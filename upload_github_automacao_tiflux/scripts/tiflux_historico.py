from __future__ import annotations

import argparse
import codecs
import getpass
import os
import re
import sqlite3
import sys
import time
from pathlib import Path

from playwright.sync_api import Error, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SESSION_FILE = PROJECT_ROOT / "storage" / "tiflux" / "session.json"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output" / "tiflux"
DEFAULT_BASE_URL = "https://app.tiflux.com"
DEFAULT_ENTITY_PATH = "entities_3622"
EMPTY_HISTORY_VALUES = {"", "Sem dados"}
OUTLOOK_NOTIFICATIONS_DB = Path.home() / "AppData" / "Local" / "Microsoft" / "Windows" / "Notifications" / "wpndatabase.db"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Atualiza o campo 'Historico da Fatura' de um ticket no TiFlux."
    )
    parser.add_argument("ticket", help="Numero do ticket no TiFlux.")
    parser.add_argument("historico", help="Texto a ser salvo em 'Historico da Fatura'.")
    parser.add_argument("--email", default=os.getenv("TIFLUX_EMAIL"), help="Login do TiFlux.")
    parser.add_argument(
        "--password",
        default=os.getenv("TIFLUX_PASSWORD"),
        help="Senha do TiFlux. Se omitida, sera solicitada quando necessario.",
    )
    parser.add_argument(
        "--auth-code",
        default=os.getenv("TIFLUX_AUTH_CODE"),
        help="Codigo de autenticacao por e-mail, quando a sessao expirar.",
    )
    parser.add_argument(
        "--base-url",
        default=os.getenv("TIFLUX_BASE_URL", DEFAULT_BASE_URL),
        help="URL base do TiFlux.",
    )
    parser.add_argument(
        "--entity-path",
        default=os.getenv("TIFLUX_ENTITY_PATH", DEFAULT_ENTITY_PATH),
        help="Trecho final da URL do ticket. Ex.: entities_3622",
    )
    parser.add_argument(
        "--session-file",
        default=str(DEFAULT_SESSION_FILE),
        help="Arquivo JSON para reaproveitar a sessao autenticada.",
    )
    parser.add_argument(
        "--output-dir",
        default=str(DEFAULT_OUTPUT_DIR),
        help="Diretorio para screenshots e evidencias.",
    )
    parser.add_argument(
        "--timeout-ms",
        type=int,
        default=int(os.getenv("TIFLUX_TIMEOUT_MS", "45000")),
        help="Timeout padrao do navegador em milissegundos.",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Forca execucao sem abrir a janela do navegador.",
    )
    parser.add_argument(
        "--show-browser",
        action="store_true",
        help="Forca execucao com a janela do navegador visivel.",
    )
    return parser.parse_args()


def ensure_directory(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def sanitize_filename(value: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_.-]+", "_", value).strip("_") or "arquivo"


def resolve_historico_text(value: str) -> str:
    if "\\u" in value or "\\x" in value:
        return codecs.decode(value, "unicode_escape")
    return value


def ticket_url(base_url: str, ticket: str, entity_path: str) -> str:
    return f"{base_url.rstrip('/')}/v/tickets/{ticket}/{entity_path.strip('/')}"


def is_login_screen(page: Page) -> bool:
    return page.locator("text=Acessar conta").first.is_visible()


def is_auth_code_screen(page: Page) -> bool:
    return page.locator("text=C\u00f3digo de autentica\u00e7\u00e3o enviado por e-mail").first.is_visible()


def click_submit_button(page: Page) -> None:
    strategies = (
        lambda: page.locator("button[type='submit']:visible").click(),
        lambda: page.get_by_role("button", name="Acessar conta").last.click(),
        lambda: page.get_by_role("button", name="Acessar conta").click(),
    )
    last_error: Exception | None = None
    for strategy in strategies:
        try:
            strategy()
            return
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    raise RuntimeError("Nao foi possivel acionar o botao de envio da autenticacao.") from last_error


def get_account_blocked_message(page: Page) -> str:
    body_text = page.locator("body").inner_text()
    pattern = re.compile(r"Sua conta foi bloqueada.*", re.IGNORECASE)
    match = pattern.search(body_text)
    return match.group(0).strip() if match else ""


def find_latest_tiflux_auth_code() -> str:
    code, _arrival_time = find_latest_tiflux_auth_code_info()
    return code


def find_latest_tiflux_auth_code_info() -> tuple[str, int]:
    if not OUTLOOK_NOTIFICATIONS_DB.exists():
        return "", 0

    query = (
        "SELECT ArrivalTime, Payload "
        "FROM Notification "
        "WHERE HandlerId = 199 AND Type = 'toast' "
        "ORDER BY ArrivalTime DESC LIMIT 20"
    )
    try:
        connection = sqlite3.connect(f"file:{OUTLOOK_NOTIFICATIONS_DB}?mode=ro", uri=True)
        cursor = connection.cursor()
        rows = cursor.execute(query).fetchall()
        connection.close()
    except sqlite3.Error:
        return "", 0

    for arrival_time, payload in rows:
        text = payload.decode("utf-8", errors="ignore") if isinstance(payload, (bytes, bytearray)) else str(payload)
        if "Tiflux" not in text and "tiflux" not in text:
            continue
        matches = re.findall(r"\b\d{6}\b", text)
        if matches:
            return matches[-1], int(arrival_time or 0)
    return "", 0


def wait_for_fresh_tiflux_auth_code(previous_arrival_time: int = 0, timeout_seconds: int = 45) -> str:
    deadline = time.time() + timeout_seconds
    latest_code = ""
    while time.time() < deadline:
        code, arrival_time = find_latest_tiflux_auth_code_info()
        if code:
            latest_code = code
        if code and arrival_time > previous_arrival_time:
            return code
        time.sleep(2)
    return latest_code


def prompt_password_if_needed(password: str | None) -> str:
    if password:
        return password
    return getpass.getpass("Senha do TiFlux: ")


def fill_auth_code(page: Page, auth_code: str) -> bool:
    code = re.sub(r"\D+", "", auth_code or "")
    if not code:
        return False

    single_inputs = page.locator("input[maxlength='1']:visible")
    if single_inputs.count() >= len(code):
        for index, digit in enumerate(code):
            single_inputs.nth(index).fill(digit)
        return True

    indexed_inputs = page.locator("input[data-id]:visible")
    if indexed_inputs.count() >= len(code):
        for index, digit in enumerate(code):
            indexed_inputs.nth(index).fill(digit)
        return True

    visible_inputs = page.locator("input:visible")
    if visible_inputs.count() >= len(code):
        for index, digit in enumerate(code):
            candidate = visible_inputs.nth(index)
            try:
                candidate.fill(digit)
            except Exception:  # noqa: BLE001
                try:
                    candidate.click()
                    page.keyboard.press("Control+A")
                    page.keyboard.type(digit)
                except Exception:  # noqa: BLE001
                    continue
        return True

    generic_inputs = page.locator("input:not([type='hidden']):visible")
    if generic_inputs.count() == 0:
        return False

    generic_inputs.first.fill(code)
    return True


def finish_login_if_needed(
    page: Page,
    email: str | None,
    password: str | None,
    auth_code: str | None,
    headless: bool,
    timeout_ms: int,
) -> None:
    if not is_login_screen(page):
        return

    if not email:
        raise RuntimeError("Sessao expirada e nenhum e-mail do TiFlux foi informado.")

    _previous_auth_code, previous_auth_arrival_time = find_latest_tiflux_auth_code_info()
    page.get_by_label("E-mail").fill(email)
    page.get_by_label("Senha").fill(prompt_password_if_needed(password))
    click_submit_button(page)
    page.wait_for_timeout(3000)

    blocked_message = get_account_blocked_message(page)
    if blocked_message:
        raise RuntimeError(blocked_message)

    if not is_auth_code_screen(page):
        return

    resolved_auth_code = auth_code or wait_for_fresh_tiflux_auth_code(previous_auth_arrival_time)
    if resolved_auth_code and fill_auth_code(page, resolved_auth_code):
        click_submit_button(page)
        page.wait_for_timeout(3000)
        blocked_message = get_account_blocked_message(page)
        if blocked_message:
            raise RuntimeError(blocked_message)
        if is_auth_code_screen(page):
            raise RuntimeError("O codigo de autenticacao nao foi aceito. Gere um novo codigo e tente novamente.")
        return

    print(
        "\nAutenticacao adicional necessaria no TiFlux."
        "\nSe o navegador estiver aberto, conclua o codigo e pressione ENTER aqui."
        "\nSe preferir rodar totalmente em background, execute novamente com --auth-code.\n"
    )
    if headless:
        typed_code = input("Codigo de autenticacao: ").strip()
        if not fill_auth_code(page, typed_code):
            raise RuntimeError("Nao foi possivel localizar os campos do codigo de autenticacao.")
        click_submit_button(page)
        return

    wait_deadline = time.time() + max(timeout_ms / 1000, 180)
    print("Aguardando conclusao da autenticacao manual no navegador...")
    while time.time() < wait_deadline:
        page.wait_for_timeout(1000)
        try:
            if not is_auth_code_screen(page):
                return
        except Exception:  # noqa: BLE001
            return

    raise RuntimeError("A tela de codigo ainda esta aberta. Revise a autenticacao e tente novamente.")


def save_storage_state(page: Page, session_file: Path) -> None:
    ensure_directory(session_file.parent)
    page.context.storage_state(path=str(session_file))


def click_area_de_faturas_tab(page: Page) -> None:
    strategies = (
        lambda: page.get_by_role("tab", name=re.compile(r"\u00c1rea de Faturas", re.I)).click(),
        lambda: page.get_by_text("\u00c1rea de Faturas", exact=True).click(),
        lambda: page.locator("text=\u00c1rea de Faturas").first.click(),
    )
    last_error: Exception | None = None
    for strategy in strategies:
        try:
            strategy()
            page.wait_for_timeout(800)
            return
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    raise RuntimeError("Nao foi possivel abrir a aba 'Area de Faturas'.") from last_error


def open_edit_modal(page: Page) -> None:
    if page.locator(".ant-modal-content").first.is_visible():
        return

    strategies = (
        lambda: page.get_by_role("button", name=re.compile(r"editar", re.I)).click(),
        lambda: page.locator("button[title*='Editar']").first.click(),
        lambda: page.locator("[aria-label*='Editar']").first.click(),
        lambda: page.locator("div.ant-card-head-wrapper div.ant-card-extra .cursor-pointer").first.click(),
        lambda: page.locator("div.ant-card-head-wrapper div.ant-card-extra span[role='img']").first.click(),
        lambda: page.locator("div, section").filter(has=page.get_by_text("\u00c1rea de Faturas", exact=True)).locator("button").first.click(),
    )
    last_error: Exception | None = None
    for strategy in strategies:
        try:
            strategy()
            page.locator(".ant-modal-content").first.wait_for(state="visible", timeout=5000)
            return
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    raise RuntimeError("Nao foi possivel abrir o modal 'Editar Area de Faturas'.") from last_error


def read_historico_value(page: Page) -> str:
    candidates = (
        page.locator("div.ant-card-small.entity").filter(has=page.get_by_text("Hist\u00f3rico da Fatura")).first,
        page.locator("xpath=//*[contains(normalize-space(.), 'Histórico da Fatura')]").first,
    )
    last_error: Exception | None = None
    for candidate in candidates:
        try:
            candidate.wait_for(state="visible", timeout=5000)
            text = candidate.inner_text().strip()
            lines = [line.strip() for line in text.splitlines() if line.strip()]
            if not lines:
                return ""
            if lines[0] == "Hist\u00f3rico da Fatura":
                return "\n".join(lines[1:]).strip()
            return "\n".join(lines).strip()
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    raise RuntimeError("Nao foi possivel ler o campo 'Historico da Fatura'.") from last_error


def fill_historico(page: Page, historico: str) -> None:
    modal = page.locator(".ant-modal-content").first

    field_strategies = (
        lambda: modal.get_by_label("Hist\u00f3rico da Fatura"),
        lambda: modal.locator("label:has-text('Hist\u00f3rico da Fatura')").locator(
            "xpath=following::*[(self::input or self::textarea) and not(@type='hidden')][1]"
        ),
        lambda: modal.locator("input:visible, textarea:visible").first,
    )

    last_error: Exception | None = None
    for strategy in field_strategies:
        try:
            field = strategy()
            field.wait_for(state="visible", timeout=5000)
            field.fill(historico)
            page.wait_for_timeout(500)
            return
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    raise RuntimeError("Nao foi possivel localizar o campo 'Historico da Fatura'.") from last_error


def save_modal(page: Page) -> None:
    page.get_by_role("button", name="Salvar").click()
    page.wait_for_timeout(1500)


def take_screenshot(page: Page, output_dir: Path, ticket: str, suffix: str) -> Path:
    ensure_directory(output_dir)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    file_name = f"{sanitize_filename(ticket)}_{suffix}_{timestamp}.png"
    destination = output_dir / file_name
    page.screenshot(path=str(destination), full_page=True)
    return destination


def main() -> int:
    args = parse_args()
    args.historico = resolve_historico_text(args.historico)

    session_file = Path(args.session_file).resolve()
    output_dir = Path(args.output_dir).resolve()
    ensure_directory(output_dir)

    requested_headless = args.headless
    if args.show_browser:
        requested_headless = False
    elif not session_file.exists():
        requested_headless = False

    url = ticket_url(args.base_url, args.ticket, args.entity_path)

    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=requested_headless)
            context_kwargs: dict[str, object] = {
                "viewport": {"width": 1600, "height": 1200},
            }
            if session_file.exists():
                context_kwargs["storage_state"] = str(session_file)

            context = browser.new_context(**context_kwargs)
            page = context.new_page()
            page.set_default_timeout(args.timeout_ms)

            page.goto(url, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)

            finish_login_if_needed(
                page=page,
                email=args.email,
                password=args.password,
                auth_code=args.auth_code,
                headless=requested_headless,
                timeout_ms=args.timeout_ms,
            )

            page.goto(url, wait_until="domcontentloaded")
            page.wait_for_timeout(2500)
            save_storage_state(page, session_file)

            click_area_de_faturas_tab(page)
            open_edit_modal(page)
            fill_historico(page, args.historico)
            save_modal(page)

            success_shot = take_screenshot(page, output_dir, args.ticket, "historico_salvo")
            context.close()
            browser.close()

        print(f"Historico atualizado com sucesso no ticket {args.ticket}.")
        print(f"Screenshot: {success_shot}")
        print(f"Sessao salva em: {session_file}")
        return 0
    except (PlaywrightTimeoutError, Error, RuntimeError) as exc:
        print(f"Falha na automacao do TiFlux: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
