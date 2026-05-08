from __future__ import annotations

import os
import re
import traceback
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

import gspread
from playwright.sync_api import Error, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright
from unidecode import unidecode

from bot_faturas_v2.google_auth import build_google_credentials
from scripts.tiflux_historico import (
    click_area_de_faturas_tab,
    ensure_directory,
    finish_login_if_needed,
    is_login_screen,
    open_edit_modal,
    save_modal,
    save_storage_state,
    sanitize_filename,
    take_screenshot,
    ticket_url,
)


GOOGLE_SCOPES = (
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
)

DEFAULT_TIFLUX_EMAIL = os.getenv("TIFLUX_EMAIL", "")
DEFAULT_TIFLUX_PASSWORD = os.getenv("TIFLUX_PASSWORD", "")
DEFAULT_TIFLUX_BASE_URL = os.getenv("TIFLUX_BASE_URL", "https://app.tiflux.com")
DEFAULT_TIFLUX_ENTITY_PATH = os.getenv("TIFLUX_ENTITY_PATH", "entities_3622")
DEFAULT_TIFLUX_SESSION_FILE = Path(os.getenv("TIFLUX_SESSION_FILE", "storage/tiflux/session.json"))
DEFAULT_TIFLUX_OUTPUT_DIR = Path(os.getenv("TIFLUX_OUTPUT_DIR", "output/tiflux"))

RESULT_HEADERS = [
    "STATUS_EXECUCAO",
    "MENSAGEM_EXECUCAO",
    "PROCESSADO_EM",
    "CAMPOS_APLICADOS",
    "EVIDENCIA",
]

BUTTON_ACTION_CELL = "S2"
BUTTON_STATUS_CELL = "T2"
BUTTON_JOB_CELL = "U2"

FIELD_OPTIONS: dict[str, list[str]] = {
    "impedimento": [
        "Problema de Login no Fatura Fácil (Embratel)",
        "Sem Acesso ao site - Cliente SAAS",
        "Fatura não disponível, aberto protocolo",
        "Não disponível – Fatura do mês não está disponível.",
        "Erro no site – Problemas de acesso ao portal. Não está relacionado ...",
        "Sem acesso ao site – Problemas com login ou senha.",
        "N/A",
        "Pendente Opedora - Validando trafego",
        "Emissão atrasada devido nova NFCOM",
        "AE (Arquivo Eletrônico) – Indisponível",
        "Pendente Prorrogação",
        "Fornecedor não disponibiliza portal",
        "Não disponível – Fatura ou CNPJ não localizados nos acessos.",
        "PDF – Indisponível",
        "Erro Upload – Portal",
        "NF – Indisponível",
        "NFPREF – Arquivo RPS não foi convertido em NF PREF",
    ],
    "tratativas_observacoes": [
        "Problemas com AE – Contatar a operadora",
        "Solicitar arquivo por e-mail",
        "Solicitar arquivos para o cliente",
        "Em Cancelamento",
        "Solicitar prorrogação",
        "Corrigir inconsistência",
        "Escalado Anatel: Dentro do Prazo",
        "Acompanhar devolutiva – Cliente",
        "Abrir protocolo junto ao fornecedor",
        "N/A",
        "Reprocessar conta",
        "Ajustar Valores",
        "Escalado Ouvidoria: Dentro do Prazo",
        "Direcionado ao desenvolvedor",
        "Validação de Cancelamento",
        "Validação de Ausência de trafego",
        "Validação de bloqueio/suspensão temporária",
        "Avaliar possível alteração de vencimento",
        "NF PREF – Acompanhar conversão para NFPREF",
    ],
    "estagio": [
        "Pendente",
        "Em andamento",
        "Concluído",
    ],
}


@dataclass(frozen=True, slots=True)
class SheetFieldSpec:
    header: str
    label: str
    kind: str


FIELD_SPECS = (
    SheetFieldSpec("historico_da_fatura", "Histórico da Fatura", "text"),
    SheetFieldSpec("impedimento", "Impedimento", "select"),
    SheetFieldSpec("tratativas_observacoes", "Tratativas/observações", "select"),
    SheetFieldSpec("fatura_assumida_data", "Fatura Assumida (Data)", "date"),
    SheetFieldSpec("bo_dt_data", "BO+DT (Data)", "date"),
    SheetFieldSpec("rps_nf_data", "RPS+NF (Data)", "date"),
    SheetFieldSpec("nf_prefeitura", "NF Prefeitura", "date"),
    SheetFieldSpec("ae_data", "AE (Data)", "date"),
    SheetFieldSpec("importacao_data", "Importação (Data)", "date"),
    SheetFieldSpec("envio_data", "Envio (Data)", "date"),
    SheetFieldSpec("concluido_data", "Concluído (Data)", "date"),
    SheetFieldSpec("estagio", "Estágio", "stage"),
)

TICKET_ALIASES = {
    "ticket",
    "numero_ticket",
    "numero_do_ticket",
    "n_ticket",
    "ticket_numero",
}


class TifluxSheetService:
    def __init__(self, *, browser_timeout_ms: int, headless: bool, output_dir: Path, session_file: Path) -> None:
        self.browser_timeout_ms = browser_timeout_ms
        self.headless = headless
        self.output_dir = Path(output_dir)
        self.session_file = Path(session_file)

    @property
    def template_headers(self) -> list[str]:
        headers = ["NUMERO_TICKET"]
        headers.extend(spec.label for spec in FIELD_SPECS)
        headers.extend(RESULT_HEADERS)
        return headers

    @property
    def template_fields(self) -> list[dict[str, Any]]:
        return [
            {
                "key": spec.header,
                "label": spec.label,
                "kind": spec.kind,
                "options": FIELD_OPTIONS.get(spec.header, []),
            }
            for spec in FIELD_SPECS
        ]

    def process_google_sheet(self, spreadsheet_id: str, worksheet_name: str) -> dict[str, Any]:
        tiflux_email = DEFAULT_TIFLUX_EMAIL.strip()
        tiflux_password = DEFAULT_TIFLUX_PASSWORD
        if not tiflux_email or not tiflux_password:
            raise ValueError("Configure TIFLUX_EMAIL e TIFLUX_PASSWORD para rodar a automacao da planilha.")

        client = self._build_google_client()
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)
        rows = worksheet.get_all_values()
        if not rows:
            raise ValueError("A aba informada esta vazia.")

        header_row = rows[0]
        normalized_headers = [canonical_header_name(item) for item in header_row]
        header_positions = {name: index for index, name in enumerate(normalized_headers) if name}
        ticket_column = self._resolve_ticket_column(header_positions)

        missing_headers = [item for item in RESULT_HEADERS if item not in header_row]
        if missing_headers:
            header_row = header_row + missing_headers
            worksheet.update("A1", [header_row])
            rows = worksheet.get_all_values()
            normalized_headers = [canonical_header_name(item) for item in rows[0]]
            header_positions = {name: index for index, name in enumerate(normalized_headers) if name}
            ticket_column = self._resolve_ticket_column(header_positions)

        parsed_rows = self._parse_rows(rows, ticket_column, header_positions)
        if not parsed_rows:
            raise ValueError("Nenhuma linha valida encontrada. Preencha ao menos o numero do ticket e um campo para atualizar.")

        self.update_button_status(
            worksheet,
            status_text=f"Processando {len(parsed_rows)} linha(s)...",
        )
        self.initialize_result_rows(worksheet, header_row, parsed_rows)

        ensure_directory(self.output_dir)
        ensure_directory(self.session_file.parent)

        summary: list[dict[str, Any]] = []
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=self.headless)
            first_url = ticket_url(DEFAULT_TIFLUX_BASE_URL, parsed_rows[0].ticket, DEFAULT_TIFLUX_ENTITY_PATH)
            context, page = self._create_tiflux_page(browser, use_saved_session=True)
            context, page = self._ensure_authenticated_page(
                browser=browser,
                context=context,
                page=page,
                first_url=first_url,
                email=tiflux_email,
                password=tiflux_password,
            )

            for item in parsed_rows:
                self.mark_row_processing(worksheet, header_row, item)
                result = self._process_sheet_row(page, item)
                summary.append(result)
                self._write_result_row(worksheet, item.row_number, header_row, result)

            context.close()
            browser.close()

        payload = {
            "ok": True,
            "spreadsheet_id": spreadsheet_id,
            "worksheet_name": worksheet_name,
            "processed": len(summary),
            "updated": sum(1 for item in summary if item["status"] == "OK"),
            "failed": sum(1 for item in summary if item["status"] in {"ERRO", "ALERTA"}),
            "results": summary,
        }
        self.update_button_status(
            worksheet,
            status_text=(
                f"Concluído: {payload['updated']} OK, "
                f"{sum(1 for item in summary if item['status'] == 'ALERTA')} alerta(s), "
                f"{sum(1 for item in summary if item['status'] == 'ERRO')} erro(s)"
            ),
        )
        return payload

    def process_ticket_batch(
        self,
        tickets: list[str],
        raw_updates: dict[str, Any],
        auth_code: str = "",
        auth_code_provider: Callable[[], str] | None = None,
        progress_callback: Callable[[list[dict[str, Any]], int, str], None] | None = None,
    ) -> dict[str, Any]:
        tiflux_email = DEFAULT_TIFLUX_EMAIL.strip()
        tiflux_password = DEFAULT_TIFLUX_PASSWORD
        if not tiflux_email or not tiflux_password:
            raise ValueError("Configure TIFLUX_EMAIL e TIFLUX_PASSWORD para rodar a automacao do TiFlux.")

        updates = self._build_direct_updates(raw_updates)
        if not updates:
            raise ValueError("Preencha ao menos um campo para atualizar no TiFlux.")

        normalized_tickets = unique_tickets(tickets)
        if not normalized_tickets:
            raise ValueError("Informe ao menos um ticket valido.")

        parsed_rows = [
            ParsedSheetRow(row_number=index + 1, ticket=ticket, updates=updates)
            for index, ticket in enumerate(normalized_tickets)
        ]

        ensure_directory(self.output_dir)
        ensure_directory(self.session_file.parent)

        summary: list[dict[str, Any]] = []
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=self.headless)
            first_url = ticket_url(DEFAULT_TIFLUX_BASE_URL, parsed_rows[0].ticket, DEFAULT_TIFLUX_ENTITY_PATH)
            context, page = self._create_tiflux_page(browser, use_saved_session=True)
            context, page = self._ensure_authenticated_page(
                browser=browser,
                context=context,
                page=page,
                first_url=first_url,
                email=tiflux_email,
                password=tiflux_password,
                auth_code=auth_code,
                auth_code_provider=auth_code_provider,
            )

            for item in parsed_rows:
                if progress_callback:
                    progress_callback(summary, len(parsed_rows), item.ticket)
                result = self._process_sheet_row(page, item, auth_code=auth_code, auth_code_provider=auth_code_provider)
                summary.append(result)
                if progress_callback:
                    progress_callback(summary, len(parsed_rows), "")

            context.close()
            browser.close()

        return {
            "ok": True,
            "processed": len(summary),
            "updated": sum(1 for item in summary if item["status"] == "OK"),
            "failed": sum(1 for item in summary if item["status"] in {"ERRO", "ALERTA"}),
            "results": summary,
            "tickets": normalized_tickets,
            "fields_applied": [update.spec.label for update in updates],
        }

    def _create_tiflux_page(self, browser, *, use_saved_session: bool):
        context_kwargs: dict[str, object] = {"viewport": {"width": 1600, "height": 1200}}
        if use_saved_session and self.session_file.exists():
            context_kwargs["storage_state"] = str(self.session_file)
        context = browser.new_context(**context_kwargs)
        page = context.new_page()
        page.set_default_timeout(self.browser_timeout_ms)
        return context, page

    def _ensure_authenticated_page(
        self,
        *,
        browser,
        context,
        page: Page,
        first_url: str,
        email: str,
        password: str,
        auth_code: str = "",
        auth_code_provider: Callable[[], str] | None = None,
    ):
        page.goto(first_url, wait_until="domcontentloaded")
        page.wait_for_timeout(2000)
        if not is_login_screen(page):
            return context, page

        try:
            finish_login_if_needed(
                page=page,
                email=email,
                password=password,
                auth_code=auth_code,
                auth_code_provider=auth_code_provider,
                headless=self.headless,
                timeout_ms=self.browser_timeout_ms,
            )
        except Exception:
            context.close()
            context, page = self._create_tiflux_page(browser, use_saved_session=False)
            page.goto(first_url, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)
            finish_login_if_needed(
                page=page,
                email=email,
                password=password,
                auth_code=auth_code,
                auth_code_provider=auth_code_provider,
                headless=self.headless,
                timeout_ms=self.browser_timeout_ms,
            )

        page.goto(first_url, wait_until="domcontentloaded")
        page.wait_for_timeout(2000)
        save_storage_state(page, self.session_file)
        return context, page

    def set_job_banner(
        self,
        spreadsheet_id: str,
        worksheet_name: str,
        *,
        job_id: str,
        status_text: str,
        reset_checkbox: bool = False,
    ) -> None:
        client = self._build_google_client()
        worksheet = client.open_by_key(spreadsheet_id).worksheet(worksheet_name)
        self.update_button_status(
            worksheet,
            status_text=status_text,
            job_id=job_id,
            reset_checkbox=reset_checkbox,
        )

    def update_button_status(
        self,
        worksheet: gspread.Worksheet,
        *,
        status_text: str,
        job_id: str | None = None,
        reset_checkbox: bool = False,
    ) -> None:
        values = [[status_text]]
        worksheet.update(range_name=BUTTON_STATUS_CELL, values=values, value_input_option="USER_ENTERED")
        if job_id is not None:
            worksheet.update(range_name=BUTTON_JOB_CELL, values=[[job_id]], value_input_option="USER_ENTERED")
        if reset_checkbox:
            worksheet.update(range_name=BUTTON_ACTION_CELL, values=[[False]], value_input_option="USER_ENTERED")

    def initialize_result_rows(
        self,
        worksheet: gspread.Worksheet,
        header_row: list[str],
        parsed_rows: list["ParsedSheetRow"],
    ) -> None:
        for item in parsed_rows:
            result = {
                "status": "NA_FILA",
                "message": "Linha identificada e aguardando processamento.",
                "processed_at": now_iso(),
                "fields_applied": ", ".join(update.spec.label for update in item.updates),
                "evidence": "",
            }
            self._write_result_row(worksheet, item.row_number, header_row, result)

    def mark_row_processing(
        self,
        worksheet: gspread.Worksheet,
        header_row: list[str],
        item: "ParsedSheetRow",
    ) -> None:
        result = {
            "status": "PROCESSANDO",
            "message": f"Processando ticket {item.ticket}...",
            "processed_at": now_iso(),
            "fields_applied": ", ".join(update.spec.label for update in item.updates),
            "evidence": "",
        }
        self._write_result_row(worksheet, item.row_number, header_row, result)

    def _build_google_client(self) -> gspread.Client:
        credentials = build_google_credentials(GOOGLE_SCOPES)
        return gspread.authorize(credentials)

    def _build_direct_updates(self, raw_updates: dict[str, Any]) -> list["FieldUpdate"]:
        updates: list[FieldUpdate] = []
        spec_by_header = {spec.header: spec for spec in FIELD_SPECS}
        for raw_key, raw_value in (raw_updates or {}).items():
            canonical_key = canonical_header_name(raw_key)
            spec = spec_by_header.get(canonical_key)
            if not spec:
                continue
            value = clean_text(raw_value)
            if not value:
                continue
            if spec.kind == "date":
                value = normalize_date_string(value)
            updates.append(FieldUpdate(spec=spec, value=value))
        return updates

    def _resolve_ticket_column(self, header_positions: dict[str, int]) -> str:
        for alias in TICKET_ALIASES:
            if alias in header_positions:
                return alias
        raise ValueError("A planilha precisa ter a primeira coluna com o numero do ticket.")

    def _parse_rows(
        self,
        rows: list[list[str]],
        ticket_column: str,
        header_positions: dict[str, int],
    ) -> list["ParsedSheetRow"]:
        parsed: list[ParsedSheetRow] = []
        spec_by_header = {spec.header: spec for spec in FIELD_SPECS}
        for row_number, raw_row in enumerate(rows[1:], start=2):
            row = list(raw_row)
            if not any(clean_text(cell) for cell in row):
                continue

            ticket = clean_text(self._cell(row, header_positions[ticket_column]))
            updates: list[FieldUpdate] = []
            for spec in FIELD_SPECS:
                column_index = header_positions.get(spec.header)
                if column_index is None:
                    continue
                value = clean_text(self._cell(row, column_index))
                if not value:
                    continue
                if spec.kind == "date":
                    value = normalize_date_string(value)
                updates.append(FieldUpdate(spec=spec_by_header[spec.header], value=value))

            if not ticket and not updates:
                continue

            parsed.append(ParsedSheetRow(row_number=row_number, ticket=ticket, updates=updates))
        return parsed

    def _process_sheet_row(
        self,
        page: Page,
        item: "ParsedSheetRow",
        *,
        auth_code: str = "",
        auth_code_provider: Callable[[], str] | None = None,
    ) -> dict[str, Any]:
        fields_applied = ", ".join(update.spec.label for update in item.updates)
        if not item.ticket:
            return {
                "row_number": item.row_number,
                "ticket": "",
                "status": "ERRO",
                "message": "Numero do ticket nao informado.",
                "fields_applied": fields_applied,
                "processed_at": now_iso(),
                "evidence": "",
            }
        if not item.updates:
            return {
                "row_number": item.row_number,
                "ticket": item.ticket,
                "status": "IGNORADO",
                "message": "Nenhum campo foi preenchido para esta linha.",
                "fields_applied": "",
                "processed_at": now_iso(),
                "evidence": "",
            }

        try:
            page.goto(ticket_url(DEFAULT_TIFLUX_BASE_URL, item.ticket, DEFAULT_TIFLUX_ENTITY_PATH), wait_until="domcontentloaded")
            page.wait_for_timeout(1800)
            if is_login_screen(page):
                finish_login_if_needed(
                    page=page,
                    email=DEFAULT_TIFLUX_EMAIL,
                    password=DEFAULT_TIFLUX_PASSWORD,
                    auth_code=auth_code,
                    auth_code_provider=auth_code_provider,
                    headless=self.headless,
                    timeout_ms=self.browser_timeout_ms,
                )
                page.goto(ticket_url(DEFAULT_TIFLUX_BASE_URL, item.ticket, DEFAULT_TIFLUX_ENTITY_PATH), wait_until="domcontentloaded")
                page.wait_for_timeout(1800)
                save_storage_state(page, self.session_file)

            modal_updates = [update for update in item.updates if update.spec.kind != "stage"]
            stage_updates = [update for update in item.updates if update.spec.kind == "stage"]
            mismatches: list[str] = []
            errors: list[str] = []

            if stage_updates:
                for update in stage_updates:
                    try:
                        set_stage_value(page, update.value)
                        actual_stage = normalize_compare_value(read_stage_value(page), "stage")
                        expected_stage = normalize_compare_value(update.value, "stage")
                        if actual_stage != expected_stage:
                            mismatches.append(update.spec.label)
                    except Exception as exc:  # noqa: BLE001
                        errors.append(f"{update.spec.label}: {exc}")

            if modal_updates:
                try:
                    click_area_de_faturas_tab(page)
                    open_edit_modal(page)
                    for update in modal_updates:
                        set_modal_field_value(page, update.spec.label, update.value, update.spec.kind)
                    save_modal(page)

                    page.wait_for_timeout(700)
                    open_edit_modal(page)
                    mismatches.extend(verify_modal_field_values(page, modal_updates))
                except Exception as exc:  # noqa: BLE001
                    errors.append(f"Area de Faturas: {exc}")

            evidence = str(take_screenshot(page, self.output_dir, item.ticket, "planilha_tiflux"))
            if errors:
                message = f"Atualizacao parcial com erro: {'; '.join(errors)}"
                status = "ALERTA"
            elif mismatches:
                message = f"Validacao parcial: {', '.join(mismatches)}"
                status = "ALERTA"
            else:
                message = "Campos salvos com sucesso."
                status = "OK"
            return {
                "row_number": item.row_number,
                "ticket": item.ticket,
                "status": status,
                "message": message,
                "fields_applied": fields_applied,
                "processed_at": now_iso(),
                "evidence": evidence,
            }
        except (PlaywrightTimeoutError, Error, RuntimeError, ValueError) as exc:
            evidence = ""
            try:
                evidence = str(take_screenshot(page, self.output_dir, item.ticket or f"linha_{item.row_number}", "planilha_tiflux_erro"))
            except Exception:  # noqa: BLE001
                evidence = ""
            return {
                "row_number": item.row_number,
                "ticket": item.ticket,
                "status": "ERRO",
                "message": str(exc),
                "fields_applied": fields_applied,
                "processed_at": now_iso(),
                "evidence": evidence,
                "traceback": traceback.format_exc(limit=3),
            }

    def _write_result_row(
        self,
        worksheet: gspread.Worksheet,
        row_number: int,
        header_row: list[str],
        result: dict[str, Any],
    ) -> None:
        header_positions = {canonical_header_name(name): index for index, name in enumerate(header_row)}
        payload = {
            "status_execucao": result.get("status", ""),
            "mensagem_execucao": result.get("message", ""),
            "processado_em": result.get("processed_at", ""),
            "campos_aplicados": result.get("fields_applied", ""),
            "evidencia": result.get("evidence", ""),
        }
        updates = []
        for header in RESULT_HEADERS:
            normalized = canonical_header_name(header)
            updates.append(
                {
                    "range": f"{column_letter(header_positions[normalized] + 1)}{row_number}",
                    "values": [[payload.get(normalized, "")]],
                }
            )
        worksheet.batch_update(updates)

    @staticmethod
    def _cell(row: list[str], index: int) -> str:
        return row[index] if index < len(row) else ""


@dataclass(frozen=True, slots=True)
class FieldUpdate:
    spec: SheetFieldSpec
    value: str


@dataclass(frozen=True, slots=True)
class ParsedSheetRow:
    row_number: int
    ticket: str
    updates: list[FieldUpdate]


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def unique_tickets(values: list[str]) -> list[str]:
    tickets: list[str] = []
    seen: set[str] = set()
    for value in values:
        ticket = clean_text(value)
        if not ticket or ticket in seen:
            continue
        seen.add(ticket)
        tickets.append(ticket)
    return tickets


def normalize_header(value: Any) -> str:
    text = unidecode(clean_text(value)).lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return re.sub(r"_+", "_", text).strip("_")


def canonical_header_name(value: Any) -> str:
    normalized = normalize_header(value)
    if not normalized:
        return ""

    if "ticket" in normalized:
        return "numero_ticket"
    if ("histor" in normalized or normalized.startswith("hist")) and "fatura" in normalized:
        return "historico_da_fatura"
    if "imped" in normalized:
        return "impedimento"
    if "tratativ" in normalized or "observ" in normalized:
        return "tratativas_observacoes"
    if "fatura" in normalized and "assum" in normalized:
        return "fatura_assumida_data"
    if "bo" in normalized and "dt" in normalized:
        return "bo_dt_data"
    if "rps" in normalized and "nf" in normalized:
        return "rps_nf_data"
    if "nf" in normalized and "prefeitura" in normalized:
        return "nf_prefeitura"
    if normalized.startswith("ae") or normalized.endswith("_ae") or "ae_data" in normalized:
        return "ae_data"
    if "import" in normalized:
        return "importacao_data"
    if "envio" in normalized:
        return "envio_data"
    if "conclu" in normalized:
        return "concluido_data"
    if "estag" in normalized or normalized.startswith("est"):
        return "estagio"
    if "status" in normalized and "exec" in normalized:
        return "status_execucao"
    if "mensagem" in normalized and "exec" in normalized:
        return "mensagem_execucao"
    if "processado" in normalized:
        return "processado_em"
    if "campos" in normalized and "aplic" in normalized:
        return "campos_aplicados"
    if "evidenc" in normalized:
        return "evidencia"
    return normalized


def normalize_date_string(value: str) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", text):
        return text
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        year, month, day = text.split("-")
        return f"{day}/{month}/{year}"
    for fmt in ("%d-%m-%Y", "%d.%m.%Y", "%Y/%m/%d"):
        try:
            parsed = datetime.strptime(text, fmt)
            return parsed.strftime("%d/%m/%Y")
        except ValueError:
            continue
    return text


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def column_letter(index: int) -> str:
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def xpath_literal(value: str) -> str:
    if "'" not in value:
        return f"'{value}'"
    if '"' not in value:
        return f'"{value}"'
    parts = value.split("'")
    return "concat(" + ", \"'\", ".join(f"'{part}'" for part in parts) + ")"


def locate_form_item(page: Page, label: str):
    modal = page.locator(".ant-modal-content").first
    label_text = xpath_literal(label)
    strategies = (
        lambda: modal.locator(
            f"xpath=.//label[normalize-space(.)={label_text}]/ancestor::*[contains(@class,'ant-form-item')][1]"
        ).first,
        lambda: modal.locator(
            f"xpath=.//*[contains(@class,'ant-form-item')][.//label[contains(normalize-space(.), {label_text})]]"
        ).first,
    )
    last_error: Exception | None = None
    for strategy in strategies:
        try:
            field = strategy()
            field.wait_for(state="visible", timeout=4000)
            return field
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    raise RuntimeError(f"Nao foi possivel localizar o campo '{label}' no modal.") from last_error


def set_modal_field_value(page: Page, label: str, value: str, kind: str) -> None:
    field = locate_form_item(page, label)
    if kind in {"text", "date"}:
        fill_text_like_control(page, field, value)
        return
    if kind == "select":
        fill_select_control(page, field, value)
        return
    raise ValueError(f"Tipo de campo nao suportado: {kind}")


def fill_text_like_control(page: Page, field, value: str) -> None:
    strategies = (
        lambda: field.locator("textarea:visible").first,
        lambda: field.locator("input:not([type='hidden']):visible").first,
    )
    last_error: Exception | None = None
    for strategy in strategies:
        try:
            control = strategy()
            control.wait_for(state="visible", timeout=3000)
            control.click()
            page.keyboard.press("Control+A")
            control.fill(value)
            page.wait_for_timeout(250)
            return
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    raise RuntimeError("Nao foi possivel preencher o campo de texto/data.") from last_error


def fill_select_control(page: Page, field, value: str) -> None:
    strategies = (
        lambda: field.locator(".ant-select-selector:visible").first,
        lambda: field.locator(".ant-select:visible").first,
        lambda: field.locator("[role='combobox']:visible").first,
    )
    last_error: Exception | None = None
    for strategy in strategies:
        try:
            control = strategy()
            control.wait_for(state="visible", timeout=3000)
            control.click()
            page.wait_for_timeout(350)

            search_inputs = page.locator(".ant-select-dropdown:visible input:visible")
            if search_inputs.count():
                search = search_inputs.last
                search.click()
                page.keyboard.press("Control+A")
                search.fill(value)
            else:
                page.keyboard.type(value, delay=30)

            page.wait_for_timeout(400)
            page.keyboard.press("Enter")
            page.wait_for_timeout(400)
            return
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    raise RuntimeError(f"Nao foi possivel selecionar o valor '{value}'.") from last_error


def read_modal_field_value(page: Page, label: str, kind: str) -> str:
    field = locate_form_item(page, label)
    if kind in {"text", "date"}:
        controls = (
            field.locator("textarea:visible").first,
            field.locator("input:not([type='hidden']):visible").first,
        )
        for control in controls:
            try:
                control.wait_for(state="visible", timeout=1500)
                return clean_text(control.input_value())
            except Exception:  # noqa: BLE001
                continue
        return ""
    selection = field.locator(".ant-select-selection-item:visible").first
    try:
        selection.wait_for(state="visible", timeout=1500)
        return clean_text(selection.inner_text())
    except Exception:  # noqa: BLE001
        pass
    input_control = field.locator("input:visible").first
    try:
        input_control.wait_for(state="visible", timeout=1500)
        return clean_text(input_control.input_value())
    except Exception:  # noqa: BLE001
        return ""


def verify_modal_field_values(page: Page, updates: list[FieldUpdate]) -> list[str]:
    mismatches: list[str] = []
    for update in updates:
        actual = normalize_compare_value(read_modal_field_value(page, update.spec.label, update.spec.kind), update.spec.kind)
        expected = normalize_compare_value(update.value, update.spec.kind)
        if actual != expected:
            mismatches.append(update.spec.label)
    return mismatches


def set_stage_value(page: Page, value: str) -> None:
    target_value = canonical_stage_value(value)
    try:
        page.get_by_text("Informações gerais", exact=True).click()
        page.wait_for_timeout(500)
    except Exception:  # noqa: BLE001
        pass

    selectors = (
        "xpath=//*[normalize-space()='Estágio']/following::*[@role='combobox'][1]",
        "xpath=//*[normalize-space()='Estágio']/following::*[contains(@class,'ant-select-selector')][1]",
        "xpath=//*[normalize-space()='Estágio']/following::input[1]",
        "xpath=//*[contains(normalize-space(),'Est') and contains(normalize-space(),'gio')]/following::*[@role='combobox'][1]",
        "xpath=//*[contains(normalize-space(),'Est') and contains(normalize-space(),'gio')]/following::*[contains(@class,'ant-select-selector')][1]",
    )
    last_error: Exception | None = None
    for selector in selectors:
        try:
            control = page.locator(selector).first
            control.wait_for(state="visible", timeout=5000)
            control.click()
            page.wait_for_timeout(300)
            if click_stage_option(page, target_value):
                page.wait_for_timeout(900)
                return
            page.keyboard.press("Control+A")
            page.keyboard.type(unidecode(target_value), delay=30)
            page.wait_for_timeout(400)
            if click_stage_option(page, target_value):
                page.wait_for_timeout(900)
                return
            page.keyboard.press("Enter")
            page.wait_for_timeout(900)
            return
        except Exception as exc:  # noqa: BLE001
            last_error = exc
    raise RuntimeError(f"Nao foi possivel definir o estágio '{value}'.") from last_error


def canonical_stage_value(value: str) -> str:
    text = unidecode(clean_text(value)).casefold()
    if "andamento" in text:
        return "Em andamento"
    if "conclu" in text:
        return "Conclu\u00eddo"
    if "pendente" in text:
        return "Pendente"
    return clean_text(value)


def click_stage_option(page: Page, value: str) -> bool:
    target_key = normalize_stage_key(value)
    option_locators = (
        page.get_by_role("option"),
        page.locator(".ant-select-item-option:visible"),
        page.locator("[title]:visible"),
    )
    for locator in option_locators:
        try:
            count = min(locator.count(), 30)
        except Exception:  # noqa: BLE001
            continue
        for index in range(count):
            option = locator.nth(index)
            try:
                option_text = clean_text(option.inner_text() or option.get_attribute("title") or "")
                if normalize_stage_key(option_text) == target_key:
                    option.click()
                    return True
            except Exception:  # noqa: BLE001
                continue
    return False


def read_stage_value(page: Page) -> str:
    selectors = (
        "xpath=//*[normalize-space()='Estágio']/following::*[@role='combobox'][1]",
        "xpath=//*[normalize-space()='Estágio']/following::*[contains(@class,'ant-select-selector')][1]",
        "xpath=//*[normalize-space()='Estágio']/following::input[1]",
        "xpath=//*[contains(normalize-space(),'Est') and contains(normalize-space(),'gio')]/following::*[@role='combobox'][1]",
        "xpath=//*[contains(normalize-space(),'Est') and contains(normalize-space(),'gio')]/following::*[contains(@class,'ant-select-selector')][1]",
    )
    for selector in selectors:
        try:
            control = page.locator(selector).first
            control.wait_for(state="visible", timeout=2000)
            text = clean_text(control.inner_text())
            if text:
                return text
            value = clean_text(control.input_value())
            if value:
                return value
        except Exception:  # noqa: BLE001
            continue
    return ""


def normalize_compare_value(value: str, kind: str) -> str:
    text = clean_text(value)
    if kind == "date":
        return normalize_date_string(text)
    if kind == "stage":
        return normalize_stage_key(text)
    return text.casefold()


def normalize_stage_key(value: str) -> str:
    text = unidecode(clean_text(value)).casefold()
    if "andamento" in text:
        return "em_andamento"
    if "conclu" in text:
        return "concluido"
    if "pendente" in text:
        return "pendente"
    return re.sub(r"[^a-z0-9]+", "_", text).strip("_")
