from __future__ import annotations

import re
from pathlib import Path
from urllib.parse import urlparse

from playwright.async_api import Browser, BrowserContext, Error, Page, TimeoutError as PlaywrightTimeoutError

from ..config import BotFaturasSettings
from ..constants import DownloadStatus, ErrorCode, LineStatus
from ..errors import BotLineError
from ..models import DownloadArtifact, LineProcessingResult
from ..normalization import clean_text, now_local, safe_filename
from ..storage import StorageManager


class BaseConnector:
    connector_name = "BASE"
    invoice_nav_texts = ("Faturas", "Financeiro", "Segunda via", "Boletos", "Conta", "Downloads")
    pdf_keywords = ("pdf", "fatura", "boleto", "segunda via")
    ae_keywords = ("ae", "csv", "txt", "zip", "arquivo eletronico")
    login_button_texts = ("Entrar", "Acessar", "Login", "Continuar", "Prosseguir")

    def __init__(self, settings: BotFaturasSettings, storage: StorageManager) -> None:
        self.settings = settings
        self.storage = storage

    async def process(self, browser: Browser, line: dict[str, object]) -> LineProcessingResult:
        context = await browser.new_context(
            accept_downloads=True,
            ignore_https_errors=True,
            locale="pt-BR",
            viewport={"width": 1440, "height": 960},
        )
        page = await context.new_page()
        await page.set_viewport_size({"width": 1440, "height": 960})
        page.set_default_timeout(self.settings.browser_timeout_ms)
        log_path = ""

        try:
            await self._log(line, "inicio", f"Conector {self.connector_name} iniciado.")
            await page.goto(str(line.get("link_portal") or ""), wait_until="domcontentloaded")
            await self._log(line, "portal", f"Portal aberto em {page.url}")
            await self.login(page, line)
            await self.selecionar_empresa_ou_conta(page, line)
            await self.localizar_fatura(page, line)

            downloads: list[DownloadArtifact] = []
            pdf_artifact = None
            ae_artifact = None
            pdf_status = DownloadStatus.NAO_SOLICITADO
            ae_status = DownloadStatus.NAO_SOLICITADO

            if bool(line.get("baixar_pdf")):
                pdf_artifact = await self.baixar_pdf(page, context, line)
                pdf_status = DownloadStatus.BAIXADO if pdf_artifact else DownloadStatus.NAO_ENCONTRADO
                if pdf_artifact:
                    downloads.append(pdf_artifact)

            if bool(line.get("baixar_ae")):
                ae_artifact = await self.baixar_arquivo_eletronico(page, context, line)
                ae_status = DownloadStatus.BAIXADO if ae_artifact else DownloadStatus.NAO_ENCONTRADO
                if ae_artifact:
                    downloads.append(ae_artifact)

            status_final, error_code, error_description = self._summarize_outcome(
                line=line,
                pdf_artifact=pdf_artifact,
                ae_artifact=ae_artifact,
            )
            log_path = await self._log(
                line,
                "finalizacao",
                error_description or "Linha processada com retorno positivo.",
            )
            return LineProcessingResult(
                status_final=status_final,
                pdf_status=pdf_status,
                ae_status=ae_status,
                erro_codigo=error_code,
                erro_descricao=error_description,
                observacao_execucao=error_description or "Arquivos coletados com sucesso.",
                downloads=downloads,
                log_execucao_path=log_path,
                processado_em=now_local(self.settings.timezone_name),
            )
        except BotLineError as exc:
            screenshot_path, html_path = await self.coletar_evidencias_de_erro(page, line, exc.code)
            log_path = await self._log(line, "erro", f"{exc.code}: {exc.description}", exc.technical_details)
            return LineProcessingResult(
                status_final=LineStatus.ERRO_PROCESSAMENTO,
                pdf_status=DownloadStatus.FALHA if bool(line.get("baixar_pdf")) else DownloadStatus.NAO_SOLICITADO,
                ae_status=DownloadStatus.FALHA if bool(line.get("baixar_ae")) else DownloadStatus.NAO_SOLICITADO,
                erro_codigo=exc.code,
                erro_descricao=exc.description,
                observacao_execucao=exc.description,
                screenshot_erro_path=screenshot_path,
                html_erro_path=html_path,
                log_execucao_path=log_path,
                processado_em=now_local(self.settings.timezone_name),
            )
        except PlaywrightTimeoutError as exc:
            screenshot_path, html_path = await self.coletar_evidencias_de_erro(page, line, ErrorCode.ERRO_TIMEOUT)
            log_path = await self._log(line, "timeout", "Timeout durante a automacao.", str(exc))
            return LineProcessingResult(
                status_final=LineStatus.ERRO_PROCESSAMENTO,
                pdf_status=DownloadStatus.FALHA if bool(line.get("baixar_pdf")) else DownloadStatus.NAO_SOLICITADO,
                ae_status=DownloadStatus.FALHA if bool(line.get("baixar_ae")) else DownloadStatus.NAO_SOLICITADO,
                erro_codigo=ErrorCode.ERRO_TIMEOUT,
                erro_descricao="O portal demorou alem do limite configurado para responder.",
                observacao_execucao="Timeout durante a automacao.",
                screenshot_erro_path=screenshot_path,
                html_erro_path=html_path,
                log_execucao_path=log_path,
                processado_em=now_local(self.settings.timezone_name),
            )
        except Exception as exc:  # noqa: BLE001
            screenshot_path, html_path = await self.coletar_evidencias_de_erro(page, line, ErrorCode.ERRO_DESCONHECIDO)
            log_path = await self._log(line, "erro_desconhecido", "Erro nao tratado na automacao.", str(exc))
            return LineProcessingResult(
                status_final=LineStatus.ERRO_PROCESSAMENTO,
                pdf_status=DownloadStatus.FALHA if bool(line.get("baixar_pdf")) else DownloadStatus.NAO_SOLICITADO,
                ae_status=DownloadStatus.FALHA if bool(line.get("baixar_ae")) else DownloadStatus.NAO_SOLICITADO,
                erro_codigo=ErrorCode.ERRO_DESCONHECIDO,
                erro_descricao=f"Falha nao tratada pela automacao: {clean_text(exc)}",
                observacao_execucao="Falha nao tratada pela automacao.",
                screenshot_erro_path=screenshot_path,
                html_erro_path=html_path,
                log_execucao_path=log_path,
                processado_em=now_local(self.settings.timezone_name),
            )
        finally:
            await page.close()
            await context.close()

    async def login(self, page: Page, line: dict[str, object]) -> None:
        user_input = await self._first_visible(
            page,
            (
                'input[autocomplete="username"]',
                'input[type="email"]',
                'input[name*="user" i]',
                'input[name*="login" i]',
                'input[id*="user" i]',
                'input[id*="login" i]',
                'input[type="text"]',
            ),
        )
        password_input = await self._first_visible(
            page,
            (
                'input[type="password"]',
                'input[name*="senha" i]',
                'input[id*="senha" i]',
            ),
        )
        if not user_input and not password_input:
            await self._log(line, "login", "Formulario de login nao encontrado; fluxo vai seguir.")
            return

        if user_input:
            await user_input.fill(str(line.get("usuario_login") or ""))
        if password_input:
            await password_input.fill(str(line.get("senha") or ""))

        clicked = await self._click_by_text(page, self.login_button_texts)
        if not clicked and password_input:
            await password_input.press("Enter")
        await page.wait_for_load_state("domcontentloaded")
        login_error = await self._extract_login_error(page)
        if login_error:
            raise login_error
        await self._log(line, "login", "Tentativa de login executada.")

    async def selecionar_empresa_ou_conta(self, page: Page, line: dict[str, object]) -> None:
        identifier = next(
            (
                candidate
                for candidate in (
                    clean_text(line.get("codigo_cliente")),
                    clean_text(line.get("conta")),
                    clean_text(line.get("titulo_conta")),
                    clean_text(line.get("numero_ticket")),
                )
                if candidate
            ),
            "",
        )
        if not identifier:
            return

        selector = await self._first_visible(
            page,
            (
                'input[placeholder*="cliente" i]',
                'input[placeholder*="conta" i]',
                'input[placeholder*="titulo" i]',
                'input[name*="cliente" i]',
                'input[name*="conta" i]',
                'input[type="search"]',
            ),
        )
        if not selector:
            return
        await selector.fill(identifier)
        await selector.press("Enter")
        await page.wait_for_timeout(800)
        await self._log(line, "selecionar_conta", f"Busca por identificador executada com {identifier}.")

    async def localizar_fatura(self, page: Page, line: dict[str, object]) -> None:
        await self._click_by_text(page, self.invoice_nav_texts, optional=True)
        await page.wait_for_timeout(1000)
        target_terms = [
            clean_text(line.get("mes_referencia")),
            clean_text(line.get("vencimento")),
            clean_text(line.get("conta")),
            clean_text(line.get("codigo_cliente")),
            clean_text(line.get("titulo_conta")),
            clean_text(line.get("nome_arquivo_esperado")),
        ]
        page_content = clean_text(await page.content()).lower()
        lowered_terms = [term.lower() for term in target_terms if term]
        if lowered_terms and not any(term in page_content for term in lowered_terms):
            search_input = await self._first_visible(
                page,
                (
                    'input[placeholder*="buscar" i]',
                    'input[placeholder*="pesquisar" i]',
                    'input[type="search"]',
                ),
            )
            if search_input:
                await search_input.fill(lowered_terms[0])
                await search_input.press("Enter")
                await page.wait_for_timeout(1000)
                page_content = clean_text(await page.content()).lower()
        if not any(keyword in page_content for keyword in self.pdf_keywords + self.ae_keywords + self.invoice_nav_texts):
            raise BotLineError(
                ErrorCode.ERRO_FATURA_NAO_ENCONTRADA,
                "A area de faturas foi aberta, mas a fatura desejada nao apareceu com os filtros desta linha.",
            )
        await self._log(line, "localizar_fatura", "Tela de fatura localizada ou indiciada no portal.")

    async def baixar_pdf(self, page: Page, context: BrowserContext, line: dict[str, object]) -> DownloadArtifact | None:
        artifact = await self._download_using_candidates(
            page=page,
            context=context,
            line=line,
            kind="pdf",
            extensions={"pdf"},
            keywords=self.pdf_keywords,
        )
        await self._log(
            line,
            "baixar_pdf",
            "PDF encontrado e salvo." if artifact else "PDF nao encontrado nesta etapa.",
        )
        return artifact

    async def baixar_arquivo_eletronico(
        self,
        page: Page,
        context: BrowserContext,
        line: dict[str, object],
    ) -> DownloadArtifact | None:
        allowed = {item.lower() for item in (line.get("formatos_aceitos_ae") or [])}
        artifact = await self._download_using_candidates(
            page=page,
            context=context,
            line=line,
            kind="ae",
            extensions=allowed or {"ae", "txt", "zip", "csv"},
            keywords=self.ae_keywords,
        )
        await self._log(
            line,
            "baixar_ae",
            "Arquivo eletronico encontrado e salvo." if artifact else "Arquivo eletronico nao encontrado nesta etapa.",
        )
        return artifact

    async def coletar_evidencias_de_erro(
        self,
        page: Page,
        line: dict[str, object],
        error_code: ErrorCode,
    ) -> tuple[str, str]:
        if not self.settings.screenshot_on_error:
            return "", ""
        screenshot_path = self.storage.line_root(str(line.get("lote_id")), str(line.get("id"))) / f"{error_code}.png"
        html_path = self.storage.line_root(str(line.get("lote_id")), str(line.get("id"))) / f"{error_code}.html"
        try:
            await page.screenshot(path=str(screenshot_path), full_page=True)
            html_path.write_text(await page.content(), encoding="utf-8")
        except Exception:  # noqa: BLE001
            return "", ""
        return str(screenshot_path), str(html_path)

    async def _download_using_candidates(
        self,
        *,
        page: Page,
        context: BrowserContext,
        line: dict[str, object],
        kind: str,
        extensions: set[str],
        keywords: tuple[str, ...],
    ) -> DownloadArtifact | None:
        anchors = await page.eval_on_selector_all(
            "a[href]",
            """items => items.map(item => ({
                href: item.href || '',
                text: (item.innerText || item.textContent || '').trim()
            }))""",
        )
        for candidate in anchors:
            href = clean_text(candidate.get("href"))
            text = clean_text(candidate.get("text"))
            if not href:
                continue
            if not self._candidate_matches(href, text, extensions, keywords):
                continue
            artifact = await self._download_direct_link(context, line, kind, href, text, extensions)
            if artifact:
                return artifact

        for keyword in keywords:
            locator = await self._first_visible(
                page,
                (
                    f'text="{keyword}"',
                    f'role=link[name="{keyword}"]',
                    f'role=button[name="{keyword}"]',
                ),
            )
            if not locator:
                continue
            try:
                async with page.expect_download(timeout=4000) as download_info:
                    await locator.click(timeout=4000)
                download = await download_info.value
                suggested = safe_filename(download.suggested_filename or f"{kind}.{next(iter(extensions), 'bin')}")
                content_path = self.storage.line_root(str(line.get("lote_id")), str(line.get("id"))) / suggested
                await download.save_as(str(content_path))
                return DownloadArtifact(
                    kind=kind,
                    file_name=suggested,
                    file_path=str(content_path),
                    size_bytes=content_path.stat().st_size,
                    detected_inner_type=content_path.suffix.lstrip(".").upper(),
                )
            except (PlaywrightTimeoutError, Error):
                continue
        return None

    async def _download_direct_link(
        self,
        context: BrowserContext,
        line: dict[str, object],
        kind: str,
        href: str,
        text: str,
        extensions: set[str],
    ) -> DownloadArtifact | None:
        try:
            response = await context.request.get(href, timeout=self.settings.browser_timeout_ms)
        except Error:
            return None
        if not response.ok:
            return None
        parsed = urlparse(href)
        extension = Path(parsed.path).suffix.lower().lstrip(".")
        if extension and extension not in extensions:
            return None
        if not extension:
            extension = self._infer_extension_from_headers(
                str(response.headers.get("content-type") or ""),
                str(response.headers.get("content-disposition") or ""),
            )
        if extension not in extensions:
            return None
        file_name = Path(parsed.path).name or f"{safe_filename(text or kind)}.{extension}"
        file_name = safe_filename(file_name)
        content = await response.body()
        file_path = self.storage.line_file(str(line.get("lote_id")), str(line.get("id")), file_name, content)
        return DownloadArtifact(
            kind=kind,
            file_name=file_name,
            file_path=file_path,
            content_type=str(response.headers.get("content-type") or ""),
            size_bytes=len(content),
            detected_inner_type=Path(file_path).suffix.lstrip(".").upper(),
        )

    async def _first_visible(self, page: Page, selectors: tuple[str, ...]):
        for selector in selectors:
            locator = page.locator(selector).first
            try:
                if await locator.is_visible(timeout=1000):
                    return locator
            except Error:
                continue
        return None

    async def _click_by_text(self, page: Page, texts: tuple[str, ...], optional: bool = False) -> bool:
        for text in texts:
            locator = await self._first_visible(
                page,
                (
                    f'text="{text}"',
                    f'role=button[name="{text}"]',
                    f'role=link[name="{text}"]',
                ),
            )
            if not locator:
                continue
            try:
                await locator.click(timeout=3000)
                await page.wait_for_load_state("domcontentloaded")
                return True
            except (PlaywrightTimeoutError, Error):
                continue
        if optional:
            return False
        raise BotLineError(
            ErrorCode.ERRO_LAYOUT_ALTERADO,
            "Nao foi possivel abrir a area de faturas porque os atalhos do portal nao apareceram.",
        )

    async def _extract_login_error(self, page: Page) -> BotLineError | None:
        messages = []
        for selector in (
            '[role="alert"]',
            ".error",
            ".erro",
            ".alert",
            ".warning",
            ".invalid-feedback",
        ):
            locator = page.locator(selector)
            try:
                count = min(await locator.count(), 4)
            except Error:
                continue
            for index in range(count):
                text = clean_text(await locator.nth(index).inner_text())
                if text:
                    messages.append(text)
        merged = " | ".join(messages).lower()
        if not merged:
            content = clean_text(await page.content()).lower()
            if "captcha" in content:
                return BotLineError(ErrorCode.ERRO_CAPTCHA, "O portal exigiu captcha e a automacao foi interrompida.")
            if "sem permiss" in content:
                return BotLineError(ErrorCode.ERRO_SEM_PERMISSAO, "O usuario autenticado nao tem permissao para acessar a conta.")
            return None
        if "senha" in merged and ("incorret" in merged or "invalid" in merged):
            return BotLineError(ErrorCode.ERRO_SENHA_INVALIDA, "O portal informou senha invalida.")
        if "usuario" in merged and ("incorret" in merged or "invalid" in merged):
            return BotLineError(ErrorCode.ERRO_USUARIO_INVALIDO, "O portal informou usuario invalido.")
        if "captcha" in merged:
            return BotLineError(ErrorCode.ERRO_CAPTCHA, "O portal exigiu captcha e a automacao foi interrompida.")
        if "indispon" in merged or "temporariamente fora" in merged:
            return BotLineError(ErrorCode.ERRO_PORTAL_INDISPONIVEL, "O portal retornou indisponibilidade durante o login.")
        return BotLineError(ErrorCode.ERRO_LOGIN, messages[0])

    async def _log(self, line: dict[str, object], step: str, message: str, technical: str = "") -> str:
        return self.storage.append_line_log(
            str(line.get("lote_id")),
            str(line.get("id")),
            {
                "timestamp": now_local(self.settings.timezone_name),
                "connector": self.connector_name,
                "step": step,
                "message": message,
                "technical": technical,
            },
        )

    def _candidate_matches(self, href: str, text: str, extensions: set[str], keywords: tuple[str, ...]) -> bool:
        href_lower = href.lower()
        text_lower = text.lower()
        ext_hit = any(f".{extension}" in href_lower for extension in extensions)
        keyword_hit = any(keyword in href_lower or keyword in text_lower for keyword in keywords)
        return ext_hit or keyword_hit

    @staticmethod
    def _infer_extension_from_headers(content_type: str, content_disposition: str) -> str:
        lowered = f"{content_type.lower()} {content_disposition.lower()}"
        if ".pdf" in lowered or "application/pdf" in lowered:
            return "pdf"
        if ".csv" in lowered or "text/csv" in lowered:
            return "csv"
        if ".txt" in lowered or "text/plain" in lowered:
            return "txt"
        if ".zip" in lowered or "zip" in lowered:
            return "zip"
        if ".ae" in lowered:
            return "ae"
        return "bin"

    @staticmethod
    def _summarize_outcome(
        *,
        line: dict[str, object],
        pdf_artifact: DownloadArtifact | None,
        ae_artifact: DownloadArtifact | None,
    ) -> tuple[LineStatus, ErrorCode | None, str]:
        requested_pdf = bool(line.get("baixar_pdf"))
        requested_ae = bool(line.get("baixar_ae"))
        pdf_ok = (not requested_pdf) or bool(pdf_artifact)
        ae_ok = (not requested_ae) or bool(ae_artifact)

        if pdf_ok and ae_ok and (pdf_artifact or ae_artifact):
            return LineStatus.SUCESSO_TOTAL, None, ""
        if pdf_artifact or ae_artifact:
            if requested_pdf and not pdf_artifact:
                return LineStatus.SUCESSO_PARCIAL, ErrorCode.ERRO_PDF_NAO_DISPONIVEL, "O PDF nao estava disponivel nesta fatura."
            if requested_ae and not ae_artifact:
                return LineStatus.SUCESSO_PARCIAL, ErrorCode.ERRO_AE_NAO_DISPONIVEL, "O arquivo eletronico nao estava disponivel nesta fatura."
        if requested_pdf and requested_ae:
            return (
                LineStatus.ERRO_PROCESSAMENTO,
                ErrorCode.ERRO_FATURA_NAO_ENCONTRADA,
                "Nenhum arquivo foi baixado porque a fatura nao apareceu com os identificadores desta linha.",
            )
        if requested_pdf:
            return LineStatus.ERRO_PROCESSAMENTO, ErrorCode.ERRO_PDF_NAO_DISPONIVEL, "O PDF da fatura nao foi encontrado."
        if requested_ae:
            return (
                LineStatus.ERRO_PROCESSAMENTO,
                ErrorCode.ERRO_AE_NAO_DISPONIVEL,
                "O arquivo eletronico da fatura nao foi encontrado.",
            )
        return LineStatus.ERRO_PROCESSAMENTO, ErrorCode.ERRO_DESCONHECIDO, "Nenhum download foi solicitado para esta linha."
