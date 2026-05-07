from __future__ import annotations

from playwright.async_api import Browser

from ..audit import AuditService
from ..config import BotFaturasSettings
from ..connectors import resolve_connector
from ..database import BotFaturasDatabase
from ..exporter import ResultExporter
from ..security import CredentialCipher
from ..storage import StorageManager


class LineProcessingService:
    def __init__(
        self,
        *,
        settings: BotFaturasSettings,
        database: BotFaturasDatabase,
        storage: StorageManager,
        cipher: CredentialCipher,
        audit: AuditService,
        exporter: ResultExporter,
    ) -> None:
        self.settings = settings
        self.database = database
        self.storage = storage
        self.cipher = cipher
        self.audit = audit
        self.exporter = exporter

    async def process_line(self, browser: Browser, line_id: str) -> None:
        line = self.database.get_line_for_processing(line_id, self.cipher)
        if not line:
            return
        if str(line.get("status_processamento")) == "ERRO_VALIDACAO":
            await self.finalize_lot_if_ready(str(line.get("lote_id")))
            return

        self.database.mark_line_processing(line_id)
        connector_class = resolve_connector(line)
        connector = connector_class(self.settings, self.storage)
        result = await connector.process(browser, line)

        self.database.update_line_result(
            line_id,
            status_processamento=str(result.status_final),
            status_final=str(result.status_final),
            pdf_status=str(result.pdf_status),
            ae_status=str(result.ae_status),
            erro_codigo=str(result.erro_codigo or ""),
            erro_descricao=result.erro_descricao,
            arquivo_pdf_path=result.pdf_artifact.file_path if result.pdf_artifact else "",
            arquivo_ae_path=result.ae_artifact.file_path if result.ae_artifact else "",
            log_execucao=result.log_execucao_path,
            screenshot_erro_path=result.screenshot_erro_path,
            html_erro_path=result.html_erro_path,
            observacao_execucao=result.observacao_execucao,
            processado_em=result.processado_em,
        )
        self.audit.record(
            "line",
            line_id,
            "LINE_PROCESSED",
            {
                "lot_id": line.get("lote_id"),
                "status_final": str(result.status_final),
                "error_code": str(result.erro_codigo or ""),
            },
        )
        await self.finalize_lot_if_ready(str(line.get("lote_id")))

    async def finalize_lot_if_ready(self, lot_id: str) -> None:
        lot = self.database.refresh_lot_summary(lot_id)
        if not lot:
            return
        if lot.get("status") == "PROCESSANDO":
            return
        lines = self.database.get_lines_for_lot(lot_id)
        report_path, consolidated_log_path = self.exporter.export_lot(lot_id, lines, lot)
        archive_path = self.storage.create_lot_archive(lot_id)
        self.database.mark_lot_report_files(
            lot_id,
            report_path=report_path,
            archive_path=archive_path,
            consolidated_log_path=consolidated_log_path,
        )
        self.audit.record(
            "lot",
            lot_id,
            "LOT_FINALIZED",
            {
                "status": lot.get("status"),
                "report_path": report_path,
                "archive_path": archive_path,
            },
        )
