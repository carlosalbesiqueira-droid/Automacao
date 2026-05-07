from __future__ import annotations

import json
import uuid
from pathlib import Path

from ..audit import AuditService
from ..config import BotFaturasSettings
from ..constants import LineStatus, LotStatus
from ..database import BotFaturasDatabase
from ..normalization import clean_text, mask_secret, now_local
from ..parser import SpreadsheetParser
from ..security import CredentialCipher
from ..storage import StorageManager


class LotService:
    def __init__(
        self,
        *,
        settings: BotFaturasSettings,
        database: BotFaturasDatabase,
        storage: StorageManager,
        parser: SpreadsheetParser,
        cipher: CredentialCipher,
        audit: AuditService,
        processing_service,
        queue_manager,
    ) -> None:
        self.settings = settings
        self.database = database
        self.storage = storage
        self.parser = parser
        self.cipher = cipher
        self.audit = audit
        self.processing_service = processing_service
        self.queue_manager = queue_manager

    async def create_lot_from_upload(self, *, file_name: str, content: bytes) -> dict[str, object]:
        suffix = Path(file_name).suffix.lower()
        if suffix not in {".csv", ".xlsx", ".xlsm", ".xls"}:
            raise ValueError("Envie apenas arquivos CSV, XLSX, XLSM ou XLS.")
        if len(content) > self.settings.max_upload_size_bytes:
            raise ValueError(
                f"O arquivo enviado excede o limite de {self.settings.max_upload_size_mb} MB configurado para upload."
            )

        lot_id = uuid.uuid4().hex[:14]
        upload_path = self.storage.save_upload(lot_id, file_name, content)
        parsed = self.parser.parse_file(upload_path)
        valid_lines = [line for line in parsed.lines if not line.validation_errors]
        invalid_lines = [line for line in parsed.lines if line.validation_errors]
        lot_status = LotStatus.PROCESSANDO if valid_lines else LotStatus.CONCLUIDO_COM_ERROS

        self.database.create_lot(
            lot_id=lot_id,
            source_filename=file_name,
            source_type="upload",
            source_path=upload_path,
            sheet_name=parsed.sheet_name,
            status=str(lot_status),
            total_linhas=len(parsed.lines),
            linhas_validas=len(valid_lines),
            linhas_invalidas=len(invalid_lines),
            header_mapping=parsed.header_mapping,
            source_headers=parsed.source_headers,
            warnings=parsed.warnings,
        )

        queued_ids: list[str] = []
        for line in parsed.lines:
            line.lote_id = lot_id
            line.criado_em = now_local(self.settings.timezone_name)
            if not line.validation_errors:
                line.status_processamento = LineStatus.NA_FILA
            line_id = self.database.insert_line(lot_id, line, self.cipher)
            if not line.validation_errors:
                queued_ids.append(line_id)

        self.audit.record(
            "lot",
            lot_id,
            "LOT_CREATED",
            {
                "source_filename": file_name,
                "valid_lines": len(valid_lines),
                "invalid_lines": len(invalid_lines),
            },
        )
        if queued_ids:
            await self.queue_manager.enqueue_many(queued_ids)
        else:
            await self.processing_service.finalize_lot_if_ready(lot_id)
        return self.get_lot_payload(lot_id)

    def list_history(self, limit: int = 50) -> list[dict[str, object]]:
        return [self._serialize_lot(lot) for lot in self.database.list_lots(limit=limit)]

    def get_lot_payload(self, lot_id: str) -> dict[str, object]:
        lot = self.database.get_lot(lot_id)
        if not lot:
            raise ValueError("Lote nao encontrado.")
        lines = self.database.get_lines_for_lot(lot_id)
        return {
            "lot": self._serialize_lot(lot),
            "lines": [self._serialize_line_summary(line) for line in lines],
        }

    def get_line_detail(self, line_id: str) -> dict[str, object]:
        line = self.database.get_line(line_id)
        if not line:
            raise ValueError("Linha nao encontrada.")
        original_data = json.loads(line.get("original_data_json") or "{}")
        normalized_data = json.loads(line.get("normalized_data_json") or "{}")
        return {
            "line": self._serialize_line_summary(line),
            "original_data": self._mask_payload(original_data),
            "normalized_data": self._mask_payload(normalized_data),
            "evidences": {
                "screenshot_erro_path": line.get("screenshot_erro_path", ""),
                "html_erro_path": line.get("html_erro_path", ""),
                "log_execucao": line.get("log_execucao", ""),
            },
        }

    async def reprocess_lot(self, lot_id: str) -> dict[str, object]:
        lines = self.database.get_lines_for_lot(lot_id)
        queue_ids: list[str] = []
        for line in lines:
            if str(line.get("status_processamento")) == str(LineStatus.ERRO_VALIDACAO):
                continue
            self.database.reset_line_for_reprocess(str(line.get("id")))
            queue_ids.append(str(line.get("id")))
        self.database.set_lot_status(lot_id, str(LotStatus.PROCESSANDO))
        self.audit.record("lot", lot_id, "LOT_REPROCESS_REQUESTED", {"queued_lines": len(queue_ids)})
        if queue_ids:
            await self.queue_manager.enqueue_many(queue_ids)
        return self.get_lot_payload(lot_id)

    async def reprocess_line(self, line_id: str) -> dict[str, object]:
        line = self.database.get_line(line_id)
        if not line:
            raise ValueError("Linha nao encontrada.")
        if str(line.get("status_processamento")) == str(LineStatus.ERRO_VALIDACAO):
            raise ValueError("Esta linha esta em ERRO_VALIDACAO e precisa de nova planilha ou correcao na origem.")
        self.database.reset_line_for_reprocess(line_id)
        self.database.set_lot_status(str(line.get("lote_id")), str(LotStatus.PROCESSANDO))
        self.audit.record("line", line_id, "LINE_REPROCESS_REQUESTED", {"lot_id": line.get("lote_id")})
        await self.queue_manager.enqueue(line_id)
        return self.get_line_detail(line_id)

    @staticmethod
    def _mask_payload(payload: dict[str, object]) -> dict[str, object]:
        masked = {}
        for key, value in payload.items():
            normalized = key.lower()
            if "senha" in normalized or "password" in normalized:
                masked[key] = "***"
            elif "usuario" in normalized or "login" in normalized:
                masked[key] = mask_secret(value)
            else:
                masked[key] = value
        return masked

    def _serialize_lot(self, lot: dict[str, object]) -> dict[str, object]:
        return {
            "id": lot.get("id", ""),
            "source_filename": lot.get("source_filename", ""),
            "source_type": lot.get("source_type", ""),
            "sheet_name": lot.get("sheet_name", ""),
            "status": lot.get("status", ""),
            "total_linhas": lot.get("total_linhas", 0),
            "linhas_validas": lot.get("linhas_validas", 0),
            "linhas_invalidas": lot.get("linhas_invalidas", 0),
            "linhas_processadas": lot.get("linhas_processadas", 0),
            "sucesso_total": lot.get("sucesso_total", 0),
            "sucesso_parcial": lot.get("sucesso_parcial", 0),
            "erro_total": lot.get("erro_total", 0),
            "headers": json.loads(lot.get("headers_json") or "[]"),
            "mapping": json.loads(lot.get("mapping_json") or "{}"),
            "warnings": json.loads(lot.get("warnings_json") or "[]"),
            "report_path": lot.get("report_path", ""),
            "archive_path": lot.get("archive_path", ""),
            "consolidated_log_path": lot.get("consolidated_log_path", ""),
            "criado_em": lot.get("criado_em", ""),
            "atualizado_em": lot.get("atualizado_em", ""),
            "finalizado_em": lot.get("finalizado_em", ""),
        }

    def _serialize_line_summary(self, line: dict[str, object]) -> dict[str, object]:
        return {
            "id": line.get("id", ""),
            "lote_id": line.get("lote_id", ""),
            "numero_linha_origem": line.get("numero_linha_origem", 0),
            "numero_ticket": line.get("numero_ticket", ""),
            "empresa": line.get("empresa", ""),
            "cnpj": line.get("cnpj", ""),
            "operadora_original": line.get("operadora_original", ""),
            "operadora_padronizada": line.get("operadora_padronizada", ""),
            "plataforma_login": line.get("plataforma_login", ""),
            "link_portal": line.get("link_portal", ""),
            "codigo_cliente": line.get("codigo_cliente", ""),
            "conta": line.get("conta", ""),
            "titulo_conta": line.get("titulo_conta", ""),
            "mes_referencia": line.get("mes_referencia", ""),
            "vencimento": line.get("vencimento", ""),
            "nome_arquivo_esperado": line.get("nome_arquivo_esperado", ""),
            "baixar_pdf": bool(line.get("baixar_pdf")),
            "baixar_ae": bool(line.get("baixar_ae")),
            "formatos_aceitos_ae": clean_text(line.get("formatos_aceitos_ae")),
            "status_processamento": line.get("status_processamento", ""),
            "status_final": line.get("status_final", ""),
            "pdf_status": line.get("pdf_status", ""),
            "ae_status": line.get("ae_status", ""),
            "erro_codigo": line.get("erro_codigo", ""),
            "erro_descricao": line.get("erro_descricao", ""),
            "arquivo_pdf_path": line.get("arquivo_pdf_path", ""),
            "arquivo_ae_path": line.get("arquivo_ae_path", ""),
            "observacao_execucao": line.get("observacao_execucao", ""),
            "log_execucao": line.get("log_execucao", ""),
            "screenshot_erro_path": line.get("screenshot_erro_path", ""),
            "html_erro_path": line.get("html_erro_path", ""),
            "processado_em": line.get("processado_em", ""),
            "tentativas": line.get("attempt_count", 0),
        }
