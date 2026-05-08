from __future__ import annotations

import json
import sqlite3
import threading
import uuid
from pathlib import Path
from typing import Any

from .constants import LineStatus, LotStatus, TERMINAL_LINE_STATUSES
from .models import NormalizedInvoiceLine
from .normalization import dump_json, ensure_directory, now_local
from .security import CredentialCipher


class BotFaturasDatabase:
    def __init__(self, database_path: Path, timezone_name: str) -> None:
        self.database_path = Path(database_path)
        ensure_directory(self.database_path.parent)
        self.timezone_name = timezone_name
        self._lock = threading.RLock()
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.database_path, check_same_thread=False)
        connection.row_factory = sqlite3.Row
        return connection

    def _initialize(self) -> None:
        with self._connect() as connection:
            connection.executescript(
                """
                PRAGMA journal_mode=WAL;
                CREATE TABLE IF NOT EXISTS lots (
                    id TEXT PRIMARY KEY,
                    source_filename TEXT NOT NULL,
                    source_type TEXT NOT NULL,
                    source_path TEXT NOT NULL,
                    sheet_name TEXT,
                    status TEXT NOT NULL,
                    total_linhas INTEGER NOT NULL,
                    linhas_validas INTEGER NOT NULL,
                    linhas_invalidas INTEGER NOT NULL,
                    linhas_processadas INTEGER NOT NULL DEFAULT 0,
                    sucesso_total INTEGER NOT NULL DEFAULT 0,
                    sucesso_parcial INTEGER NOT NULL DEFAULT 0,
                    erro_total INTEGER NOT NULL DEFAULT 0,
                    mapping_json TEXT NOT NULL,
                    headers_json TEXT NOT NULL,
                    warnings_json TEXT NOT NULL,
                    report_path TEXT,
                    archive_path TEXT,
                    consolidated_log_path TEXT,
                    criado_em TEXT NOT NULL,
                    atualizado_em TEXT NOT NULL,
                    finalizado_em TEXT
                );
                CREATE TABLE IF NOT EXISTS lines (
                    id TEXT PRIMARY KEY,
                    lote_id TEXT NOT NULL,
                    numero_linha_origem INTEGER NOT NULL,
                    original_data_json TEXT NOT NULL,
                    normalized_data_json TEXT NOT NULL,
                    numero_ticket TEXT,
                    empresa TEXT,
                    cnpj TEXT,
                    operadora_original TEXT,
                    operadora_padronizada TEXT,
                    plataforma_login TEXT,
                    link_portal TEXT,
                    usuario_login_enc TEXT,
                    senha_enc TEXT,
                    codigo_cliente TEXT,
                    conta TEXT,
                    titulo_conta TEXT,
                    mes_referencia TEXT,
                    vencimento TEXT,
                    nome_arquivo_esperado TEXT,
                    baixar_pdf INTEGER NOT NULL DEFAULT 1,
                    baixar_ae INTEGER NOT NULL DEFAULT 1,
                    formatos_aceitos_ae TEXT NOT NULL,
                    status_processamento TEXT NOT NULL,
                    status_final TEXT,
                    pdf_status TEXT,
                    ae_status TEXT,
                    erro_codigo TEXT,
                    erro_descricao TEXT,
                    arquivo_pdf_path TEXT,
                    arquivo_ae_path TEXT,
                    log_execucao TEXT,
                    screenshot_erro_path TEXT,
                    html_erro_path TEXT,
                    observacao_execucao TEXT,
                    criado_em TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    processado_em TEXT,
                    attempt_count INTEGER NOT NULL DEFAULT 0,
                    FOREIGN KEY(lote_id) REFERENCES lots(id)
                );
                CREATE INDEX IF NOT EXISTS idx_lines_lot ON lines(lote_id);
                CREATE TABLE IF NOT EXISTS audit_events (
                    id TEXT PRIMARY KEY,
                    entity_type TEXT NOT NULL,
                    entity_id TEXT NOT NULL,
                    action TEXT NOT NULL,
                    payload_json TEXT NOT NULL,
                    created_at TEXT NOT NULL
                );
                """
            )

    def create_lot(
        self,
        *,
        lot_id: str,
        source_filename: str,
        source_type: str,
        source_path: str,
        sheet_name: str,
        status: str,
        total_linhas: int,
        linhas_validas: int,
        linhas_invalidas: int,
        header_mapping: dict[str, str],
        source_headers: list[str],
        warnings: list[str],
    ) -> None:
        now = now_local(self.timezone_name)
        with self._lock, self._connect() as connection:
            connection.execute(
                """
                INSERT INTO lots (
                    id, source_filename, source_type, source_path, sheet_name, status,
                    total_linhas, linhas_validas, linhas_invalidas, mapping_json, headers_json,
                    warnings_json, criado_em, atualizado_em
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    lot_id,
                    source_filename,
                    source_type,
                    source_path,
                    sheet_name,
                    status,
                    total_linhas,
                    linhas_validas,
                    linhas_invalidas,
                    dump_json(header_mapping),
                    dump_json(source_headers),
                    dump_json(warnings),
                    now,
                    now,
                ),
            )

    def insert_line(self, lot_id: str, line: NormalizedInvoiceLine, cipher: CredentialCipher) -> str:
        line_id = line.linha_id or uuid.uuid4().hex[:12]
        now = now_local(self.timezone_name)
        with self._lock, self._connect() as connection:
            connection.execute(
                """
                INSERT INTO lines (
                    id, lote_id, numero_linha_origem, original_data_json, normalized_data_json,
                    numero_ticket, empresa, cnpj, operadora_original, operadora_padronizada,
                    plataforma_login, link_portal, usuario_login_enc, senha_enc, codigo_cliente,
                    conta, titulo_conta, mes_referencia, vencimento, nome_arquivo_esperado,
                    baixar_pdf, baixar_ae, formatos_aceitos_ae, status_processamento, status_final,
                    erro_codigo, erro_descricao, criado_em, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    line_id,
                    lot_id,
                    line.numero_linha_origem,
                    dump_json(line.original_data),
                    dump_json(line.normalized_data),
                    line.numero_ticket,
                    line.empresa,
                    line.cnpj,
                    line.operadora_original,
                    line.operadora_padronizada,
                    line.plataforma_login,
                    line.link_portal,
                    cipher.encrypt(line.usuario_login),
                    cipher.encrypt(line.senha),
                    line.codigo_cliente,
                    line.conta,
                    line.titulo_conta,
                    line.mes_referencia,
                    line.vencimento,
                    line.nome_arquivo_esperado,
                    1 if line.baixar_pdf else 0,
                    1 if line.baixar_ae else 0,
                    ",".join(line.formatos_aceitos_ae),
                    str(line.status_processamento),
                    str(line.status_processamento) if line.status_processamento == LineStatus.ERRO_VALIDACAO else None,
                    line.erro_codigo,
                    line.erro_descricao,
                    now,
                    now,
                ),
            )
        return line_id

    def insert_audit_event(self, entity_type: str, entity_id: str, action: str, payload: dict[str, object]) -> None:
        with self._lock, self._connect() as connection:
            connection.execute(
                """
                INSERT INTO audit_events (id, entity_type, entity_id, action, payload_json, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    uuid.uuid4().hex,
                    entity_type,
                    entity_id,
                    action,
                    dump_json(payload),
                    now_local(self.timezone_name),
                ),
            )

    def get_lot(self, lot_id: str) -> dict[str, Any] | None:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM lots WHERE id = ?", (lot_id,)).fetchone()
        return dict(row) if row else None

    def list_lots(self, limit: int = 50) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT * FROM lots ORDER BY criado_em DESC LIMIT ?", (limit,)
            ).fetchall()
        return [dict(row) for row in rows]

    def set_lot_status(self, lot_id: str, status: str) -> None:
        with self._lock, self._connect() as connection:
            connection.execute(
                "UPDATE lots SET status = ?, atualizado_em = ? WHERE id = ?",
                (status, now_local(self.timezone_name), lot_id),
            )

    def get_lines_for_lot(self, lot_id: str) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT * FROM lines WHERE lote_id = ? ORDER BY numero_linha_origem ASC",
                (lot_id,),
            ).fetchall()
        return [dict(row) for row in rows]

    def get_line(self, line_id: str) -> dict[str, Any] | None:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM lines WHERE id = ?", (line_id,)).fetchone()
        return dict(row) if row else None

    def list_lines_by_statuses(self, statuses: list[str]) -> list[dict[str, Any]]:
        placeholders = ",".join("?" for _ in statuses)
        with self._connect() as connection:
            rows = connection.execute(
                f"SELECT * FROM lines WHERE status_processamento IN ({placeholders}) ORDER BY criado_em ASC",
                tuple(statuses),
            ).fetchall()
        return [dict(row) for row in rows]

    def get_line_for_processing(self, line_id: str, cipher: CredentialCipher) -> dict[str, Any] | None:
        row = self.get_line(line_id)
        if not row:
            return None
        row["usuario_login"] = cipher.decrypt(row.get("usuario_login_enc"))
        row["senha"] = cipher.decrypt(row.get("senha_enc"))
        row["original_data"] = json.loads(row.get("original_data_json") or "{}")
        row["normalized_data"] = json.loads(row.get("normalized_data_json") or "{}")
        row["formatos_aceitos_ae"] = [
            item.strip().upper() for item in str(row.get("formatos_aceitos_ae") or "").split(",") if item.strip()
        ]
        row["baixar_pdf"] = bool(row.get("baixar_pdf"))
        row["baixar_ae"] = bool(row.get("baixar_ae"))
        return row

    def mark_line_processing(self, line_id: str) -> None:
        with self._lock, self._connect() as connection:
            connection.execute(
                """
                UPDATE lines
                SET status_processamento = ?, updated_at = ?, attempt_count = attempt_count + 1
                WHERE id = ?
                """,
                (str(LineStatus.EM_PROCESSAMENTO), now_local(self.timezone_name), line_id),
            )

    def update_line_result(
        self,
        line_id: str,
        *,
        status_processamento: str,
        status_final: str,
        pdf_status: str,
        ae_status: str,
        erro_codigo: str,
        erro_descricao: str,
        arquivo_pdf_path: str,
        arquivo_ae_path: str,
        log_execucao: str,
        screenshot_erro_path: str,
        html_erro_path: str,
        observacao_execucao: str,
        processado_em: str,
    ) -> None:
        with self._lock, self._connect() as connection:
            connection.execute(
                """
                UPDATE lines
                SET status_processamento = ?, status_final = ?, pdf_status = ?, ae_status = ?,
                    erro_codigo = ?, erro_descricao = ?, arquivo_pdf_path = ?, arquivo_ae_path = ?,
                    log_execucao = ?, screenshot_erro_path = ?, html_erro_path = ?,
                    observacao_execucao = ?, processado_em = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    status_processamento,
                    status_final,
                    pdf_status,
                    ae_status,
                    erro_codigo,
                    erro_descricao,
                    arquivo_pdf_path,
                    arquivo_ae_path,
                    log_execucao,
                    screenshot_erro_path,
                    html_erro_path,
                    observacao_execucao,
                    processado_em,
                    now_local(self.timezone_name),
                    line_id,
                ),
            )

    def reset_line_for_reprocess(self, line_id: str) -> None:
        with self._lock, self._connect() as connection:
            connection.execute(
                """
                UPDATE lines
                SET status_processamento = ?, status_final = NULL, pdf_status = NULL, ae_status = NULL,
                    erro_codigo = '', erro_descricao = '', arquivo_pdf_path = '', arquivo_ae_path = '',
                    log_execucao = '', screenshot_erro_path = '', html_erro_path = '',
                    observacao_execucao = '', processado_em = NULL, updated_at = ?
                WHERE id = ?
                """,
                (str(LineStatus.NA_FILA), now_local(self.timezone_name), line_id),
            )

    def refresh_lot_summary(self, lot_id: str) -> dict[str, Any]:
        with self._lock, self._connect() as connection:
            rows = connection.execute(
                "SELECT status_final, status_processamento FROM lines WHERE lote_id = ?",
                (lot_id,),
            ).fetchall()
            total = len(rows)
            processed = 0
            success_total = 0
            success_partial = 0
            error_total = 0
            has_queue = False
            for row in rows:
                current = row["status_final"] or row["status_processamento"]
                if current in {str(LineStatus.NA_FILA), str(LineStatus.EM_PROCESSAMENTO)}:
                    has_queue = True
                if current in {str(item) for item in TERMINAL_LINE_STATUSES}:
                    processed += 1
                if current == str(LineStatus.SUCESSO_TOTAL):
                    success_total += 1
                elif current == str(LineStatus.SUCESSO_PARCIAL):
                    success_partial += 1
                elif current in {str(LineStatus.ERRO_VALIDACAO), str(LineStatus.ERRO_PROCESSAMENTO)}:
                    error_total += 1

            if has_queue and processed < total:
                lot_status = LotStatus.PROCESSANDO
                finalizado_em = None
            else:
                lot_status = LotStatus.CONCLUIDO if error_total == 0 else LotStatus.CONCLUIDO_COM_ERROS
                finalizado_em = now_local(self.timezone_name)

            connection.execute(
                """
                UPDATE lots
                SET status = ?, linhas_processadas = ?, sucesso_total = ?, sucesso_parcial = ?,
                    erro_total = ?, atualizado_em = ?, finalizado_em = COALESCE(?, finalizado_em)
                WHERE id = ?
                """,
                (
                    str(lot_status),
                    processed,
                    success_total,
                    success_partial,
                    error_total,
                    now_local(self.timezone_name),
                    finalizado_em,
                    lot_id,
                ),
            )
        return self.get_lot(lot_id) or {}

    def mark_lot_report_files(
        self,
        lot_id: str,
        *,
        report_path: str,
        archive_path: str,
        consolidated_log_path: str,
    ) -> None:
        with self._lock, self._connect() as connection:
            connection.execute(
                """
                UPDATE lots
                SET report_path = ?, archive_path = ?, consolidated_log_path = ?, atualizado_em = ?
                WHERE id = ?
                """,
                (report_path, archive_path, consolidated_log_path, now_local(self.timezone_name), lot_id),
            )
