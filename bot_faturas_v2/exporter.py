from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from .storage import StorageManager


RETURN_COLUMNS = [
    "lote_id",
    "linha_id",
    "numero_ticket",
    "empresa",
    "operadora",
    "titulo_conta",
    "mes_referencia",
    "vencimento",
    "status_final",
    "pdf_status",
    "ae_status",
    "pdf_nome_arquivo",
    "ae_nome_arquivo",
    "pdf_caminho",
    "ae_caminho",
    "erro_codigo",
    "erro_descricao",
    "observacao_execucao",
    "processado_em",
]


class ResultExporter:
    def __init__(self, storage: StorageManager) -> None:
        self.storage = storage

    def export_lot(self, lote_id: str, lines: list[dict[str, object]], lot_payload: dict[str, object]) -> tuple[str, str]:
        report_dir = self.storage.lot_reports_root(lote_id)
        frame = pd.DataFrame([self._map_line(row) for row in lines], columns=RETURN_COLUMNS)
        report_path = report_dir / f"resultado_lote_{lote_id}.xlsx"
        with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
            frame.to_excel(writer, sheet_name="resultado", index=False)

        consolidated_log_path = report_dir / f"log_lote_{lote_id}.json"
        consolidated_log_path.write_text(
            json.dumps({"lote": lot_payload, "linhas": lines}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return str(report_path), str(consolidated_log_path)

    @staticmethod
    def _map_line(row: dict[str, object]) -> dict[str, object]:
        pdf_name = Path(str(row.get("arquivo_pdf_path") or "")).name if row.get("arquivo_pdf_path") else ""
        ae_name = Path(str(row.get("arquivo_ae_path") or "")).name if row.get("arquivo_ae_path") else ""
        return {
            "lote_id": row.get("lote_id", ""),
            "linha_id": row.get("id", ""),
            "numero_ticket": row.get("numero_ticket", ""),
            "empresa": row.get("empresa", ""),
            "operadora": row.get("operadora_padronizada", ""),
            "titulo_conta": row.get("titulo_conta", ""),
            "mes_referencia": row.get("mes_referencia", ""),
            "vencimento": row.get("vencimento", ""),
            "status_final": row.get("status_final", row.get("status_processamento", "")),
            "pdf_status": row.get("pdf_status", ""),
            "ae_status": row.get("ae_status", ""),
            "pdf_nome_arquivo": pdf_name,
            "ae_nome_arquivo": ae_name,
            "pdf_caminho": row.get("arquivo_pdf_path", ""),
            "ae_caminho": row.get("arquivo_ae_path", ""),
            "erro_codigo": row.get("erro_codigo", ""),
            "erro_descricao": row.get("erro_descricao", ""),
            "observacao_execucao": row.get("observacao_execucao", ""),
            "processado_em": row.get("processado_em", ""),
        }
