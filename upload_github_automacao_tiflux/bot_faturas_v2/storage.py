from __future__ import annotations

import json
import shutil
from pathlib import Path

from .normalization import ensure_directory, safe_filename


class StorageManager:
    def __init__(self, root: Path) -> None:
        self.root = ensure_directory(Path(root))

    def lot_root(self, lote_id: str) -> Path:
        return ensure_directory(self.root / lote_id)

    def line_root(self, lote_id: str, linha_id: str) -> Path:
        return ensure_directory(self.lot_root(lote_id) / linha_id)

    def lot_reports_root(self, lote_id: str) -> Path:
        return ensure_directory(self.lot_root(lote_id) / "_reports")

    def lot_input_root(self, lote_id: str) -> Path:
        return ensure_directory(self.lot_root(lote_id) / "_input")

    def save_upload(self, lote_id: str, file_name: str, content: bytes) -> str:
        target = self.lot_input_root(lote_id) / safe_filename(file_name, "upload.xlsx")
        target.write_bytes(content)
        return str(target)

    def line_file(self, lote_id: str, linha_id: str, file_name: str, content: bytes) -> str:
        target = self.line_root(lote_id, linha_id) / safe_filename(file_name)
        target.write_bytes(content)
        return str(target)

    def line_text_file(self, lote_id: str, linha_id: str, file_name: str, content: str) -> str:
        target = self.line_root(lote_id, linha_id) / safe_filename(file_name)
        target.write_text(content, encoding="utf-8")
        return str(target)

    def line_json_file(self, lote_id: str, linha_id: str, file_name: str, payload: object) -> str:
        target = self.line_root(lote_id, linha_id) / safe_filename(file_name)
        target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return str(target)

    def append_line_log(self, lote_id: str, linha_id: str, entry: dict[str, object]) -> str:
        target = self.line_root(lote_id, linha_id) / "log_tecnico.jsonl"
        with target.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(entry, ensure_ascii=False) + "\n")
        return str(target)

    def create_lot_archive(self, lote_id: str) -> str:
        lot_root = self.lot_root(lote_id)
        archive_base = ensure_directory(self.root / "_archives") / f"lote_{lote_id}"
        return shutil.make_archive(str(archive_base), "zip", root_dir=lot_root)
