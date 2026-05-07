from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class BotFaturasSettings:
    api_host: str
    api_port: int
    database_path: Path
    storage_root: Path
    encryption_key_file: Path
    encryption_key_value: str
    worker_count: int
    headless: bool
    browser_timeout_ms: int
    timezone_name: str
    max_upload_size_mb: int
    max_downloads_per_line: int
    screenshot_on_error: bool
    cors_allow_origins: list[str]

    @property
    def max_upload_size_bytes(self) -> int:
        return self.max_upload_size_mb * 1024 * 1024

    @classmethod
    def from_env(cls) -> "BotFaturasSettings":
        storage_root = Path(os.getenv("BOT_FATURAS_STORAGE_DIR", "storage/faturas"))
        return cls(
            api_host=os.getenv("BOT_FATURAS_API_HOST", "127.0.0.1"),
            api_port=int(os.getenv("BOT_FATURAS_API_PORT", "8321")),
            database_path=Path(os.getenv("BOT_FATURAS_DB_PATH", str(storage_root / "bot_faturas.db"))),
            storage_root=storage_root,
            encryption_key_file=Path(
                os.getenv("BOT_FATURAS_ENCRYPTION_KEY_FILE", str(storage_root / "_keys" / "bot_faturas.key"))
            ),
            encryption_key_value=os.getenv("BOT_FATURAS_ENCRYPTION_KEY", ""),
            worker_count=max(1, int(os.getenv("BOT_FATURAS_WORKER_COUNT", "2"))),
            headless=os.getenv("BOT_FATURAS_HEADLESS", "true").strip().lower() not in {"0", "false", "no"},
            browser_timeout_ms=int(os.getenv("BOT_FATURAS_TIMEOUT_MS", "45000")),
            timezone_name=os.getenv("BOT_FATURAS_TIMEZONE", "America/Sao_Paulo"),
            max_upload_size_mb=max(5, int(os.getenv("BOT_FATURAS_MAX_UPLOAD_MB", "25"))),
            max_downloads_per_line=max(1, int(os.getenv("BOT_FATURAS_MAX_DOWNLOADS_PER_LINE", "6"))),
            screenshot_on_error=os.getenv("BOT_FATURAS_SCREENSHOT_ON_ERROR", "true").strip().lower()
            not in {"0", "false", "no"},
            cors_allow_origins=[
                item.strip()
                for item in os.getenv(
                    "BOT_FATURAS_CORS_ALLOW_ORIGINS",
                    "http://localhost:3210,http://127.0.0.1:3210,http://localhost:8321,http://127.0.0.1:8321",
                ).split(",")
                if item.strip()
            ],
        )
