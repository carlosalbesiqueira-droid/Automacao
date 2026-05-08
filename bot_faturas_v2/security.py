from __future__ import annotations

from pathlib import Path

from cryptography.fernet import Fernet

from .config import BotFaturasSettings
from .normalization import clean_text, ensure_directory


class CredentialCipher:
    def __init__(self, key: bytes) -> None:
        self._fernet = Fernet(key)

    @classmethod
    def from_settings(cls, settings: BotFaturasSettings) -> "CredentialCipher":
        if settings.encryption_key_value:
            return cls(settings.encryption_key_value.encode("utf-8"))

        key_file = Path(settings.encryption_key_file)
        ensure_directory(key_file.parent)
        if key_file.exists():
            return cls(key_file.read_bytes().strip())

        key = Fernet.generate_key()
        key_file.write_bytes(key)
        return cls(key)

    def encrypt(self, value: object) -> str:
        text = clean_text(value)
        if not text:
            return ""
        return self._fernet.encrypt(text.encode("utf-8")).decode("utf-8")

    def decrypt(self, value: object) -> str:
        token = clean_text(value)
        if not token:
            return ""
        return self._fernet.decrypt(token.encode("utf-8")).decode("utf-8")
