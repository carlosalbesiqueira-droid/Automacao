from __future__ import annotations

from dataclasses import dataclass

from .constants import ErrorCode


@dataclass(slots=True)
class BotLineError(Exception):
    code: ErrorCode
    description: str
    technical_details: str = ""

    def __str__(self) -> str:
        return self.description
