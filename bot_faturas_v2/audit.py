from __future__ import annotations

from .database import BotFaturasDatabase


class AuditService:
    def __init__(self, database: BotFaturasDatabase) -> None:
        self.database = database

    def record(self, entity_type: str, entity_id: str, action: str, payload: dict[str, object]) -> None:
        self.database.insert_audit_event(entity_type, entity_id, action, payload)
