from __future__ import annotations

from .generic import GenericPortalConnector


class EmbratelConnector(GenericPortalConnector):
    connector_name = "EMBRATEL"
    invoice_nav_texts = ("Embratel", "Faturas", "Conta", "Downloads")
