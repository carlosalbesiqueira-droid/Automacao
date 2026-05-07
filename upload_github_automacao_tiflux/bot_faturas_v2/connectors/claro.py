from __future__ import annotations

from .generic import GenericPortalConnector


class ClaroConnector(GenericPortalConnector):
    connector_name = "CLARO"
    invoice_nav_texts = ("Minha Claro", "Faturas", "Conta", "Segunda via", "Downloads")
