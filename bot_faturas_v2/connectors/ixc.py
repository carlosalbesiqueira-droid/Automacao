from __future__ import annotations

from .generic import GenericPortalConnector


class IXCConnector(GenericPortalConnector):
    connector_name = "IXC"
    invoice_nav_texts = ("Financeiro", "Faturas", "Boletos")
