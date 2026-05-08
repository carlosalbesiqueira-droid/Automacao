from __future__ import annotations

from .generic import GenericPortalConnector


class OiConnector(GenericPortalConnector):
    connector_name = "OI"
    invoice_nav_texts = ("Oi", "Faturas", "Financeiro", "Boletos")
