from __future__ import annotations

from .generic import GenericPortalConnector


class HubsoftConnector(GenericPortalConnector):
    connector_name = "HUBSOFT"
    invoice_nav_texts = ("Financeiro", "Titulos", "Boletos")
