from __future__ import annotations

from .generic import GenericPortalConnector


class TimConnector(GenericPortalConnector):
    connector_name = "TIM"
    invoice_nav_texts = ("Meu TIM", "Faturas", "2 via", "Financeiro")
