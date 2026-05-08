from __future__ import annotations

from .generic import GenericPortalConnector


class AlgarConnector(GenericPortalConnector):
    connector_name = "ALGAR"
    invoice_nav_texts = ("Algar", "Faturas", "Financeiro", "Minhas faturas")
