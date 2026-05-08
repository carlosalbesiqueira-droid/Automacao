from __future__ import annotations

from .generic import GenericPortalConnector


class VivoConnector(GenericPortalConnector):
    connector_name = "VIVO"
    invoice_nav_texts = ("MVE", "Faturas", "Minhas faturas", "Contas", "Segunda via")
