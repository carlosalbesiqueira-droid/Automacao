from __future__ import annotations

from .algar import AlgarConnector
from .claro import ClaroConnector
from .embratel import EmbratelConnector
from .generic import GenericPortalConnector
from .hubsoft import HubsoftConnector
from .ixc import IXCConnector
from .oi import OiConnector
from .tim import TimConnector
from .vivo import VivoConnector


def resolve_connector(line: dict[str, object]):
    platform = str(line.get("plataforma_login") or "").upper()
    operator = str(line.get("operadora_padronizada") or "").upper()

    if platform == "MEU_TIM" or operator == "TIM":
        return TimConnector
    if platform in {"MVE_VIVO", "VIVO_EM_DIA"} or operator == "VIVO":
        return VivoConnector
    if platform in {"CLARO_CONTA_ONLINE", "CLARO_OPERA360", "CLARO_PRESTO360"} or operator == "CLARO":
        return ClaroConnector
    if operator == "EMBRATEL":
        return EmbratelConnector
    if platform == "OI_SOLUCOES" or operator == "OI":
        return OiConnector
    if platform == "ALGAR_PORTAL" or operator == "ALGAR":
        return AlgarConnector
    if platform == "HUBSOFT" or operator == "HUBSOFT":
        return HubsoftConnector
    if platform == "IXC" or operator == "IXC":
        return IXCConnector
    return GenericPortalConnector
