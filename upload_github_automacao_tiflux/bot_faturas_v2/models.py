from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any

from .constants import DEFAULT_AE_FORMATS, DownloadStatus, ErrorCode, LineStatus


@dataclass(slots=True)
class DownloadArtifact:
    kind: str
    file_name: str
    file_path: str
    content_type: str = ""
    size_bytes: int = 0
    detected_inner_type: str = ""


@dataclass(slots=True)
class NormalizedInvoiceLine:
    lote_id: str = ""
    linha_id: str = ""
    numero_linha_origem: int = 0
    numero_ticket: str = ""
    empresa: str = ""
    cnpj: str = ""
    operadora_original: str = ""
    operadora_padronizada: str = ""
    plataforma_login: str = ""
    link_portal: str = ""
    usuario_login: str = ""
    senha: str = ""
    codigo_cliente: str = ""
    conta: str = ""
    titulo_conta: str = ""
    mes_referencia: str = ""
    vencimento: str = ""
    nome_arquivo_esperado: str = ""
    baixar_pdf: bool = True
    baixar_ae: bool = True
    formatos_aceitos_ae: tuple[str, ...] = DEFAULT_AE_FORMATS
    status_processamento: LineStatus = LineStatus.NA_FILA
    erro_codigo: str = ""
    erro_descricao: str = ""
    arquivo_pdf_path: str = ""
    arquivo_ae_path: str = ""
    log_execucao: str = ""
    screenshot_erro_path: str = ""
    html_erro_path: str = ""
    criado_em: str = ""
    atualizado_em: str = ""
    original_data: dict[str, Any] = field(default_factory=dict)
    normalized_data: dict[str, Any] = field(default_factory=dict)
    validation_errors: list[str] = field(default_factory=list)

    def masked(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["senha"] = "***" if self.senha else ""
        payload["usuario_login"] = self.usuario_login[:3] + "***" if self.usuario_login else ""
        return payload


@dataclass(slots=True)
class ParsedSpreadsheet:
    source_name: str
    sheet_name: str
    source_headers: list[str]
    header_mapping: dict[str, str]
    lines: list[NormalizedInvoiceLine]
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class LineProcessingResult:
    status_final: LineStatus
    pdf_status: DownloadStatus
    ae_status: DownloadStatus
    erro_codigo: ErrorCode | None
    erro_descricao: str
    observacao_execucao: str
    downloads: list[DownloadArtifact] = field(default_factory=list)
    screenshot_erro_path: str = ""
    html_erro_path: str = ""
    log_execucao_path: str = ""
    processado_em: str = ""

    @property
    def pdf_artifact(self) -> DownloadArtifact | None:
        for artifact in self.downloads:
            if artifact.kind == "pdf":
                return artifact
        return None

    @property
    def ae_artifact(self) -> DownloadArtifact | None:
        for artifact in self.downloads:
            if artifact.kind == "ae":
                return artifact
        return None
