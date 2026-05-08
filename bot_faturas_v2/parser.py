from __future__ import annotations

from pathlib import Path

import pandas as pd

from .constants import LineStatus
from .models import NormalizedInvoiceLine, ParsedSpreadsheet
from .normalization import (
    clean_text,
    default_portal_url,
    header_score,
    infer_platform,
    normalize_accepted_ae_formats,
    normalize_cnpj,
    normalize_key,
    normalize_url,
    parse_bool,
    parse_date,
    parse_reference_month,
    standardize_operator,
)


STANDARD_FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "numero_ticket": ("numero ticket", "ticket", "numero do ticket", "id ticket"),
    "empresa": ("empresa", "cliente", "cliente empresa", "cliente/empresa", "razao social"),
    "cnpj": ("cnpj", "documento", "cnpj cliente"),
    "operadora_original": ("operadora", "operadora original", "fornecedor", "provedor"),
    "plataforma_login": ("plataforma", "plataforma login", "portal", "sistema"),
    "link_portal": ("link portal", "portal", "url", "link", "site", "link_portal"),
    "usuario_login": ("usuario", "usuario login", "login", "e-mail login", "email login"),
    "senha": ("senha", "password", "pass", "senha portal"),
    "codigo_cliente": ("cod cliente", "codigo cliente", "matricula", "cod_cliente", "cliente id"),
    "conta": ("conta", "numero conta", "conta faturamento", "conta cliente"),
    "titulo_conta": ("titulo", "titulo conta", "descricao fatura", "descricao", "nome conta"),
    "mes_referencia": (
        "mes referencia",
        "mes de consulta",
        "mes consulta",
        "competencia",
        "referencia",
        "mes_ref",
    ),
    "vencimento": ("vencimento", "data vencimento", "dt vencimento"),
    "nome_arquivo_esperado": (
        "nome arquivo",
        "area de faturas nome arquivo",
        "arquivo esperado",
        "nome da fatura",
    ),
    "baixar_pdf": ("baixar pdf", "pdf", "download pdf"),
    "baixar_ae": ("baixar ae", "arquivo eletronico", "baixar arquivo eletronico", "baixar csv/txt/zip"),
    "formatos_aceitos_ae": ("formatos aceitos ae", "tipo arquivo", "arquivo eletronico formato", "formato ae"),
}


class SpreadsheetParser:
    def parse_file(self, input_path: str | Path) -> ParsedSpreadsheet:
        path = Path(input_path)
        frame, sheet_name = self._load_frame(path)
        source_headers = [str(header) for header in frame.columns]
        header_mapping = self._map_headers(source_headers)
        lines: list[NormalizedInvoiceLine] = []

        for row_number, row in enumerate(frame.fillna("").to_dict(orient="records"), start=2):
            original_data = {str(key): clean_text(value) for key, value in row.items()}
            if not any(original_data.values()):
                continue
            line = self._normalize_row(original_data, header_mapping, row_number)
            lines.append(line)

        warnings = []
        if not lines:
            warnings.append("A planilha enviada nao possui linhas preenchidas para processar.")

        return ParsedSpreadsheet(
            source_name=path.name,
            sheet_name=sheet_name,
            source_headers=source_headers,
            header_mapping=header_mapping,
            lines=lines,
            warnings=warnings,
        )

    def _load_frame(self, path: Path) -> tuple[pd.DataFrame, str]:
        suffix = path.suffix.lower()
        if suffix == ".csv":
            return pd.read_csv(path, dtype=str, keep_default_na=False), "CSV"
        if suffix in {".xlsx", ".xlsm", ".xls"}:
            workbook = pd.ExcelFile(path)
            sheet_name = workbook.sheet_names[0]
            return pd.read_excel(workbook, sheet_name=sheet_name, dtype=str), sheet_name
        raise ValueError("Envie um arquivo CSV, XLSX, XLSM ou XLS.")

    def _map_headers(self, headers: list[str]) -> dict[str, str]:
        mapping: dict[str, str] = {}
        for internal_field, aliases in STANDARD_FIELD_ALIASES.items():
            best_header = ""
            best_score = 0
            for header in headers:
                current_score = max(header_score(header, alias) for alias in aliases)
                if current_score > best_score:
                    best_header = header
                    best_score = current_score
            if best_header and best_score >= 84:
                mapping[internal_field] = best_header
        return mapping

    def _normalize_row(
        self,
        original_data: dict[str, str],
        header_mapping: dict[str, str],
        row_number: int,
    ) -> NormalizedInvoiceLine:
        def value(field_name: str) -> str:
            source_header = header_mapping.get(field_name, "")
            return clean_text(original_data.get(source_header, ""))

        operadora_original = value("operadora_original")
        operadora_padronizada = standardize_operator(operadora_original)
        link_portal = normalize_url(value("link_portal"))
        plataforma_login = infer_platform(value("plataforma_login"), link_portal, operadora_padronizada)
        if not link_portal:
            link_portal = default_portal_url(plataforma_login)

        line = NormalizedInvoiceLine(
            numero_linha_origem=row_number,
            numero_ticket=value("numero_ticket"),
            empresa=value("empresa"),
            cnpj=normalize_cnpj(value("cnpj")),
            operadora_original=operadora_original,
            operadora_padronizada=operadora_padronizada,
            plataforma_login=plataforma_login,
            link_portal=link_portal,
            usuario_login=value("usuario_login"),
            senha=value("senha"),
            codigo_cliente=value("codigo_cliente"),
            conta=value("conta"),
            titulo_conta=value("titulo_conta"),
            mes_referencia=parse_reference_month(value("mes_referencia")),
            vencimento=parse_date(value("vencimento")),
            nome_arquivo_esperado=value("nome_arquivo_esperado"),
            baixar_pdf=parse_bool(value("baixar_pdf"), default=True),
            baixar_ae=parse_bool(value("baixar_ae"), default=True),
            formatos_aceitos_ae=normalize_accepted_ae_formats(value("formatos_aceitos_ae")),
            original_data=original_data,
        )
        line.normalized_data = self._build_normalized_preview(line)
        line.validation_errors = self._validate_line(line)
        if line.validation_errors:
            line.status_processamento = LineStatus.ERRO_VALIDACAO
            line.erro_codigo = LineStatus.ERRO_VALIDACAO
            line.erro_descricao = "; ".join(line.validation_errors)
        return line

    @staticmethod
    def _build_normalized_preview(line: NormalizedInvoiceLine) -> dict[str, str]:
        return {
            "numero_ticket": line.numero_ticket,
            "empresa": line.empresa,
            "cnpj": line.cnpj,
            "operadora_original": line.operadora_original,
            "operadora_padronizada": line.operadora_padronizada,
            "plataforma_login": line.plataforma_login,
            "link_portal": line.link_portal,
            "usuario_login": line.usuario_login,
            "senha": "***" if line.senha else "",
            "codigo_cliente": line.codigo_cliente,
            "conta": line.conta,
            "titulo_conta": line.titulo_conta,
            "mes_referencia": line.mes_referencia,
            "vencimento": line.vencimento,
            "nome_arquivo_esperado": line.nome_arquivo_esperado,
            "baixar_pdf": "SIM" if line.baixar_pdf else "NAO",
            "baixar_ae": "SIM" if line.baixar_ae else "NAO",
            "formatos_aceitos_ae": ", ".join(line.formatos_aceitos_ae),
        }

    @staticmethod
    def _validate_line(line: NormalizedInvoiceLine) -> list[str]:
        issues: list[str] = []
        if not line.empresa:
            issues.append("Campo obrigatorio ausente: empresa")
        if not line.operadora_original:
            issues.append("Campo obrigatorio ausente: operadora")
        if not any([line.codigo_cliente, line.conta, line.titulo_conta, line.numero_ticket]):
            issues.append("Falta algum identificador de conta, titulo ou ticket")
        if not any([line.mes_referencia, line.vencimento]):
            issues.append("Falta mes_referencia ou vencimento")
        if not line.usuario_login:
            issues.append("Campo obrigatorio ausente: usuario_login")
        if not line.senha:
            issues.append("Campo obrigatorio ausente: senha")
        if not line.link_portal and normalize_key(line.plataforma_login) in {"", "generic_portal"}:
            issues.append("Falta link_portal ou plataforma reconhecida")
        return issues
