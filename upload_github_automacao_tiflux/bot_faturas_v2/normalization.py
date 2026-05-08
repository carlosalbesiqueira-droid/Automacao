from __future__ import annotations

import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse
from zoneinfo import ZoneInfo

from rapidfuzz import fuzz

from .constants import DEFAULT_AE_FORMATS


DEFAULT_PLATFORM_URLS = {
    "MEU_TIM": "https://meutim.tim.com.br/",
    "MVE_VIVO": "https://mve.vivo.com.br/",
    "VIVO_EM_DIA": "https://www.vivoemdia.com.br/",
    "CLARO_CONTA_ONLINE": "https://minhaclaro.claro.com.br/",
    "CLARO_OPERA360": "https://empresa.claro.com.br/",
    "CLARO_PRESTO360": "https://empresa.claro.com.br/",
    "OI_SOLUCOES": "https://oicontasb2b.com.br/",
}


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_key(value: object) -> str:
    text = unicodedata.normalize("NFKD", clean_text(value))
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"[^a-zA-Z0-9]+", "_", text).strip("_").lower()
    return text


def ensure_directory(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def safe_filename(value: object, fallback: str = "arquivo") -> str:
    text = clean_text(value)
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", text)
    text = re.sub(r"\s+", "_", text).strip("._ ")
    return text or fallback


def slugify(value: object, fallback: str = "item") -> str:
    text = normalize_key(value).replace("_", "-").strip("-")
    return text or fallback


def now_local(timezone_name: str) -> str:
    return datetime.now(ZoneInfo(timezone_name)).replace(microsecond=0).isoformat()


def parse_bool(value: object, default: bool = False) -> bool:
    text = normalize_key(value)
    if not text:
        return default
    return text in {"sim", "s", "true", "1", "yes", "y", "baixar"}


def normalize_cnpj(value: object) -> str:
    digits = re.sub(r"\D+", "", clean_text(value))
    return digits if len(digits) == 14 else ""


def normalize_url(value: object) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if not re.match(r"^[a-z][a-z0-9+.-]*://", text, flags=re.IGNORECASE):
        text = f"https://{text.lstrip('/')}"
    return text


def extract_domain(value: object) -> str:
    url = normalize_url(value)
    if not url:
        return ""
    return urlparse(url).netloc.lower().removeprefix("www.")


def mask_secret(value: object) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if len(text) <= 3:
        return "***"
    return f"{text[:2]}***{text[-1:]}"


def dump_json(payload: object) -> str:
    return json.dumps(payload, ensure_ascii=False, indent=2)


def parse_reference_month(value: object) -> str:
    text = clean_text(value)
    if not text:
        return ""
    digits = re.sub(r"[^\d]", "", text)
    if re.fullmatch(r"\d{6}", digits):
        return f"{digits[:4]}-{digits[4:6]}"
    if re.fullmatch(r"\d{2}\d{4}", digits):
        return f"{digits[2:6]}-{digits[:2]}"
    match = re.search(r"(20\d{2})[-/](0[1-9]|1[0-2])", text)
    if match:
        return f"{match.group(1)}-{match.group(2)}"
    match = re.search(r"(0[1-9]|1[0-2])[-/](20\d{2})", text)
    if match:
        return f"{match.group(2)}-{match.group(1)}"
    return text


def parse_date(value: object) -> str:
    text = clean_text(value)
    if not text:
        return ""
    for pattern in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, pattern).date().isoformat()
        except ValueError:
            continue
    return text


def normalize_accepted_ae_formats(value: object) -> tuple[str, ...]:
    text = clean_text(value).upper()
    if not text:
        return DEFAULT_AE_FORMATS
    parts = [item.strip(" .") for item in re.split(r"[,;/| ]+", text) if item.strip(" .")]
    allowed = [item for item in parts if item in {"AE", "TXT", "ZIP", "CSV"}]
    return tuple(dict.fromkeys(allowed or list(DEFAULT_AE_FORMATS)))


def standardize_operator(raw_value: object) -> str:
    text = normalize_key(raw_value).replace("_", " ").strip()
    if not text:
        return ""
    aliases = {
        "tim": "TIM",
        "tim celular": "TIM",
        "vivo": "VIVO",
        "mve vivo": "VIVO",
        "vivo em dia": "VIVO",
        "claro": "CLARO",
        "embratel": "EMBRATEL",
        "oi": "OI",
        "oi solucoes": "OI",
        "algar": "ALGAR",
        "hubsoft": "HUBSOFT",
        "ixc": "IXC",
    }
    if text in aliases:
        return aliases[text]
    return text.replace(" ", "_").upper()


def infer_platform(
    raw_platform: object,
    link_portal: object,
    operadora_padronizada: object,
) -> str:
    platform = normalize_key(raw_platform)
    domain = extract_domain(link_portal)
    operator = clean_text(operadora_padronizada).upper()

    domain_map = {
        "meutim.tim.com.br": "MEU_TIM",
        "meu.tim.com.br": "MEU_TIM",
        "mve.vivo.com.br": "MVE_VIVO",
        "vivoemdia.com.br": "VIVO_EM_DIA",
        "minhaclaro.claro.com.br": "CLARO_CONTA_ONLINE",
        "empresa.claro.com.br": "CLARO_OPERA360",
        "oicontasb2b.com.br": "OI_SOLUCOES",
    }
    if domain in domain_map:
        return domain_map[domain]

    explicit_map = {
        "meu_tim": "MEU_TIM",
        "mve_vivo": "MVE_VIVO",
        "vivo_em_dia": "VIVO_EM_DIA",
        "claro_conta_online": "CLARO_CONTA_ONLINE",
        "claro_opera360": "CLARO_OPERA360",
        "claro_presto360": "CLARO_PRESTO360",
        "hubsoft": "HUBSOFT",
        "ixc": "IXC",
        "oi_solucoes": "OI_SOLUCOES",
        "algar_portal": "ALGAR_PORTAL",
        "generic_portal": "GENERIC_PORTAL",
    }
    if platform in explicit_map:
        return explicit_map[platform]

    operator_map = {
        "TIM": "MEU_TIM",
        "VIVO": "MVE_VIVO",
        "CLARO": "CLARO_CONTA_ONLINE",
        "EMBRATEL": "GENERIC_PORTAL",
        "OI": "OI_SOLUCOES",
        "ALGAR": "ALGAR_PORTAL",
        "HUBSOFT": "HUBSOFT",
        "IXC": "IXC",
    }
    return operator_map.get(operator, "GENERIC_PORTAL")


def default_portal_url(plataforma_login: object) -> str:
    return DEFAULT_PLATFORM_URLS.get(clean_text(plataforma_login).upper(), "")


def header_score(header: str, candidate: str) -> int:
    normalized_header = normalize_key(header)
    normalized_candidate = normalize_key(candidate)
    if normalized_header == normalized_candidate:
        return 100
    if normalized_header in normalized_candidate or normalized_candidate in normalized_header:
        return 95
    return int(fuzz.token_sort_ratio(normalized_header, normalized_candidate))
