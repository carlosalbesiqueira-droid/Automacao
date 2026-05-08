from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

import gspread
from googleapiclient.discovery import build
from gspread.utils import ValidationConditionType

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from bot_faturas_v2.google_auth import build_google_credentials
from bot_faturas_v2.tiflux_sheet_service import FIELD_SPECS, RESULT_HEADERS

GOOGLE_SCOPES = (
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
)

DEFAULT_CREDENTIAL_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "google-service-account.json")
DEFAULT_SPREADSHEET_TITLE = os.getenv("TIFLUX_SHEET_TITLE", "TiFlux Y3 - Preenchimento")
DEFAULT_WORKSHEET_TITLE = os.getenv("TIFLUX_WORKSHEET_TITLE", "TIFLUX_PREENCHIMENTO")
DEFAULT_SHARE_EMAIL = os.getenv("TIFLUX_SHEET_SHARE_EMAIL", "carlos.siqueira@y3gestao.com.br")
CHECKBOX_CELL = "S2"
STATUS_CELL = "T2"
IMPEDIMENTO_OPTIONS = [
    "Problema de Login no Fatura Fácil (Embratel)",
    "Sem Acesso ao site - Cliente SAAS",
    "Fatura não disponível, aberto protocolo",
    "Não disponível – Fatura do mês não está disponível.",
    "Erro no site – Problemas de acesso ao portal. Não está relacionado ...",
    "Sem acesso ao site – Problemas com login ou senha.",
    "N/A",
    "Pendente Opedora - Validando trafego",
    "Emissão atrasada devido nova NFCOM",
    "AE (Arquivo Eletrônico) – Indisponível",
    "Pendente Prorrogação",
    "Fornecedor não disponibiliza portal",
    "Não disponível – Fatura ou CNPJ não localizados nos acessos.",
    "PDF – Indisponível",
    "Erro Upload – Portal",
    "NF – Indisponível",
    "NFPREF – Arquivo RPS não foi convertido em NF PREF",
]
TRATATIVAS_OPTIONS = [
    "Problemas com AE – Contatar a operadora",
    "Solicitar arquivo por e-mail",
    "Solicitar arquivos para o cliente",
    "Em Cancelamento",
    "Solicitar prorrogação",
    "Corrigir inconsistência",
    "Escalado Anatel: Dentro do Prazo",
    "Acompanhar devolutiva – Cliente",
    "Abrir protocolo junto ao fornecedor",
    "N/A",
    "Reprocessar conta",
    "Ajustar Valores",
    "Escalado Ouvidoria: Dentro do Prazo",
    "Direcionado ao desenvolvedor",
    "Validação de Cancelamento",
    "Validação de Ausência de trafego",
    "Validação de bloqueio/suspensão temporária",
    "Avaliar possível alteração de vencimento",
    "NF PREF – Acompanhar conversão para NFPREF",
]
ESTAGIO_OPTIONS = [
    "Pendente",
    "Em andamento",
    "Concluído",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Cria ou atualiza a planilha modelo do TiFlux.")
    parser.add_argument("--title", default=DEFAULT_SPREADSHEET_TITLE, help="Titulo da planilha Google.")
    parser.add_argument("--worksheet", default=DEFAULT_WORKSHEET_TITLE, help="Nome da aba principal.")
    parser.add_argument("--share-email", default=DEFAULT_SHARE_EMAIL, help="E-mail para compartilhar a planilha.")
    parser.add_argument(
        "--credentials-file",
        default=DEFAULT_CREDENTIAL_FILE,
        help="Arquivo JSON da service account Google.",
    )
    parser.add_argument(
        "--spreadsheet-id",
        default="",
        help="Se informado, reutiliza a planilha existente em vez de criar uma nova.",
    )
    parser.add_argument(
        "--prune-extra-tabs",
        action="store_true",
        help="Remove abas extras, mantendo apenas a aba principal e a aba oculta de listas.",
    )
    return parser.parse_args()


def load_clients(credentials_file: str) -> tuple[gspread.Client, object]:
    if credentials_file and credentials_file != DEFAULT_CREDENTIAL_FILE:
        os.environ["GOOGLE_SERVICE_ACCOUNT_FILE"] = credentials_file
    creds = build_google_credentials(GOOGLE_SCOPES)
    client = gspread.authorize(creds)
    sheets_service = build("sheets", "v4", credentials=creds, cache_discovery=False)
    return client, sheets_service


def desired_headers() -> list[str]:
    headers = ["NUMERO_TICKET"]
    headers.extend(spec.label for spec in FIELD_SPECS)
    headers.extend(RESULT_HEADERS)
    return headers


def ensure_main_worksheet(spreadsheet: gspread.Spreadsheet, worksheet_title: str) -> gspread.Worksheet:
    try:
        worksheet = spreadsheet.worksheet(worksheet_title)
        worksheet.clear()
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=worksheet_title, rows=2000, cols=30)

    headers = desired_headers()
    required_cols = max(len(headers), 21)
    current_cols = int(getattr(worksheet, "col_count", required_cols) or required_cols)
    if current_cols < required_cols:
        worksheet.add_cols(required_cols - current_cols)
    worksheet.update(range_name="A1", values=[headers], value_input_option="USER_ENTERED")
    worksheet.freeze(rows=1)
    worksheet.columns_auto_resize(0, len(headers) - 1)
    worksheet.update(
        range_name="S1:U3",
        values=[
            ["ACAO", "STATUS_BOTAO", "ULTIMO_JOB"],
            [False, "Aguardando", ""],
            ["Marque a caixa S2 para executar", "", ""],
        ],
        value_input_option="USER_ENTERED",
    )
    worksheet.format("S1:U3", {"textFormat": {"bold": True}})
    worksheet.add_validation(
        CHECKBOX_CELL,
        ValidationConditionType.boolean,
        [],
        strict=True,
        showCustomUi=True,
        inputMessage="Marque para executar a automacao.",
    )
    worksheet.update_acell(CHECKBOX_CELL, "FALSE")
    worksheet.update_note(
        CHECKBOX_CELL,
        "Marque esta caixa para disparar a automacao da aba TIFLUX_PREENCHIMENTO via Apps Script.",
    )
    return worksheet


def apply_data_validations(main_ws: gspread.Worksheet) -> None:
    main_ws.add_validation(
        "C2:C2000",
        ValidationConditionType.one_of_list,
        IMPEDIMENTO_OPTIONS,
        strict=False,
        showCustomUi=True,
        inputMessage="Selecione ou digite o impedimento desejado.",
    )
    main_ws.add_validation(
        "D2:D2000",
        ValidationConditionType.one_of_list,
        TRATATIVAS_OPTIONS,
        strict=False,
        showCustomUi=True,
        inputMessage="Selecione ou digite a tratativa desejada.",
    )
    main_ws.add_validation(
        "M2:M2000",
        ValidationConditionType.one_of_list,
        ESTAGIO_OPTIONS,
        strict=False,
        showCustomUi=True,
        inputMessage="Selecione ou digite o estágio desejado.",
    )


def prune_extra_tabs(spreadsheet: gspread.Spreadsheet, keep_titles: set[str]) -> None:
    for worksheet in spreadsheet.worksheets():
        if worksheet.title in keep_titles:
            continue
        spreadsheet.del_worksheet(worksheet)


def format_sheet(sheets_service: object, spreadsheet_id: str, main_ws: gspread.Worksheet) -> None:
    main_id = main_ws.id
    header_count = len(desired_headers())
    requests = [
        {
            "repeatCell": {
                "range": {"sheetId": main_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": header_count},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.85, "green": 0.93, "blue": 0.87},
                        "textFormat": {"bold": True},
                        "horizontalAlignment": "CENTER",
                        "wrapStrategy": "WRAP",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)",
            }
        },
        {
            "updateSheetProperties": {
                "properties": {"sheetId": main_id, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount",
            }
        },
        {
            "repeatCell": {
                "range": {"sheetId": main_id, "startRowIndex": 0, "endRowIndex": 3, "startColumnIndex": 18, "endColumnIndex": 21},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.98, "green": 0.93, "blue": 0.83},
                        "textFormat": {"bold": True},
                        "horizontalAlignment": "CENTER",
                        "wrapStrategy": "WRAP",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)",
            }
        },
    ]
    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()


def create_or_open_spreadsheet(client: gspread.Client, args: argparse.Namespace) -> gspread.Spreadsheet:
    if args.spreadsheet_id:
        return client.open_by_key(args.spreadsheet_id)
    return client.create(args.title)


def share_spreadsheet(spreadsheet: gspread.Spreadsheet, share_email: str) -> None:
    if not share_email or share_email.strip() in {"-", "none", "null"}:
        return
    spreadsheet.share(share_email, perm_type="user", role="writer", notify=False)


def load_apps_script_template() -> str:
    script_path = Path(__file__).resolve().parent / "google_apps_script_tiflux.gs"
    return script_path.read_text(encoding="utf-8")


def main() -> None:
    args = parse_args()
    client, sheets_service = load_clients(args.credentials_file)
    spreadsheet = create_or_open_spreadsheet(client, args)
    share_spreadsheet(spreadsheet, args.share_email)

    main_ws = ensure_main_worksheet(spreadsheet, args.worksheet)
    apply_data_validations(main_ws)
    format_sheet(sheets_service, spreadsheet.id, main_ws)
    if args.prune_extra_tabs:
        prune_extra_tabs(spreadsheet, {args.worksheet})

    payload = {
        "spreadsheet_id": spreadsheet.id,
        "spreadsheet_title": spreadsheet.title,
        "spreadsheet_url": f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}/edit",
        "worksheet_name": args.worksheet,
        "share_email": args.share_email,
        "headers": desired_headers(),
        "button_checkbox_cell": CHECKBOX_CELL,
        "status_cell": STATUS_CELL,
        "apps_script_template": load_apps_script_template(),
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
