from __future__ import annotations

import json
import os
from pathlib import Path

from google.oauth2.service_account import Credentials


GOOGLE_SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "google-service-account.json")
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")


def build_google_credentials(scopes: tuple[str, ...] | list[str]) -> Credentials:
    raw_json = GOOGLE_SERVICE_ACCOUNT_JSON.strip()
    if raw_json:
        info = json.loads(raw_json)
        return Credentials.from_service_account_info(info, scopes=scopes)

    credential_path = Path(GOOGLE_SERVICE_ACCOUNT_FILE)
    return Credentials.from_service_account_file(str(credential_path), scopes=scopes)
