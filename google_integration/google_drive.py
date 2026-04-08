from pathlib import Path
from typing import Any, Dict, Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from drive.gdrive_oauth_api import GDriveUserClient


GOOGLE_GUI_SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]


def get_gui_credentials(client_secret_file: str, token_file: str) -> Credentials:
    """Authenticate the desktop user and cache the OAuth token locally."""
    creds: Optional[Credentials] = None
    token_path = Path(token_file)

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), GOOGLE_GUI_SCOPES)

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())

    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            client_secret_file,
            GOOGLE_GUI_SCOPES,
        )
        creds = flow.run_local_server(port=0, open_browser=True)

    token_path.parent.mkdir(parents=True, exist_ok=True)
    token_path.write_text(creds.to_json(), encoding="utf-8")
    return creds


class GoogleDriveIntegration:
    """Thin GUI wrapper over the existing Drive OAuth client."""

    def __init__(
        self,
        client_secret_file: str,
        token_file: str,
        folder_id: Optional[str] = None,
    ) -> None:
        creds = get_gui_credentials(client_secret_file, token_file)
        client = GDriveUserClient.__new__(GDriveUserClient)
        client.client_secret_file = client_secret_file
        client.token_file = token_file
        client.folder_id = folder_id
        client._port = 0
        client._service = build(
            "drive",
            "v3",
            credentials=creds,
            cache_discovery=False,
        )
        self._client = client

    def create_spreadsheet(self, title: str) -> Dict[str, Any]:
        return self._client.create_google_sheet(title, parent_id=self._client.folder_id)
