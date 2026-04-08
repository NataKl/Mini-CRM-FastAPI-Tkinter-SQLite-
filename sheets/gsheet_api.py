import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
import re

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

try:
    # Optional: allow configuration via .env without hard dependency during import
    from dotenv import load_dotenv

    # Load .env from project root (same dir as this script) if present.
    # Allow already-set env vars to take precedence.
    _env_path = Path(__file__).with_name(".env")
    if _env_path.exists():
        load_dotenv(dotenv_path=_env_path, override=False)
    else:
        load_dotenv(override=False)
except Exception:
    # It's fine if python-dotenv is not installed; env vars can still be provided by the OS
    pass


SHEETS_API_SCOPE = "https://www.googleapis.com/auth/spreadsheets"


class GSheetClient:
    """
    High-level Google Sheets client using a service account.

    Configuration:
    - Provide one of:
      - service_account_file (path to JSON key), OR
      - credentials_dict (parsed JSON key as dict)
    - spreadsheet_id: target Spreadsheet ID

    Notes:
    - All public methods raise HttpError on API failures. Callers may catch for handling.
    - For convenience, methods accept either a sheet title (e.g., \"Sheet1\")
      or full A1 range (e.g., \"Sheet1!A1:C10\").
    """

    def __init__(
        self,
        spreadsheet_id: str,
        service_account_file: Optional[str] = None,
        credentials_dict: Optional[Dict[str, Any]] = None,
        app_name: str = "GSheetClient",
    ) -> None:
        if not spreadsheet_id:
            raise ValueError("spreadsheet_id is required")
        if not service_account_file and not credentials_dict:
            raise ValueError("Provide either service_account_file or credentials_dict")

        self.spreadsheet_id: str = spreadsheet_id
        self._service = self._build_service(
            service_account_file=service_account_file,
            credentials_dict=credentials_dict,
            app_name=app_name,
        )

    @staticmethod
    def _build_service(
        service_account_file: Optional[str],
        credentials_dict: Optional[Dict[str, Any]],
        app_name: str,
    ):
        if service_account_file:
            credentials = service_account.Credentials.from_service_account_file(
                service_account_file, scopes=[SHEETS_API_SCOPE]
            )
        else:
            credentials = service_account.Credentials.from_service_account_info(
                credentials_dict, scopes=[SHEETS_API_SCOPE]
            )
        return build("sheets", "v4", credentials=credentials, cache_discovery=False)

    # -------- Spreadsheet metadata helpers --------
    def get_spreadsheet_metadata(self) -> Dict[str, Any]:
        return (
            self._service.spreadsheets()
            .get(spreadsheetId=self.spreadsheet_id, includeGridData=False)
            .execute()
        )

    def get_sheet_titles(self) -> List[str]:
        metadata = self.get_spreadsheet_metadata()
        sheets = metadata.get("sheets", [])
        return [s["properties"]["title"] for s in sheets if "properties" in s]

    def title_to_sheet_id(self, title: str) -> Optional[int]:
        metadata = self.get_spreadsheet_metadata()
        for sheet in metadata.get("sheets", []):
            props = sheet.get("properties", {})
            if props.get("title") == title:
                return props.get("sheetId")
        return None

    def get_default_sheet_title(self) -> str:
        titles = self.get_sheet_titles()
        if not titles:
            raise ValueError("Spreadsheet has no sheets. Create one first.")
        return titles[0]

    # -------- READ --------
    def read_values(self, a1_range_or_title: Optional[str] = None) -> List[List[Any]]:
        """
        Read values from a given A1 range or entire sheet by title.
        If None, reads all values from the first sheet.
        """
        if a1_range_or_title is None:
            a1_range_or_title = self.get_default_sheet_title()
        result = (
            self._service.spreadsheets()
            .values()
            .get(spreadsheetId=self.spreadsheet_id, range=a1_range_or_title)
            .execute()
        )
        return result.get("values", [])

    # -------- CREATE (append) --------
    def append_rows(
        self,
        a1_range_or_title: str,
        rows: Sequence[Sequence[Any]],
        value_input_option: str = "USER_ENTERED",
    ) -> Dict[str, Any]:
        """
        Append rows to the bottom of the given range or sheet.
        """
        body = {"values": [list(r) for r in rows]}
        return (
            self._service.spreadsheets()
            .values()
            .append(
                spreadsheetId=self.spreadsheet_id,
                range=a1_range_or_title,
                valueInputOption=value_input_option,
                insertDataOption="INSERT_ROWS",
                body=body,
            )
            .execute()
        )

    # -------- UPDATE --------
    def update_range(
        self,
        a1_range: str,
        values: Sequence[Sequence[Any]],
        value_input_option: str = "USER_ENTERED",
    ) -> Dict[str, Any]:
        """
        Update a specific A1 range with provided values (2D list).
        """
        body = {"range": a1_range, "values": [list(r) for r in values], "majorDimension": "ROWS"}
        return (
            self._service.spreadsheets()
            .values()
            .update(
                spreadsheetId=self.spreadsheet_id,
                range=a1_range,
                valueInputOption=value_input_option,
                body=body,
            )
            .execute()
        )

    def batch_update_ranges(
        self,
        data: Sequence[Tuple[str, Sequence[Sequence[Any]]]],
        value_input_option: str = "USER_ENTERED",
    ) -> Dict[str, Any]:
        """
        Batch update multiple A1 ranges.
        data: list of (a1_range, values)
        """
        body = {
            "valueInputOption": value_input_option,
            "data": [{"range": r, "values": [list(row) for row in vals]} for r, vals in data],
        }
        return (
            self._service.spreadsheets()
            .values()
            .batchUpdate(spreadsheetId=self.spreadsheet_id, body=body)
            .execute()
        )

    # -------- DELETE / CLEAR --------
    def clear_range(self, a1_range: str) -> Dict[str, Any]:
        return (
            self._service.spreadsheets()
            .values()
            .clear(spreadsheetId=self.spreadsheet_id, range=a1_range, body={})
            .execute()
        )

    def delete_rows(self, sheet_title: str, start_index: int, end_index: int) -> Dict[str, Any]:
        """
        Delete rows in [start_index, end_index) (0-based, inclusive start, exclusive end).
        """
        sheet_id = self.title_to_sheet_id(sheet_title)
        if sheet_id is None:
            raise ValueError(f"Sheet '{sheet_title}' not found")
        requests = [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_index,
                        "endIndex": end_index,
                    }
                }
            }
        ]
        body = {"requests": requests}
        return (
            self._service.spreadsheets()
            .batchUpdate(spreadsheetId=self.spreadsheet_id, body=body)
            .execute()
        )

    # -------- GENERIC BATCH UPDATE --------
    def apply_requests(self, requests: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Send arbitrary batchUpdate requests to the Sheets API.
        Used for complex formatting, merging cells, resizing rows/columns, etc.

        Example request types:
          - mergeCells
          - repeatCell  (apply formatting to a range)
          - updateDimensionProperties  (column/row sizes)
          - updateSheetProperties  (freeze rows, rename sheet, etc.)
        """
        body = {"requests": requests}
        return (
            self._service.spreadsheets()
            .batchUpdate(spreadsheetId=self.spreadsheet_id, body=body)
            .execute()
        )

    # -------- SHEET MANAGEMENT --------
    def create_sheet(self, title: str, rows: int = 1000, cols: int = 26) -> Dict[str, Any]:
        requests = [
            {
                "addSheet": {
                    "properties": {
                        "title": title,
                        "gridProperties": {"rowCount": rows, "columnCount": cols},
                    }
                }
            }
        ]
        body = {"requests": requests}
        return (
            self._service.spreadsheets()
            .batchUpdate(spreadsheetId=self.spreadsheet_id, body=body)
            .execute()
        )

    def delete_sheet(self, sheet_title: str) -> Dict[str, Any]:
        sheet_id = self.title_to_sheet_id(sheet_title)
        if sheet_id is None:
            raise ValueError(f"Sheet '{sheet_title}' not found")
        requests = [{"deleteSheet": {"sheetId": sheet_id}}]
        body = {"requests": requests}
        return (
            self._service.spreadsheets()
            .batchUpdate(spreadsheetId=self.spreadsheet_id, body=body)
            .execute()
        )


def _load_config_from_env() -> Tuple[str, Optional[str], Optional[Dict[str, Any]]]:
    """
    Load configuration from environment variables.
    - GOOGLE_SPREADSHEET_ID (required)
    - GOOGLE_SERVICE_ACCOUNT_FILE (path) OR GOOGLE_SERVICE_ACCOUNT_JSON (raw JSON string)
    """
    spreadsheet_id = os.getenv("GOOGLE_SPREADSHEET_ID", "").strip()

    # Allow passing a full Google Sheets URL; extract the ID.
    if "docs.google.com/spreadsheets" in spreadsheet_id:
        m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", spreadsheet_id)
        if m:
            spreadsheet_id = m.group(1)

    service_account_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip() or None
    sa_json_raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip() or None
    credentials_dict: Optional[Dict[str, Any]] = None
    if sa_json_raw:
        try:
            import json as _json

            credentials_dict = _json.loads(sa_json_raw)
        except Exception as exc:  # noqa: BLE001
            raise ValueError("Invalid JSON in GOOGLE_SERVICE_ACCOUNT_JSON") from exc

    # Resolve relative paths robustly: try as-is, then relative to script directory.
    if service_account_file:
        candidate = Path(service_account_file).expanduser()
        if not candidate.is_absolute():
            # First, try relative to current working directory
            if not candidate.exists():
                # Then, try relative to the script directory
                script_dir_candidate = Path(__file__).parent.joinpath(service_account_file)
                if script_dir_candidate.exists():
                    candidate = script_dir_candidate
        service_account_file = str(candidate)

    return spreadsheet_id, service_account_file, credentials_dict


if __name__ == "__main__":
    """
    Entry point for quick verification.
    Reads all values from the first (default) sheet and prints them.

    Configuration options:
    - Put service account key file in project root and set in .env:
        GOOGLE_SERVICE_ACCOUNT_FILE=service_account.json
      or provide raw JSON:
        GOOGLE_SERVICE_ACCOUNT_JSON={...}
    - Set target spreadsheet:
        GOOGLE_SPREADSHEET_ID=your_spreadsheet_id_here
    """
    try:
        # Ensure .env is loaded when executing directly
        try:
            from dotenv import load_dotenv as _ld

            _ld(dotenv_path=Path(__file__).with_name(".env"), override=False)
        except Exception:
            pass

        spreadsheet_id, sa_file, sa_creds = _load_config_from_env()
        if not spreadsheet_id:
            raise RuntimeError(
                "GOOGLE_SPREADSHEET_ID is not set. Provide it via environment or .env file."
            )
        if not sa_file and not sa_creds:
            # Fallback: look for a common filename in project root
            candidate = "service_account.json"
            script_dir_candidate = Path(__file__).parent.joinpath(candidate)
            if os.path.exists(candidate):
                sa_file = candidate
            elif script_dir_candidate.exists():
                sa_file = str(script_dir_candidate)
            else:
                # Try to auto-detect a likely service account JSON file in script directory
                try:
                    possible = list(Path(__file__).parent.glob("*.json"))
                    selected: Optional[Path] = None
                    for p in possible:
                        try:
                            import json as _json

                            with p.open("r", encoding="utf-8") as fh:
                                data = _json.load(fh)
                            if isinstance(data, dict) and data.get("type") == "service_account":
                                selected = p
                                break
                        except Exception:
                            continue
                    if selected:
                        sa_file = str(selected)
                    else:
                        raise RuntimeError(
                            "Service account credentials not provided. "
                            "Set GOOGLE_SERVICE_ACCOUNT_FILE or GOOGLE_SERVICE_ACCOUNT_JSON."
                        )
                except Exception as _:
                    raise RuntimeError(
                        "Service account credentials not provided. "
                        "Set GOOGLE_SERVICE_ACCOUNT_FILE or GOOGLE_SERVICE_ACCOUNT_JSON."
                    )

        client = GSheetClient(
            spreadsheet_id=spreadsheet_id,
            service_account_file=sa_file,
            credentials_dict=sa_creds,
        )
        values = client.read_values()  # default sheet
        for r_idx, row in enumerate(values, start=1):
            print(f"{r_idx}: {row}")
    except HttpError as api_err:
        # Provide a clearer hint if the document is not a Google Sheet
        try:
            status = getattr(api_err, "status_code", None) or getattr(api_err, "resp", None).status
        except Exception:
            status = None
        message = str(api_err)
        if status == 400 and "not supported for this document" in message:
            print(
                "Google API error: The provided ID does not refer to a Google Spreadsheet.\n"
                "Make sure you use a Sheets document (docs.google.com/spreadsheets) and pass only the ID "
                "(not the full URL)."
            )
        elif status == 404 and "Requested entity was not found" in message:
            print(
                "Google API error: Spreadsheet not found. Check the ID and that the service account has access."
            )
        else:
            print(f"Google API error: {api_err}")
    except Exception as err:  # noqa: BLE001
        print(f"Error: {err}")

