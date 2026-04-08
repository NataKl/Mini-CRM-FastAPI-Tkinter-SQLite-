import io
import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

try:
    from dotenv import load_dotenv

    _env_path = Path(__file__).with_name(".env")
    if _env_path.exists():
        load_dotenv(dotenv_path=_env_path, override=False)
    else:
        load_dotenv(override=False)
except Exception:
    pass


DRIVE_API_SCOPES = ["https://www.googleapis.com/auth/drive"]

# MIME-типы Google Workspace для экспорта
GOOGLE_MIME_TYPES = {
    "application/vnd.google-apps.spreadsheet",
    "application/vnd.google-apps.document",
    "application/vnd.google-apps.presentation",
    "application/vnd.google-apps.drawing",
}


class GDriveClient:
    """
    High-level Google Drive client using a service account.

    Supports CRUD operations on files and folders:
      - CREATE  : create_folder, upload_file, create_google_sheet, create_google_doc
      - READ    : list_files, get_file_metadata, download_file, export_google_doc
      - UPDATE  : rename_file, move_file, update_file_content, update_file_metadata
      - DELETE  : delete_file, trash_file

    Configuration:
    - Provide one of:
      - service_account_file (path to JSON key), OR
      - credentials_dict (parsed JSON key as dict)
    - folder_id (optional): default working folder; used when parent_id is not specified.

    Notes:
    - All public methods raise HttpError on API failures. Callers may catch for handling.
    - The service account must be granted access to the target folder/file in Google Drive.
    """

    def __init__(
        self,
        service_account_file: Optional[str] = None,
        credentials_dict: Optional[Dict[str, Any]] = None,
        folder_id: Optional[str] = None,
        app_name: str = "GDriveClient",
    ) -> None:
        if not service_account_file and not credentials_dict:
            raise ValueError("Provide either service_account_file or credentials_dict")

        # Default working folder (can be overridden per-call)
        self.folder_id: Optional[str] = folder_id
        self._service = self._build_service(
            service_account_file=service_account_file,
            credentials_dict=credentials_dict,
            app_name=app_name,
        )

    # ------------------------------------------------------------------ #
    #  Internal helpers                                                    #
    # ------------------------------------------------------------------ #

    @staticmethod
    def _build_service(
        service_account_file: Optional[str],
        credentials_dict: Optional[Dict[str, Any]],
        app_name: str,
    ):
        if service_account_file:
            credentials = service_account.Credentials.from_service_account_file(
                service_account_file, scopes=DRIVE_API_SCOPES
            )
        else:
            credentials = service_account.Credentials.from_service_account_info(
                credentials_dict, scopes=DRIVE_API_SCOPES
            )
        return build("drive", "v3", credentials=credentials, cache_discovery=False)

    def _resolve_parent(self, parent_id: Optional[str]) -> Optional[str]:
        """Return explicit parent_id, falling back to default folder_id."""
        return parent_id if parent_id is not None else self.folder_id

    # ------------------------------------------------------------------ #
    #  READ                                                                #
    # ------------------------------------------------------------------ #

    def list_files(
        self,
        folder_id: Optional[str] = None,
        query: Optional[str] = None,
        page_size: int = 100,
        order_by: str = "modifiedTime desc",
        fields: str = "nextPageToken, files(id, name, mimeType, size, modifiedTime, parents, trashed)",
    ) -> List[Dict[str, Any]]:
        """
        List files visible to the service account.

        Args:
            folder_id: Limit results to a specific folder. Falls back to self.folder_id.
                       Pass an empty string ``""`` to list *all* files without folder filter.
            query: Additional Drive query string (q parameter).
            page_size: Max items per page (1-1000).
            order_by: Comma-separated sort fields.
            fields: API field mask.

        Returns:
            List of file resource dicts.
        """
        resolved_folder = self._resolve_parent(folder_id)

        parts: List[str] = ["trashed = false"]
        if resolved_folder:
            parts.append(f"'{resolved_folder}' in parents")
        if query:
            parts.append(query)

        q = " and ".join(parts)

        all_files: List[Dict[str, Any]] = []
        page_token: Optional[str] = None

        while True:
            kwargs: Dict[str, Any] = dict(
                q=q,
                pageSize=min(page_size, 1000),
                orderBy=order_by,
                fields=fields,
            )
            if page_token:
                kwargs["pageToken"] = page_token

            response = self._service.files().list(**kwargs).execute()
            all_files.extend(response.get("files", []))

            page_token = response.get("nextPageToken")
            if not page_token:
                break

        return all_files

    def get_file_metadata(
        self,
        file_id: str,
        fields: str = "id, name, mimeType, size, modifiedTime, parents, trashed",
    ) -> Dict[str, Any]:
        """Return metadata for a single file."""
        return self._service.files().get(fileId=file_id, fields=fields).execute()

    def download_file(self, file_id: str, dest_path: str) -> str:
        """
        Download a binary file (non-Google-Workspace) to *dest_path*.

        Returns the final destination path.
        """
        request = self._service.files().get_media(fileId=file_id)
        dest = Path(dest_path)
        dest.parent.mkdir(parents=True, exist_ok=True)

        with dest.open("wb") as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()

        return str(dest)

    def export_google_doc(
        self,
        file_id: str,
        export_mime_type: str,
        dest_path: str,
    ) -> str:
        """
        Export a Google Workspace document (Docs, Sheets, Slides …) to a
        standard MIME type (e.g. "application/pdf",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet").

        Returns the final destination path.
        """
        request = self._service.files().export_media(
            fileId=file_id, mimeType=export_mime_type
        )
        dest = Path(dest_path)
        dest.parent.mkdir(parents=True, exist_ok=True)

        with dest.open("wb") as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()

        return str(dest)

    # ------------------------------------------------------------------ #
    #  CREATE                                                              #
    # ------------------------------------------------------------------ #

    def create_folder(
        self,
        name: str,
        parent_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Create a folder in Google Drive.

        Returns the created folder resource (id, name, mimeType, …).
        """
        metadata: Dict[str, Any] = {
            "name": name,
            "mimeType": "application/vnd.google-apps.folder",
        }
        resolved = self._resolve_parent(parent_id)
        if resolved:
            metadata["parents"] = [resolved]

        return (
            self._service.files()
            .create(body=metadata, fields="id, name, mimeType, parents")
            .execute()
        )

    def upload_file(
        self,
        local_path: str,
        name: Optional[str] = None,
        parent_id: Optional[str] = None,
        mime_type: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Upload a local file to Google Drive.

        Args:
            local_path: Path to the local file.
            name: Desired name in Drive; defaults to the local file name.
            parent_id: Target folder. Falls back to self.folder_id.
            mime_type: Content MIME type; auto-detected if omitted.

        Returns:
            Created file resource (id, name, mimeType, parents).
        """
        src = Path(local_path)
        file_name = name or src.name

        metadata: Dict[str, Any] = {"name": file_name}
        resolved = self._resolve_parent(parent_id)
        if resolved:
            metadata["parents"] = [resolved]

        media = MediaFileUpload(str(src), mimetype=mime_type, resumable=True)
        return (
            self._service.files()
            .create(body=metadata, media_body=media, fields="id, name, mimeType, size, parents")
            .execute()
        )

    def create_google_sheet(
        self,
        name: str,
        parent_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Create an empty Google Sheets spreadsheet."""
        metadata: Dict[str, Any] = {
            "name": name,
            "mimeType": "application/vnd.google-apps.spreadsheet",
        }
        resolved = self._resolve_parent(parent_id)
        if resolved:
            metadata["parents"] = [resolved]

        return (
            self._service.files()
            .create(body=metadata, fields="id, name, mimeType, parents")
            .execute()
        )

    def create_google_doc(
        self,
        name: str,
        parent_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Create an empty Google Docs document."""
        metadata: Dict[str, Any] = {
            "name": name,
            "mimeType": "application/vnd.google-apps.document",
        }
        resolved = self._resolve_parent(parent_id)
        if resolved:
            metadata["parents"] = [resolved]

        return (
            self._service.files()
            .create(body=metadata, fields="id, name, mimeType, parents")
            .execute()
        )

    # ------------------------------------------------------------------ #
    #  UPDATE                                                              #
    # ------------------------------------------------------------------ #

    def rename_file(self, file_id: str, new_name: str) -> Dict[str, Any]:
        """Rename a file or folder."""
        return (
            self._service.files()
            .update(fileId=file_id, body={"name": new_name}, fields="id, name")
            .execute()
        )

    def move_file(
        self,
        file_id: str,
        new_parent_id: str,
        remove_from_parents: bool = True,
    ) -> Dict[str, Any]:
        """
        Move a file to *new_parent_id*.

        When *remove_from_parents* is True the file is removed from all its
        current parents so that it appears only in the new location.
        """
        remove_parents: Optional[str] = None
        if remove_from_parents:
            meta = self.get_file_metadata(file_id, fields="parents")
            current_parents = meta.get("parents", [])
            remove_parents = ",".join(current_parents)

        kwargs: Dict[str, Any] = dict(
            fileId=file_id,
            addParents=new_parent_id,
            fields="id, name, parents",
        )
        if remove_parents:
            kwargs["removeParents"] = remove_parents

        return self._service.files().update(**kwargs).execute()

    def update_file_content(
        self,
        file_id: str,
        local_path: str,
        mime_type: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Replace the content of an existing Drive file with a local file."""
        media = MediaFileUpload(local_path, mimetype=mime_type, resumable=True)
        return (
            self._service.files()
            .update(fileId=file_id, media_body=media, fields="id, name, mimeType, modifiedTime")
            .execute()
        )

    def update_file_metadata(
        self,
        file_id: str,
        metadata: Dict[str, Any],
    ) -> Dict[str, Any]:
        """
        Update arbitrary metadata fields of a file (name, description, …).

        Example::

            client.update_file_metadata(file_id, {"description": "updated"})
        """
        return (
            self._service.files()
            .update(
                fileId=file_id,
                body=metadata,
                fields="id, name, mimeType, modifiedTime",
            )
            .execute()
        )

    # ------------------------------------------------------------------ #
    #  DELETE                                                              #
    # ------------------------------------------------------------------ #

    def delete_file(self, file_id: str) -> None:
        """
        Permanently delete a file (bypasses Trash).

        Warning: This action is irreversible.
        """
        self._service.files().delete(fileId=file_id).execute()

    def trash_file(self, file_id: str) -> Dict[str, Any]:
        """Move a file to Trash (recoverable)."""
        return (
            self._service.files()
            .update(fileId=file_id, body={"trashed": True}, fields="id, name, trashed")
            .execute()
        )


# ---------------------------------------------------------------------- #
#  Config loader (mirrors pattern from gsheet_api.py)                    #
# ---------------------------------------------------------------------- #

def _load_config_from_env() -> Tuple[Optional[str], Optional[str], Optional[Dict[str, Any]]]:
    """
    Load credentials from environment variables.

    Variables read:
    - GOOGLE_DRIVE_FOLDER_ID        (optional) default working folder
    - GOOGLE_SERVICE_ACCOUNT_FILE   path to JSON key file
    - GOOGLE_SERVICE_ACCOUNT_JSON   raw JSON string of the key
    """
    folder_id = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "").strip() or None

    service_account_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip() or None
    sa_json_raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip() or None
    credentials_dict: Optional[Dict[str, Any]] = None

    if sa_json_raw:
        try:
            credentials_dict = json.loads(sa_json_raw)
        except Exception as exc:
            raise ValueError("Invalid JSON in GOOGLE_SERVICE_ACCOUNT_JSON") from exc

    # Resolve relative paths (same logic as gsheet_api.py)
    if service_account_file:
        candidate = Path(service_account_file).expanduser()
        if not candidate.is_absolute() and not candidate.exists():
            script_dir_candidate = Path(__file__).parent / service_account_file
            if script_dir_candidate.exists():
                candidate = script_dir_candidate
        service_account_file = str(candidate)

    return folder_id, service_account_file, credentials_dict


# ---------------------------------------------------------------------- #
#  Entry point — test run: list all files                                 #
# ---------------------------------------------------------------------- #

if __name__ == "__main__":
    """
    Quick smoke-test: list all files visible to the service account.

    Configuration (via .env or OS environment):
      GOOGLE_SERVICE_ACCOUNT_FILE=excel-factory-492214-34f17627629c.json
      # or:
      GOOGLE_SERVICE_ACCOUNT_JSON={...raw JSON...}
      # optionally limit output to a specific folder:
      GOOGLE_DRIVE_FOLDER_ID=<folder_id>
    """
    import sys

    # Fix Windows console Unicode output
    if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    try:
        # Reload .env when executed directly
        try:
            from dotenv import load_dotenv as _ld

            _ld(dotenv_path=Path(__file__).with_name(".env"), override=False)
        except Exception:
            pass

        folder_id, sa_file, sa_creds = _load_config_from_env()

        # Auto-detect service account JSON in project directory if not configured
        if not sa_file and not sa_creds:
            possible = list(Path(__file__).parent.glob("*.json"))
            for p in possible:
                try:
                    with p.open("r", encoding="utf-8") as fh:
                        data = json.load(fh)
                    if isinstance(data, dict) and data.get("type") == "service_account":
                        sa_file = str(p)
                        print(f"[auto-detect] Using service account file: {p.name}")
                        break
                except Exception:
                    continue

        if not sa_file and not sa_creds:
            raise RuntimeError(
                "Service account credentials not found.\n"
                "Set GOOGLE_SERVICE_ACCOUNT_FILE or GOOGLE_SERVICE_ACCOUNT_JSON "
                "in .env or the OS environment."
            )

        client = GDriveClient(
            service_account_file=sa_file,
            credentials_dict=sa_creds,
            folder_id=folder_id,
        )

        print("=" * 60)
        label = f"folder '{folder_id}'" if folder_id else "all accessible files"
        print(f"Listing {label} …")
        print("=" * 60)

        files = client.list_files(folder_id="1K8bkCd-9YG3yQ7ZcncwgFJfOamjL3St7")

        if not files:
            print("(no files found — make sure the service account has access to the folder)")
        else:
            print(f"{'#':<4} {'ID':<45} {'MIME':<55} {'NAME'}")
            print("-" * 160)
            for idx, f in enumerate(files, start=1):
                size_str = ""
                raw_size = f.get("size")
                if raw_size is not None:
                    kb = int(raw_size) / 1024
                    size_str = f" ({kb:.1f} KB)"
                print(
                    f"{idx:<4} {f.get('id', ''):<45} "
                    f"{f.get('mimeType', ''):<55} "
                    f"{f.get('name', '')}{size_str}"
                )
            print("-" * 160)
            print(f"Total: {len(files)} file(s)")

    except HttpError as api_err:
        try:
            status = (
                getattr(api_err, "status_code", None)
                or getattr(api_err, "resp", None).status
            )
        except Exception:
            status = None
        if status == 403:
            print(
                "Google API error 403: Access denied.\n"
                "Make sure the service account has been granted access to the Drive folder."
            )
        elif status == 404:
            print("Google API error 404: Resource not found. Check the folder ID.")
        else:
            print(f"Google API error: {api_err}")
    except Exception as err:
        print(f"Error: {err}")
