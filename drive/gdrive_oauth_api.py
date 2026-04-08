"""
gdrive_oauth_api.py
-------------------
Google Drive CRUD client that authenticates as a **real user** via OAuth 2.0
(Installed-App flow).  Files created through this client appear in the user's
own Drive — not in a service-account Drive.

OAuth flow (first run only):
  1. A browser window opens asking the user to log in and grant permissions.
  2. The granted tokens are saved to *token_file* (default: ``token.json``).
  3. On subsequent runs the saved token is reused / silently refreshed.

Configuration (via .env or OS environment):
  GOOGLE_OAUTH_CLIENT_SECRET_FILE  path to client_secret_*.json
  GOOGLE_DRIVE_FOLDER_ID           (optional) default working folder ID
  GOOGLE_OAUTH_TOKEN_FILE          (optional) where to save/load the token
                                   (default: token.json in project root)
"""

import json
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

try:
    from dotenv import load_dotenv

    _env_path = Path(__file__).with_name(".env")
    load_dotenv(dotenv_path=_env_path if _env_path.exists() else None, override=False)
except Exception:
    pass


# Scopes — full Drive access so the user can read/write/delete their own files.
# Narrow to drive.file if you only need files created by this app.
OAUTH_SCOPES = ["https://www.googleapis.com/auth/drive"]


class GDriveUserClient:
    """
    Google Drive client authenticated as a personal Google account via OAuth 2.0.

    All files/folders created by this client belong to the authenticated user
    and appear in their Google Drive — unlike the service-account client where
    files live in the service account's isolated storage.

    Args:
        client_secret_file: Path to the ``client_secret_*.json`` downloaded
            from Google Cloud Console (Desktop / Installed App type).
        token_file: Where to cache the OAuth token between runs.
            Defaults to ``token.json`` next to this script.
        folder_id: Default working folder ID used when *parent_id* is not
            specified in individual calls.
        port: Local port for the OAuth redirect (default 0 = pick free port).

    CRUD surface
    ────────────
    CREATE  create_folder · create_google_sheet · create_google_doc · upload_file
    READ    list_files · get_file_metadata · download_file · export_google_doc
    UPDATE  rename_file · move_file · update_file_content · update_file_metadata
    DELETE  delete_file · trash_file
    """

    def __init__(
        self,
        client_secret_file: str,
        token_file: Optional[str] = None,
        folder_id: Optional[str] = None,
        port: int = 0,
    ) -> None:
        self.client_secret_file: str = client_secret_file
        self.token_file: str = token_file or str(
            Path(__file__).with_name("token.json")
        )
        self.folder_id: Optional[str] = folder_id
        self._port: int = port
        self._service = self._build_service()

    # ------------------------------------------------------------------ #
    #  Auth helpers                                                        #
    # ------------------------------------------------------------------ #

    def _get_credentials(self) -> Credentials:
        """Load, refresh, or obtain new OAuth2 credentials."""
        creds: Optional[Credentials] = None
        token_path = Path(self.token_file)

        # 1. Try loading cached token
        if token_path.exists():
            creds = Credentials.from_authorized_user_file(str(token_path), OAUTH_SCOPES)

        # 2. Refresh if expired
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())

        # 3. Run interactive browser flow if no valid token
        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file(
                self.client_secret_file, OAUTH_SCOPES
            )
            creds = flow.run_local_server(port=self._port, open_browser=True)

        # 4. Persist token for next run
        token_path.parent.mkdir(parents=True, exist_ok=True)
        with token_path.open("w", encoding="utf-8") as fh:
            fh.write(creds.to_json())

        return creds

    def _build_service(self):
        creds = self._get_credentials()
        return build("drive", "v3", credentials=creds, cache_discovery=False)

    def _resolve_parent(self, parent_id: Optional[str]) -> Optional[str]:
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
        fields: str = (
            "nextPageToken, "
            "files(id, name, mimeType, size, modifiedTime, parents, trashed)"
        ),
    ) -> List[Dict[str, Any]]:
        """
        List files in the user's Drive.

        Args:
            folder_id: Restrict to a specific folder. Pass ``""`` to skip
                       folder filter and list all Drive files.
            query: Extra Drive query clause (appended with ``and``).
            page_size: Results per page (1–1000).
        """
        resolved = self._resolve_parent(folder_id)
        parts: List[str] = ["trashed = false"]
        if resolved:
            parts.append(f"'{resolved}' in parents")
        if query:
            parts.append(query)

        all_files: List[Dict[str, Any]] = []
        page_token: Optional[str] = None

        while True:
            kwargs: Dict[str, Any] = dict(
                q=" and ".join(parts),
                pageSize=min(page_size, 1000),
                orderBy=order_by,
                fields=fields,
            )
            if page_token:
                kwargs["pageToken"] = page_token

            resp = self._service.files().list(**kwargs).execute()
            all_files.extend(resp.get("files", []))
            page_token = resp.get("nextPageToken")
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
        """Download a binary (non-Workspace) file to *dest_path*."""
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
        Export a Google Workspace file to a standard format.

        Common *export_mime_type* values:
          - ``"application/pdf"``
          - ``"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"`` (xlsx)
          - ``"application/vnd.openxmlformats-officedocument.wordprocessingml.document"`` (docx)
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

    def _create_workspace_file(
        self, name: str, mime_type: str, parent_id: Optional[str]
    ) -> Dict[str, Any]:
        """Internal helper — create a Google Workspace file with optional parent."""
        metadata: Dict[str, Any] = {"name": name, "mimeType": mime_type}
        resolved = self._resolve_parent(parent_id)
        if resolved:
            metadata["parents"] = [resolved]
        return (
            self._service.files()
            .create(body=metadata, fields="id, name, mimeType, parents, webViewLink")
            .execute()
        )

    def create_google_sheet(
        self,
        name: str,
        parent_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Create an empty Google Sheets spreadsheet **in the user's Drive**.

        Args:
            name: Title of the new spreadsheet.
            parent_id: Target folder ID. Falls back to ``self.folder_id``.

        Returns:
            File resource dict with ``id``, ``name``, ``mimeType``,
            ``parents``, ``webViewLink``.
        """
        return self._create_workspace_file(
            name,
            "application/vnd.google-apps.spreadsheet",
            parent_id,
        )

    def create_google_doc(
        self,
        name: str,
        parent_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Create an empty Google Docs document **in the user's Drive**.

        Args:
            name: Title of the new document.
            parent_id: Target folder ID. Falls back to ``self.folder_id``.

        Returns:
            File resource dict with ``id``, ``name``, ``mimeType``,
            ``parents``, ``webViewLink``.
        """
        return self._create_workspace_file(
            name,
            "application/vnd.google-apps.document",
            parent_id,
        )

    def create_folder(
        self,
        name: str,
        parent_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Create a folder in the user's Drive."""
        return self._create_workspace_file(
            name,
            "application/vnd.google-apps.folder",
            parent_id,
        )

    def upload_file(
        self,
        local_path: str,
        name: Optional[str] = None,
        parent_id: Optional[str] = None,
        mime_type: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Upload a local file to the user's Drive.

        Args:
            local_path: Path to the local file.
            name: Name in Drive; defaults to the local filename.
            parent_id: Target folder. Falls back to ``self.folder_id``.
            mime_type: Content MIME type; auto-detected if omitted.
        """
        src = Path(local_path)
        metadata: Dict[str, Any] = {"name": name or src.name}
        resolved = self._resolve_parent(parent_id)
        if resolved:
            metadata["parents"] = [resolved]
        media = MediaFileUpload(str(src), mimetype=mime_type, resumable=True)
        return (
            self._service.files()
            .create(
                body=metadata,
                media_body=media,
                fields="id, name, mimeType, size, parents, webViewLink",
            )
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

        When *remove_from_parents* is True (default) the file is removed from
        all current parents so it appears only in the new folder.
        """
        remove_parents: Optional[str] = None
        if remove_from_parents:
            meta = self.get_file_metadata(file_id, fields="parents")
            remove_parents = ",".join(meta.get("parents", []))

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
        """Replace the binary content of an existing Drive file."""
        media = MediaFileUpload(local_path, mimetype=mime_type, resumable=True)
        return (
            self._service.files()
            .update(
                fileId=file_id,
                media_body=media,
                fields="id, name, mimeType, modifiedTime",
            )
            .execute()
        )

    def update_file_metadata(
        self,
        file_id: str,
        metadata: Dict[str, Any],
    ) -> Dict[str, Any]:
        """
        Patch arbitrary metadata fields (name, description, …).

        Example::

            client.update_file_metadata(file_id, {"description": "Q1 report"})
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
        """Permanently delete a file (bypasses Trash). Irreversible."""
        self._service.files().delete(fileId=file_id).execute()

    def trash_file(self, file_id: str) -> Dict[str, Any]:
        """Move a file to the user's Trash (recoverable)."""
        return (
            self._service.files()
            .update(
                fileId=file_id,
                body={"trashed": True},
                fields="id, name, trashed",
            )
            .execute()
        )


# ---------------------------------------------------------------------- #
#  Config loader                                                           #
# ---------------------------------------------------------------------- #

def _load_config_from_env() -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Load OAuth configuration from environment variables.

    Returns:
        (client_secret_file, token_file, folder_id)
    """
    client_secret_file = (
        os.getenv("GOOGLE_OAUTH_CLIENT_SECRET_FILE", "").strip() or None
    )
    token_file = os.getenv("GOOGLE_OAUTH_TOKEN_FILE", "").strip() or None
    folder_id = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "").strip() or None

    # Resolve relative paths
    for var_path in (client_secret_file,):
        if var_path:
            candidate = Path(var_path).expanduser()
            if not candidate.is_absolute() and not candidate.exists():
                script_dir_candidate = Path(__file__).parent / var_path
                if script_dir_candidate.exists():
                    client_secret_file = str(script_dir_candidate)

    return client_secret_file, token_file, folder_id


def _auto_detect_client_secret() -> Optional[str]:
    """Find the first client_secret_*.json in the project directory."""
    candidates = sorted(Path(__file__).parent.glob("client_secret_*.json"))
    return str(candidates[0]) if candidates else None


# ---------------------------------------------------------------------- #
#  Entry point — smoke test                                               #
# ---------------------------------------------------------------------- #

if __name__ == "__main__":
    """
    Smoke-test: authenticates via OAuth2, creates a Google Sheet and a Google Doc
    in the configured folder (or Drive root), then lists files in that folder.

    Set in .env or OS environment:
      GOOGLE_OAUTH_CLIENT_SECRET_FILE = client_secret_*.json   (auto-detected if absent)
      GOOGLE_DRIVE_FOLDER_ID          = <your folder id>       (optional)
      GOOGLE_OAUTH_TOKEN_FILE         = token.json             (optional, default)
    """
    # Fix Windows console Unicode output
    if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    try:
        # Re-load .env when executed directly
        try:
            from dotenv import load_dotenv as _ld
            _ld(dotenv_path=Path(__file__).with_name(".env"), override=False)
        except Exception:
            pass

        client_secret_file, token_file, folder_id = _load_config_from_env()

        # Auto-detect client secret if not configured
        if not client_secret_file:
            client_secret_file = _auto_detect_client_secret()
            if client_secret_file:
                print(f"[auto-detect] Using client secret: {Path(client_secret_file).name}")
            else:
                raise RuntimeError(
                    "Client secret file not found.\n"
                    "Download it from Google Cloud Console → APIs & Services → Credentials\n"
                    "and place it in the project folder, or set GOOGLE_OAUTH_CLIENT_SECRET_FILE."
                )

        print(f"Client secret : {Path(client_secret_file).name}")
        print(f"Token file    : {token_file or 'token.json (default)'}")
        print(f"Target folder : {folder_id or '(Drive root)'}")
        print()

        client = GDriveUserClient(
            client_secret_file=client_secret_file,
            token_file=token_file,
            folder_id=folder_id,
        )

        print("✓ OAuth2 authentication successful\n")

        TARGET_FOLDER_ID = "1K8bkCd-9YG3yQ7ZcncwgFJfOamjL3St7"

        # ── CREATE: Google Sheet ──────────────────────────────────────
        print(f"Creating test Google Sheet in folder '{TARGET_FOLDER_ID}' …")
        sheet = client.create_google_sheet(
            "Test Sheet (gdrive_oauth_api)",
            parent_id=TARGET_FOLDER_ID,
        )
        print(f"  ✓ Created Sheet  | id: {sheet['id']}")
        print(f"    Link: {sheet.get('webViewLink', 'n/a')}")

        # ── CREATE: Google Doc ────────────────────────────────────────
        print(f"Creating test Google Doc in folder '{TARGET_FOLDER_ID}' …")
        doc = client.create_google_doc(
            "Test Doc (gdrive_oauth_api)",
            parent_id=TARGET_FOLDER_ID,
        )
        print(f"  ✓ Created Doc    | id: {doc['id']}")
        print(f"    Link: {doc.get('webViewLink', 'n/a')}")

        # ── READ: list files in target folder ─────────────────────────
        print()
        label = f"folder '{folder_id}'" if folder_id else "Drive root"
        print(f"Listing files in {label} …")
        print("=" * 70)

        files = client.list_files(TARGET_FOLDER_ID)
        if not files:
            print("(folder is empty)")
        else:
            print(f"{'#':<4} {'MIME TYPE':<50} NAME")
            print("-" * 70)
            for i, f in enumerate(files, 1):
                size_raw = f.get("size")
                size_str = f" ({int(size_raw)/1024:.1f} KB)" if size_raw else ""
                print(f"{i:<4} {f.get('mimeType',''):<50} {f.get('name','')}{size_str}")
            print("-" * 70)
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
                "Check that the Google Drive API is enabled in your Google Cloud project\n"
                "and that the OAuth consent screen includes the required scopes."
            )
        else:
            print(f"Google API error: {api_err}")
    except Exception as err:
        print(f"Error: {err}")
