"""Google Drive storage for plan files uploaded by planners.

Authentication uses a service account. Its JSON key is read from the
environment variable GOOGLE_SERVICE_ACCOUNT_JSON, and the target folder id
from DRIVE_ROOT_FOLDER_ID. The root folder must be shared with the service
account (Editor). Files are stored as root / site / discipline / file name.
"""
import io
import json
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]
FOLDER_MIME = "application/vnd.google-apps.folder"


class DriveNotConfigured(RuntimeError):
    """Raised when the Drive environment variables are missing or invalid."""


def is_configured():
    return bool(
        os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
        and os.environ.get("DRIVE_ROOT_FOLDER_ID", "").strip()
    )


def _root_folder_id():
    rid = os.environ.get("DRIVE_ROOT_FOLDER_ID", "").strip()
    if not rid:
        raise DriveNotConfigured("DRIVE_ROOT_FOLDER_ID לא מוגדר")
    return rid


def _service():
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        raise DriveNotConfigured("GOOGLE_SERVICE_ACCOUNT_JSON לא מוגדר")
    try:
        info = json.loads(raw)
    except ValueError as exc:
        raise DriveNotConfigured("GOOGLE_SERVICE_ACCOUNT_JSON אינו JSON תקין") from exc
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _escape_query(value):
    return value.replace("\\", "\\\\").replace("'", "\\'")


def ensure_folder(svc, parent_id, name):
    """Return the id of a child folder with this name, creating it if needed."""
    query = (
        f"name = '{_escape_query(name)}' and '{parent_id}' in parents "
        f"and mimeType = '{FOLDER_MIME}' and trashed = false"
    )
    found = svc.files().list(
        q=query,
        fields="files(id)",
        spaces="drive",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute().get("files", [])
    if found:
        return found[0]["id"]
    meta = {"name": name, "mimeType": FOLDER_MIME, "parents": [parent_id]}
    return svc.files().create(body=meta, fields="id", supportsAllDrives=True).execute()["id"]


def upload_plan_file(data, filename, mime, site_name, discipline):
    """Upload a plan file under root/site/discipline.

    Returns a dict with the Drive file id and a link that opens it in Drive.
    """
    svc = _service()
    site_folder = ensure_folder(svc, _root_folder_id(), site_name)
    discipline_folder = ensure_folder(svc, site_folder, discipline)
    media = MediaIoBaseUpload(
        io.BytesIO(data), mimetype=mime or "application/octet-stream", resumable=False
    )
    meta = {"name": filename, "parents": [discipline_folder]}
    created = svc.files().create(
        body=meta,
        media_body=media,
        fields="id,webViewLink",
        supportsAllDrives=True,
    ).execute()
    file_id = created["id"]
    return {
        "file_id": file_id,
        "url": created.get("webViewLink") or f"https://drive.google.com/file/d/{file_id}/view",
    }
