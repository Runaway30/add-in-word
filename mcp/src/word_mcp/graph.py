"""
Microsoft Graph API helpers for reading and writing .docx files on SharePoint / OneDrive.

Authentication uses the client-credentials flow (app-only).
Required environment variables:
    AZURE_TENANT_ID     – Directory (tenant) ID
    AZURE_CLIENT_ID     – Application (client) ID
    AZURE_CLIENT_SECRET – Client secret value
"""

from __future__ import annotations

import base64
import io
import os
from urllib.parse import urlparse, urlunparse

import httpx
import msal
from docx import Document

_GRAPH = "https://graph.microsoft.com/v1.0"
_SCOPE = ["https://graph.microsoft.com/.default"]

# Module-level MSAL app so the in-memory token cache is reused across calls.
_msal_app: msal.ConfidentialClientApplication | None = None


def _get_msal_app() -> msal.ConfidentialClientApplication:
    global _msal_app
    if _msal_app is None:
        tenant_id = os.environ["AZURE_TENANT_ID"]
        client_id = os.environ["AZURE_CLIENT_ID"]
        client_secret = os.environ["AZURE_CLIENT_SECRET"]
        _msal_app = msal.ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )
    return _msal_app


def _get_token() -> str:
    app = _get_msal_app()
    result = app.acquire_token_silent(_SCOPE, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=_SCOPE)
    if "access_token" not in result:
        raise RuntimeError(
            f"Failed to acquire Graph token: {result.get('error_description', result)}"
        )
    return result["access_token"]


def _encode_sharing_url(url: str) -> str:
    """Encode a SharePoint/OneDrive URL as a Graph API sharing token (u!<base64url>)."""
    parsed = urlparse(url)
    clean_url = urlunparse(parsed._replace(query="", fragment=""))
    encoded = base64.b64encode(clean_url.encode()).decode()
    encoded = encoded.rstrip("=").replace("+", "-").replace("/", "_")
    return f"u!{encoded}"


def _resolve_drive_item(url: str, token: str) -> tuple[str, str]:
    """Return (drive_id, item_id) for a SharePoint/OneDrive file URL."""
    # Handle doc2.aspx URLs (browser URL) by extracting the sourcedoc GUID
    # and resolving via the user's OneDrive drive directly.
    parsed = urlparse(url)
    from urllib.parse import parse_qs
    qs = parse_qs(parsed.query)
    if "sourcedoc" in qs:
        # Extract GUID e.g. {82E752B5-0984-423A-88DF-0DDE4A549AF0}
        guid = qs["sourcedoc"][0].strip("{}")
        # Derive the personal site path from the URL hostname + path
        # e.g. https://guggiaritest-my.sharepoint.com/personal/user/...
        path_parts = parsed.path.split("/")
        # path is like /personal/user/_layouts/15/doc2.aspx
        if "personal" in path_parts:
            idx = path_parts.index("personal")
            site_path = "/".join(path_parts[:idx + 2])  # /personal/user
        else:
            site_path = ""
        site_host = parsed.hostname
        # Get site id
        r = httpx.get(
            f"{_GRAPH}/sites/{site_host}:{site_path}",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "id"},
        )
        r.raise_for_status()
        site_id = r.json()["id"]
        # Get default drive
        r = httpx.get(
            f"{_GRAPH}/sites/{site_id}/drive",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "id"},
        )
        r.raise_for_status()
        drive_id = r.json()["id"]
        # Find item by GUID (Graph uses the GUID as the item id in OneDrive personal drives)
        r = httpx.get(
            f"{_GRAPH}/drives/{drive_id}/items/{guid}",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "id,parentReference"},
        )
        r.raise_for_status()
        return drive_id, r.json()["id"]

    share_id = _encode_sharing_url(url)
    r = httpx.get(
        f"{_GRAPH}/shares/{share_id}/driveItem",
        headers={"Authorization": f"Bearer {token}"},
        params={"$select": "id,parentReference"},
    )
    r.raise_for_status()
    data = r.json()
    return data["parentReference"]["driveId"], data["id"]


def download_docx(sharepoint_url: str) -> tuple[Document, str, str]:
    """
    Download a .docx from SharePoint / OneDrive.

    Returns:
        (Document, drive_id, item_id)  – drive_id and item_id are needed to upload back.
    """
    token = _get_token()
    drive_id, item_id = _resolve_drive_item(sharepoint_url, token)
    r = httpx.get(
        f"{_GRAPH}/drives/{drive_id}/items/{item_id}/content",
        headers={"Authorization": f"Bearer {token}"},
        follow_redirects=True,
    )
    r.raise_for_status()
    doc = Document(io.BytesIO(r.content))
    return doc, drive_id, item_id


def upload_docx(doc: Document, drive_id: str, item_id: str, retries: int = 5, retry_delay: float = 10.0) -> None:
    """Upload a modified Document back to SharePoint / OneDrive, replacing the existing file.

    Retries on HTTP 423 (file locked) up to `retries` times, waiting `retry_delay` seconds between attempts.
    """
    import time

    token = _get_token()
    buf = io.BytesIO()
    doc.save(buf)
    content = buf.getvalue()

    for attempt in range(retries):
        r = httpx.put(
            f"{_GRAPH}/drives/{drive_id}/items/{item_id}/content",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": (
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                ),
            },
            content=content,
        )
        if r.status_code == 423:
            if attempt < retries - 1:
                time.sleep(retry_delay)
                token = _get_token()
                continue
            raise RuntimeError(
                f"File is locked (HTTP 423) after {retries} attempts. "
                "Please close the document in Word and Word Online, then try again."
            )
        r.raise_for_status()
        return
