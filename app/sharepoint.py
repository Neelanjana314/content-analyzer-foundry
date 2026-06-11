from __future__ import annotations

import base64
import json
from pathlib import PurePosixPath
from typing import Any, Optional
from urllib.parse import quote

import requests
from azure.identity import DefaultAzureCredential

from app.payload import first_value


GRAPH_SCOPE = "https://graph.microsoft.com/.default"
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


def graph_headers(credential: DefaultAzureCredential, content_type: str = "application/json") -> dict[str, str]:
    token = credential.get_token(GRAPH_SCOPE).token
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": content_type,
    }


def share_id_from_url(url: str) -> str:
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return f"u!{encoded}"


def resolve_output_folder(
    credential: DefaultAzureCredential,
    output_folder_url: str,
) -> tuple[str, str]:
    share_id = share_id_from_url(output_folder_url)
    response = requests.get(
        f"{GRAPH_BASE_URL}/shares/{share_id}/driveItem",
        headers=graph_headers(credential),
        timeout=60,
    )
    response.raise_for_status()

    item = response.json()
    drive_id = item.get("parentReference", {}).get("driveId")
    item_id = item.get("id")

    if not drive_id or not item_id:
        raise ValueError("Could not resolve outputFolderUrl to a SharePoint drive item.")

    return str(drive_id), str(item_id)


def sidecar_name(payload: dict[str, Any]) -> str:
    requested_name = payload.get("sidecarFileName")
    if requested_name:
        return str(requested_name)

    file_name = payload.get("fileName") or payload.get("name") or "content-understanding-result"
    path = PurePosixPath(str(file_name))
    return f"{path.stem}.txt"


def result_text(document: dict[str, Any]) -> str:
    return json.dumps(document, indent=2, ensure_ascii=False)


def upload_sidecar_to_sharepoint(payload: dict[str, Any], document: dict[str, Any]) -> Optional[dict[str, Any]]:
    if payload.get("skipSidecar") is True:
        return None

    site_id = payload.get("siteId")
    drive_id = payload.get("driveId")
    folder_path = payload.get("folderPath")
    parent_item_id = payload.get("parentItemId")
    output_folder_url = first_value(payload, "outputFolderUrl", "folderUrl", "output_folder_url")

    if not output_folder_url and (not site_id or not drive_id or not (folder_path or parent_item_id)):
        return None

    credential = DefaultAzureCredential()
    headers = graph_headers(credential, content_type="text/plain")
    output_name = sidecar_name(payload)
    sidecar_content = result_text(document).encode("utf-8")

    if output_folder_url:
        drive_id, parent_item_id = resolve_output_folder(credential, str(output_folder_url))
        upload_url = (
            f"{GRAPH_BASE_URL}/drives/{quote(str(drive_id), safe='')}"
            f"/items/{quote(str(parent_item_id), safe='')}:/{quote(output_name)}:/content"
        )
    elif parent_item_id:
        upload_url = (
            f"{GRAPH_BASE_URL}/sites/{quote(str(site_id), safe='')}"
            f"/drives/{quote(str(drive_id), safe='')}"
            f"/items/{quote(str(parent_item_id), safe='')}:/{quote(output_name)}:/content"
        )
    else:
        clean_folder = str(folder_path).strip("/")
        relative_path = f"{clean_folder}/{output_name}" if clean_folder else output_name
        encoded_path = "/".join(quote(part) for part in relative_path.split("/"))
        upload_url = (
            f"{GRAPH_BASE_URL}/sites/{quote(str(site_id), safe='')}"
            f"/drives/{quote(str(drive_id), safe='')}"
            f"/root:/{encoded_path}:/content"
        )

    response = requests.put(upload_url, headers=headers, data=sidecar_content, timeout=60)
    response.raise_for_status()
    return response.json()
