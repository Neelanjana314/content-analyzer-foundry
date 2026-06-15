from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import PurePosixPath
from typing import Any, Optional
from urllib.parse import quote

import requests
from azure.identity import DefaultAzureCredential

from app.payload import MEDIA_EXTENSIONS


GRAPH_SCOPE = "https://graph.microsoft.com/.default"
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
DEFAULT_SCHEDULE = "0 */15 * * * *"
DEFAULT_INITIAL_LOOKBACK_MINUTES = 60


@dataclass(frozen=True)
class SharePointSettings:
    site_hostname: Optional[str]
    site_path: Optional[str]
    site_id: Optional[str]
    drive_id: Optional[str]
    output_drive_id: Optional[str]
    input_folder_path: str
    output_folder_path: str
    state_file_path: str
    initial_lookback_minutes: int
    recursive: bool


def load_sharepoint_settings() -> SharePointSettings:
    output_folder_path = normalize_drive_path(os.environ.get("SHAREPOINT_OUTPUT_FOLDER_PATH", ""))
    state_file_path = os.environ.get("SHAREPOINT_STATE_FILE_PATH")
    if not state_file_path:
        state_file_path = (
            f"{output_folder_path}/.content-analyzer-state.json"
            if output_folder_path
            else ".content-analyzer-state.json"
        )

    return SharePointSettings(
        site_hostname=os.environ.get("SHAREPOINT_SITE_HOSTNAME"),
        site_path=os.environ.get("SHAREPOINT_SITE_PATH"),
        site_id=os.environ.get("SHAREPOINT_SITE_ID"),
        drive_id=os.environ.get("SHAREPOINT_DRIVE_ID"),
        output_drive_id=os.environ.get("SHAREPOINT_OUTPUT_DRIVE_ID"),
        input_folder_path=normalize_drive_path(os.environ.get("SHAREPOINT_INPUT_FOLDER_PATH", "")),
        output_folder_path=output_folder_path,
        state_file_path=normalize_drive_path(state_file_path),
        initial_lookback_minutes=int(
            os.environ.get("SHAREPOINT_INITIAL_LOOKBACK_MINUTES", DEFAULT_INITIAL_LOOKBACK_MINUTES)
        ),
        recursive=os.environ.get("SHAREPOINT_RECURSIVE", "false").lower() in {"1", "true", "yes"},
    )


def graph_headers(credential: DefaultAzureCredential, content_type: Optional[str] = "application/json") -> dict[str, str]:
    token = credential.get_token(GRAPH_SCOPE).token
    headers = {"Authorization": f"Bearer {token}"}
    if content_type:
        headers["Content-Type"] = content_type
    return headers


def normalize_drive_path(path: str) -> str:
    return str(PurePosixPath(path.strip().strip("/")))


def encode_drive_path(path: str) -> str:
    clean_path = normalize_drive_path(path)
    if clean_path in {"", "."}:
        return ""
    return "/".join(quote(part, safe="") for part in clean_path.split("/"))


def drive_root_url(drive_id: str, path: str, suffix: str = "") -> str:
    encoded_path = encode_drive_path(path)
    drive = quote(drive_id, safe="")
    if not encoded_path:
        return f"{GRAPH_BASE_URL}/drives/{drive}/root{suffix}"
    return f"{GRAPH_BASE_URL}/drives/{drive}/root:/{encoded_path}:{suffix}"


def graph_get_json(credential: DefaultAzureCredential, url: str) -> dict[str, Any]:
    response = requests.get(url, headers=graph_headers(credential), timeout=60)
    response.raise_for_status()
    return response.json()


def graph_get_text(credential: DefaultAzureCredential, url: str) -> Optional[str]:
    response = requests.get(url, headers=graph_headers(credential, content_type=None), timeout=60)
    if response.status_code == 404:
        return None
    response.raise_for_status()
    return response.text


def graph_put_text(
    credential: DefaultAzureCredential,
    url: str,
    content: str,
    content_type: str = "text/plain",
) -> dict[str, Any]:
    response = requests.put(
        url,
        headers=graph_headers(credential, content_type=content_type),
        data=content.encode("utf-8"),
        timeout=60,
    )
    response.raise_for_status()
    return response.json()


def resolve_site_id(credential: DefaultAzureCredential, settings: SharePointSettings) -> str:
    if settings.site_id:
        return settings.site_id

    if not settings.site_hostname or not settings.site_path:
        raise ValueError(
            "Set SHAREPOINT_SITE_ID, or set both SHAREPOINT_SITE_HOSTNAME and SHAREPOINT_SITE_PATH."
        )

    site_path = settings.site_path if settings.site_path.startswith("/") else f"/{settings.site_path}"
    url = f"{GRAPH_BASE_URL}/sites/{settings.site_hostname}:{site_path}"
    return str(graph_get_json(credential, url)["id"])


def resolve_drive_id(credential: DefaultAzureCredential, settings: SharePointSettings, site_id: str) -> str:
    if settings.drive_id:
        return settings.drive_id

    url = f"{GRAPH_BASE_URL}/sites/{quote(site_id, safe='')}/drive"
    return str(graph_get_json(credential, url)["id"])


def paged_values(credential: DefaultAzureCredential, url: str) -> list[dict[str, Any]]:
    values: list[dict[str, Any]] = []
    next_url: Optional[str] = url
    while next_url:
        data = graph_get_json(credential, next_url)
        values.extend(data.get("value", []))
        next_url = data.get("@odata.nextLink")
    return values


def list_folder_items(
    credential: DefaultAzureCredential,
    drive_id: str,
    folder_path: str,
    recursive: bool = False,
) -> list[dict[str, Any]]:
    select = "id,name,file,folder,lastModifiedDateTime,createdDateTime,webUrl,parentReference,size"
    base_url = drive_root_url(drive_id, folder_path, "/children")
    items = paged_values(credential, f"{base_url}?$top=999&$select={select}")
    if not recursive:
        return items

    all_items = list(items)
    all_items.extend(list_nested_folder_items(credential, drive_id, items, select))
    return all_items


def list_nested_folder_items(
    credential: DefaultAzureCredential,
    drive_id: str,
    items: list[dict[str, Any]],
    select: str,
) -> list[dict[str, Any]]:
    nested_items: list[dict[str, Any]] = []
    folders = [item for item in items if item.get("folder") and item.get("id")]
    for folder in folders:
        folder_id = quote(str(folder["id"]), safe="")
        child_url = f"{GRAPH_BASE_URL}/drives/{quote(drive_id, safe='')}/items/{folder_id}/children"
        children = paged_values(credential, f"{child_url}?$top=999&$select={select}")
        nested_items.extend(children)
        nested_items.extend(list_nested_folder_items(credential, drive_id, children, select))
    return nested_items


def read_checkpoint(
    credential: DefaultAzureCredential,
    drive_id: str,
    state_file_path: str,
) -> Optional[datetime]:
    text = graph_get_text(credential, drive_root_url(drive_id, state_file_path, "/content"))
    if not text:
        return None

    data = json.loads(text)
    value = data.get("lastCompletedAt")
    if not value:
        return None
    return parse_graph_datetime(str(value))


def write_checkpoint(
    credential: DefaultAzureCredential,
    drive_id: str,
    state_file_path: str,
    completed_at: datetime,
    summary: dict[str, Any],
) -> dict[str, Any]:
    document = {
        "lastCompletedAt": format_graph_datetime(completed_at),
        "summary": summary,
    }
    return graph_put_text(
        credential,
        drive_root_url(drive_id, state_file_path, "/content"),
        json.dumps(document, indent=2),
        content_type="application/json",
    )


def parse_graph_datetime(value: str) -> datetime:
    normalized = value.replace("Z", "+00:00")
    return datetime.fromisoformat(normalized).astimezone(timezone.utc)


def format_graph_datetime(value: datetime) -> str:
    return value.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def fallback_checkpoint(settings: SharePointSettings, now: datetime) -> datetime:
    return now - timedelta(minutes=settings.initial_lookback_minutes)


def item_modified_at(item: dict[str, Any]) -> datetime:
    value = item.get("lastModifiedDateTime") or item.get("createdDateTime")
    if not value:
        return datetime.min.replace(tzinfo=timezone.utc)
    return parse_graph_datetime(str(value))


def is_media_item(item: dict[str, Any]) -> bool:
    if not item.get("file"):
        return False
    suffix = PurePosixPath(str(item.get("name", ""))).suffix.lower()
    return suffix in MEDIA_EXTENSIONS


def changed_media_items(
    items: list[dict[str, Any]],
    since: datetime,
    until: datetime,
) -> list[dict[str, Any]]:
    changed = []
    for item in items:
        modified_at = item_modified_at(item)
        if is_media_item(item) and since < modified_at <= until:
            changed.append(item)
    return sorted(changed, key=item_modified_at)


def item_download_url(
    credential: DefaultAzureCredential,
    drive_id: str,
    item_id: str,
) -> str:
    url = f"{GRAPH_BASE_URL}/drives/{quote(drive_id, safe='')}/items/{quote(item_id, safe='')}"
    item = graph_get_json(credential, url)
    download_url = item.get("@microsoft.graph.downloadUrl")
    if not download_url:
        raise ValueError(f"Graph did not return a download URL for item {item_id}.")
    return str(download_url)


def result_sidecar_path(output_folder_path: str, item: dict[str, Any]) -> str:
    file_name = str(item.get("name") or "content-understanding-result")
    sidecar_name = f"{PurePosixPath(file_name).stem}.txt"
    output_folder = normalize_drive_path(output_folder_path)
    return f"{output_folder}/{sidecar_name}" if output_folder else sidecar_name


def upload_result_sidecar(
    credential: DefaultAzureCredential,
    drive_id: str,
    output_folder_path: str,
    item: dict[str, Any],
    document: dict[str, Any],
) -> dict[str, Any]:
    return graph_put_text(
        credential,
        drive_root_url(drive_id, result_sidecar_path(output_folder_path, item), "/content"),
        json.dumps(document, indent=2, ensure_ascii=False),
        content_type="text/plain",
    )
