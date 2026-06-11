from __future__ import annotations

from pathlib import PurePosixPath
from typing import Any, Optional
from urllib.parse import urlparse


MEDIA_EXTENSIONS = {
    ".3g2",
    ".3gp",
    ".avi",
    ".flv",
    ".m4v",
    ".mkv",
    ".mov",
    ".mp4",
    ".mpeg",
    ".mpg",
    ".ogg",
    ".ogv",
    ".webm",
    ".wmv",
}


def first_value(payload: dict[str, Any], *names: str) -> Optional[Any]:
    for name in names:
        if name in payload and payload[name]:
            return payload[name]
    return None


def file_url_from_payload(payload: dict[str, Any]) -> Optional[str]:
    value = first_value(
        payload,
        "fileUrl",
        "file_url",
        "downloadUrl",
        "@microsoft.graph.downloadUrl",
        "sourceUrl",
        "source_url",
    )
    return str(value) if value else None


def file_name_from_payload(payload: dict[str, Any]) -> str:
    value = first_value(payload, "fileName", "name", "displayName")
    return str(value) if value else ""


def extension_from_payload(payload: dict[str, Any]) -> str:
    if payload.get("extension"):
        extension = str(payload["extension"])
        extension = extension if extension.startswith(".") else f".{extension}"
        return extension.lower()

    file_name = file_name_from_payload(payload)
    if not file_name:
        file_url = file_url_from_payload(payload) or ""
        file_name = urlparse(str(file_url)).path

    return PurePosixPath(str(file_name or "")).suffix.lower()


def build_analysis_document(
    payload: dict[str, Any],
    file_url: str,
    extension: str,
    result: dict[str, Any],
) -> dict[str, Any]:
    return {
        "status": "analyzed",
        "source": {
            "fileName": file_name_from_payload(payload),
            "extension": extension,
            "fileUrl": file_url,
            "webUrl": payload.get("webUrl"),
            "siteId": payload.get("siteId"),
            "driveId": payload.get("driveId"),
            "itemId": payload.get("itemId"),
            "folderPath": payload.get("folderPath"),
            "parentItemId": payload.get("parentItemId"),
        },
        "contentUnderstanding": result,
    }
