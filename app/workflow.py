from __future__ import annotations

from typing import Any

from app.content_understanding import analyze_media
from app.payload import (
    MEDIA_EXTENSIONS,
    build_analysis_document,
    extension_from_payload,
    file_name_from_payload,
    file_url_from_payload,
    first_value,
)
from app.sharepoint import enrich_payload_from_sharepoint_url, is_sharepoint_url, upload_sidecar_to_sharepoint


def should_resolve_sharepoint_url(payload: dict[str, Any], file_url: str, extension: str) -> bool:
    if not is_sharepoint_url(file_url):
        return False

    if not extension:
        return True

    has_direct_download_url = first_value(payload, "downloadUrl", "@microsoft.graph.downloadUrl")
    return extension in MEDIA_EXTENSIONS and not has_direct_download_url


def process_sharepoint_media(payload: dict[str, Any]) -> tuple[dict[str, Any], int]:
    file_url = file_url_from_payload(payload)

    if not file_url:
        return (
            {
                "status": "bad_request",
                "message": "Request body must include fileUrl, downloadUrl, or @microsoft.graph.downloadUrl. Content Understanding must be able to read the URL directly.",
            },
            400,
        )

    extension = extension_from_payload(payload)
    if should_resolve_sharepoint_url(payload, str(file_url), extension):
        payload, file_url = enrich_payload_from_sharepoint_url(payload, str(file_url))
        extension = extension_from_payload(payload)

    if extension not in MEDIA_EXTENSIONS:
        return (
            {
                "status": "skipped",
                "message": "File extension is not a supported media/video type.",
                "fileName": file_name_from_payload(payload),
                "extension": extension,
            },
            200,
        )

    result = analyze_media(str(file_url))
    document = build_analysis_document(payload, str(file_url), extension, result)
    sidecar = upload_sidecar_to_sharepoint(payload, document)

    return (
        {
            **document,
            "sidecarWritten": sidecar is not None,
            "sidecar": sidecar,
        },
        200,
    )
