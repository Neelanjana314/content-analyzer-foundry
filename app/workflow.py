from __future__ import annotations

import logging
from datetime import datetime, timezone
from pathlib import PurePosixPath
from typing import Any

from azure.identity import DefaultAzureCredential

from app.content_understanding import analyze_media
from app.payload import build_analysis_document
from app.sharepoint import (
    changed_media_items,
    fallback_checkpoint,
    format_graph_datetime,
    item_download_url,
    item_modified_at,
    list_folder_items,
    load_sharepoint_settings,
    read_checkpoint,
    resolve_drive_id,
    resolve_site_id,
    upload_result_sidecar,
    write_checkpoint,
)


def payload_from_drive_item(item: dict[str, Any], site_id: str, drive_id: str) -> dict[str, Any]:
    parent = item.get("parentReference", {})
    return {
        "fileName": item.get("name"),
        "webUrl": item.get("webUrl"),
        "siteId": site_id,
        "driveId": drive_id,
        "itemId": item.get("id"),
        "parentItemId": parent.get("id"),
    }


def process_drive_item(
    credential: DefaultAzureCredential,
    site_id: str,
    drive_id: str,
    output_drive_id: str,
    output_folder_path: str,
    item: dict[str, Any],
) -> dict[str, Any]:
    item_id = str(item["id"])
    file_url = item_download_url(credential, drive_id, item_id)
    extension = PurePosixPath(str(item.get("name", ""))).suffix.lower()
    payload = payload_from_drive_item(item, site_id, drive_id)

    result = analyze_media(file_url)
    document = build_analysis_document(payload, file_url, extension, result)
    sidecar = upload_result_sidecar(credential, output_drive_id, output_folder_path, item, document)

    return {
        "itemId": item_id,
        "fileName": item.get("name"),
        "modifiedAt": format_graph_datetime(item_modified_at(item)),
        "sidecarWebUrl": sidecar.get("webUrl"),
    }


def scan_sharepoint_delta() -> dict[str, Any]:
    settings = load_sharepoint_settings()
    credential = DefaultAzureCredential()
    scan_started_at = datetime.now(timezone.utc)

    site_id = resolve_site_id(credential, settings)
    drive_id = resolve_drive_id(credential, settings, site_id)
    output_drive_id = settings.output_drive_id or drive_id
    checkpoint = read_checkpoint(credential, output_drive_id, settings.state_file_path)
    since = checkpoint or fallback_checkpoint(settings, scan_started_at)

    logging.info(
        "Scanning SharePoint media changed after %s in drive %s folder %s.",
        format_graph_datetime(since),
        drive_id,
        settings.input_folder_path,
    )

    items = list_folder_items(
        credential,
        drive_id,
        settings.input_folder_path,
        recursive=settings.recursive,
    )
    candidates = changed_media_items(items, since, scan_started_at)

    processed: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    for item in candidates:
        try:
            processed.append(
                process_drive_item(
                    credential,
                    site_id,
                    drive_id,
                    output_drive_id,
                    settings.output_folder_path,
                    item,
                )
            )
        except Exception as err:
            logging.exception("Failed to process SharePoint item %s.", item.get("id"))
            errors.append(
                {
                    "itemId": item.get("id"),
                    "fileName": item.get("name"),
                    "message": str(err),
                }
            )

    summary = {
        "status": "error" if errors else "ok",
        "since": format_graph_datetime(since),
        "scanStartedAt": format_graph_datetime(scan_started_at),
        "itemsScanned": len(items),
        "mediaCandidates": len(candidates),
        "processed": processed,
        "errors": errors,
    }

    if errors:
        logging.warning(
            "SharePoint scan completed with %s error(s); checkpoint was not advanced.",
            len(errors),
        )
        return summary

    write_checkpoint(credential, output_drive_id, settings.state_file_path, scan_started_at, summary)
    logging.info("SharePoint scan completed; processed %s media file(s).", len(processed))
    return summary
