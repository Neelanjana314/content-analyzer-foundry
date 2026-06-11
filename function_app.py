import base64
import json
import logging
import os
from pathlib import PurePosixPath
from typing import Any, Optional, Union
from urllib.parse import quote, urlparse

import azure.functions as func
import requests
from azure.ai.contentunderstanding import ContentUnderstandingClient
from azure.ai.contentunderstanding.models import AnalysisInput
from azure.core.credentials import AzureKeyCredential
from azure.core.exceptions import AzureError
from azure.identity import DefaultAzureCredential
from dotenv import load_dotenv


load_dotenv()

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

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

GRAPH_SCOPE = "https://graph.microsoft.com/.default"
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


def _json_response(body: dict[str, Any], status_code: int = 200) -> func.HttpResponse:
    return func.HttpResponse(
        json.dumps(body, default=str),
        status_code=status_code,
        mimetype="application/json",
    )


def _parse_body(req: func.HttpRequest) -> dict[str, Any]:
    try:
        return req.get_json()
    except ValueError:
        return {}


def _first_value(payload: dict[str, Any], *names: str) -> Optional[Any]:
    for name in names:
        if name in payload and payload[name]:
            return payload[name]
    return None


def _file_url_from_payload(payload: dict[str, Any]) -> Optional[str]:
    value = _first_value(
        payload,
        "fileUrl",
        "file_url",
        "downloadUrl",
        "@microsoft.graph.downloadUrl",
        "sourceUrl",
        "source_url",
    )
    return str(value) if value else None


def _file_name_from_payload(payload: dict[str, Any]) -> str:
    value = _first_value(payload, "fileName", "name", "displayName")
    return str(value) if value else ""


def _extension_from_payload(payload: dict[str, Any]) -> str:
    if payload.get("extension"):
        extension = str(payload["extension"])
        extension = extension if extension.startswith(".") else f".{extension}"
        return extension.lower()

    file_name = _file_name_from_payload(payload)
    if not file_name:
        file_url = _file_url_from_payload(payload) or ""
        file_name = urlparse(str(file_url)).path

    return PurePosixPath(str(file_name or "")).suffix.lower()


def _require_setting(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise ValueError(f"Missing required app setting: {name}")
    return value


def _content_understanding_credential() -> Union[AzureKeyCredential, DefaultAzureCredential]:
    key = os.environ.get("CONTENT_UNDERSTANDING_KEY")
    if key:
        return AzureKeyCredential(key)
    return DefaultAzureCredential()


def analyze_media(file_url: str) -> dict[str, Any]:
    endpoint = _require_setting("AZURE_CONTENT_UNDERSTANDING_ENDPOINT")
    analyzer_id = os.environ.get("CONTENT_UNDERSTANDING_ANALYZER_ID", "prebuilt-videoSearch")
    api_version = os.environ.get("CONTENT_UNDERSTANDING_API_VERSION", "2025-11-01")

    client = ContentUnderstandingClient(
        endpoint=endpoint,
        credential=_content_understanding_credential(),
        api_version=api_version,
    )

    poller = client.begin_analyze(
        analyzer_id=analyzer_id,
        inputs=[AnalysisInput(url=file_url)],
    )
    result = poller.result()
    return result.as_dict()


def build_analysis_document(payload: dict[str, Any], file_url: str, extension: str, result: dict[str, Any]) -> dict[str, Any]:
    return {
        "status": "analyzed",
        "source": {
            "fileName": _file_name_from_payload(payload),
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


def _graph_headers(credential: DefaultAzureCredential, content_type: str = "application/json") -> dict[str, str]:
    token = credential.get_token(GRAPH_SCOPE).token
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": content_type,
    }


def _share_id_from_url(url: str) -> str:
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return f"u!{encoded}"


def _resolve_output_folder(
    credential: DefaultAzureCredential,
    output_folder_url: str,
) -> tuple[str, str]:
    share_id = _share_id_from_url(output_folder_url)
    response = requests.get(
        f"{GRAPH_BASE_URL}/shares/{share_id}/driveItem",
        headers=_graph_headers(credential),
        timeout=60,
    )
    response.raise_for_status()

    item = response.json()
    drive_id = item.get("parentReference", {}).get("driveId")
    item_id = item.get("id")

    if not drive_id or not item_id:
        raise ValueError("Could not resolve outputFolderUrl to a SharePoint drive item.")

    return str(drive_id), str(item_id)


def _sidecar_name(payload: dict[str, Any]) -> str:
    requested_name = payload.get("sidecarFileName")
    if requested_name:
        return str(requested_name)

    file_name = payload.get("fileName") or payload.get("name") or "content-understanding-result"
    path = PurePosixPath(str(file_name))
    return f"{path.stem}.txt"


def _result_text(document: dict[str, Any]) -> str:
    return json.dumps(document, indent=2, ensure_ascii=False)


def upload_sidecar_to_sharepoint(payload: dict[str, Any], document: dict[str, Any]) -> Optional[dict[str, Any]]:
    if payload.get("skipSidecar") is True:
        return None

    site_id = payload.get("siteId")
    drive_id = payload.get("driveId")
    folder_path = payload.get("folderPath")
    parent_item_id = payload.get("parentItemId")
    output_folder_url = _first_value(payload, "outputFolderUrl", "folderUrl", "output_folder_url")

    if not output_folder_url and (not site_id or not drive_id or not (folder_path or parent_item_id)):
        return None

    credential = DefaultAzureCredential()
    headers = _graph_headers(credential, content_type="text/plain")
    sidecar_name = _sidecar_name(payload)
    sidecar_content = _result_text(document).encode("utf-8")

    if output_folder_url:
        drive_id, parent_item_id = _resolve_output_folder(credential, str(output_folder_url))
        upload_url = (
            f"{GRAPH_BASE_URL}/drives/{quote(str(drive_id), safe='')}"
            f"/items/{quote(str(parent_item_id), safe='')}:/{quote(sidecar_name)}:/content"
        )
    elif parent_item_id:
        upload_url = (
            f"{GRAPH_BASE_URL}/sites/{quote(str(site_id), safe='')}"
            f"/drives/{quote(str(drive_id), safe='')}"
            f"/items/{quote(str(parent_item_id), safe='')}:/{quote(sidecar_name)}:/content"
        )
    else:
        clean_folder = str(folder_path).strip("/")
        relative_path = f"{clean_folder}/{sidecar_name}" if clean_folder else sidecar_name
        encoded_path = "/".join(quote(part) for part in relative_path.split("/"))
        upload_url = (
            f"{GRAPH_BASE_URL}/sites/{quote(str(site_id), safe='')}"
            f"/drives/{quote(str(drive_id), safe='')}"
            f"/root:/{encoded_path}:/content"
        )

    response = requests.put(upload_url, headers=headers, data=sidecar_content, timeout=60)
    response.raise_for_status()
    return response.json()


@app.route(route="analyze-sharepoint-media", methods=["POST"])
def analyze_sharepoint_media(req: func.HttpRequest) -> func.HttpResponse:
    payload = _parse_body(req)
    file_url = _file_url_from_payload(payload)

    if not file_url:
        return _json_response(
            {
                "status": "bad_request",
                "message": "Request body must include fileUrl, downloadUrl, or @microsoft.graph.downloadUrl. Content Understanding must be able to read the URL directly.",
            },
            status_code=400,
        )

    extension = _extension_from_payload(payload)
    if extension not in MEDIA_EXTENSIONS:
        return _json_response(
            {
                "status": "skipped",
                "message": "File extension is not a supported media/video type.",
                "extension": extension,
            },
            status_code=200,
        )

    try:
        result = analyze_media(str(file_url))
        document = build_analysis_document(payload, str(file_url), extension, result)
        sidecar = upload_sidecar_to_sharepoint(payload, document)
    except AzureError as err:
        logging.exception("Content Understanding request failed.")
        return _json_response(
            {
                "status": "error",
                "stage": "content_understanding",
                "message": getattr(err, "message", str(err)),
            },
            status_code=502,
        )
    except requests.HTTPError as err:
        logging.exception("SharePoint sidecar upload failed.")
        response_text = err.response.text if err.response is not None else str(err)
        status_code = err.response.status_code if err.response is not None else None
        url = err.response.url if err.response is not None else None
        return _json_response(
            {
                "status": "error",
                "stage": "sharepoint_upload",
                "statusCode": status_code,
                "url": url,
                "message": response_text,
            },
            status_code=502,
        )
    except Exception as err:
        logging.exception("Unexpected failure.")
        return _json_response(
            {
                "status": "error",
                "stage": "unexpected",
                "message": str(err),
            },
            status_code=500,
        )

    return _json_response(
        {
            **document,
            "sidecarWritten": sidecar is not None,
            "sidecar": sidecar,
        }
    )
