from __future__ import annotations

import json
import logging
from typing import Any

import azure.functions as func
import requests
from azure.core.exceptions import AzureError
from dotenv import load_dotenv

from app.workflow import process_sharepoint_media


load_dotenv()

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)


def _json_response(body: dict[str, Any], status_code: int = 200) -> func.HttpResponse:
    return func.HttpResponse(
        json.dumps(body, default=str),
        status_code=status_code,
        mimetype="application/json",
    )


def _parse_body(req: func.HttpRequest) -> dict[str, Any]:
    try:
        payload = req.get_json()
    except ValueError:
        return {}

    return payload if isinstance(payload, dict) else {}


def _error_response_for_exception(err: Exception) -> func.HttpResponse:
    if isinstance(err, AzureError):
        logging.exception("Content Understanding request failed.")
        return _json_response(
            {
                "status": "error",
                "stage": "content_understanding",
                "message": getattr(err, "message", str(err)),
            },
            status_code=502,
        )

    if isinstance(err, requests.HTTPError):
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

    logging.exception("Unexpected failure.")
    return _json_response(
        {
            "status": "error",
            "stage": "unexpected",
            "message": str(err),
        },
        status_code=500,
    )


def _handle_analyze_sharepoint_media(req: func.HttpRequest) -> func.HttpResponse:
    try:
        response_payload, status_code = process_sharepoint_media(_parse_body(req))
    except Exception as err:
        return _error_response_for_exception(err)

    return _json_response(response_payload, status_code=status_code)


@app.route(route="analyze-sharepoint-media", methods=["POST"])
def analyze_sharepoint_media(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Processing SharePoint media analysis request.")
    return _handle_analyze_sharepoint_media(req)
