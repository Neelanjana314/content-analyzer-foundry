from __future__ import annotations

import os
from typing import Any, Union

from azure.ai.contentunderstanding import ContentUnderstandingClient
from azure.ai.contentunderstanding.models import AnalysisInput
from azure.core.credentials import AzureKeyCredential
from azure.identity import DefaultAzureCredential


def require_setting(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise ValueError(f"Missing required app setting: {name}")
    return value


def content_understanding_credential() -> Union[AzureKeyCredential, DefaultAzureCredential]:
    key = os.environ.get("CONTENT_UNDERSTANDING_KEY")
    if key:
        return AzureKeyCredential(key)
    return DefaultAzureCredential()


def analyze_media(file_url: str) -> dict[str, Any]:
    endpoint = require_setting("AZURE_CONTENT_UNDERSTANDING_ENDPOINT")
    analyzer_id = os.environ.get("CONTENT_UNDERSTANDING_ANALYZER_ID", "prebuilt-videoSearch")
    api_version = os.environ.get("CONTENT_UNDERSTANDING_API_VERSION", "2025-11-01")

    client = ContentUnderstandingClient(
        endpoint=endpoint,
        credential=content_understanding_credential(),
        api_version=api_version,
    )

    poller = client.begin_analyze(
        analyzer_id=analyzer_id,
        inputs=[AnalysisInput(url=file_url)],
    )
    result = poller.result()
    return result.as_dict()
