import argparse
import json
import os
import sys
from pathlib import Path

from azure.ai.contentunderstanding import ContentUnderstandingClient
from azure.ai.contentunderstanding.models import AnalysisInput
from azure.core.credentials import AzureKeyCredential
from azure.identity import DefaultAzureCredential
from dotenv import load_dotenv


REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from app.payload import extension_from_payload, file_name_from_payload
from app.sharepoint import enrich_payload_from_sharepoint_url


load_dotenv(REPO_ROOT / ".env")


def main() -> None:
    parser = argparse.ArgumentParser(description="Test Content Understanding against a public media URL.")
    parser.add_argument("file_url", help="SAS URL or anonymous sharing URL for the media file.")
    parser.add_argument("--endpoint", default=os.environ.get("AZURE_CONTENT_UNDERSTANDING_ENDPOINT"))
    parser.add_argument("--analyzer-id", default=os.environ.get("CONTENT_UNDERSTANDING_ANALYZER_ID", "prebuilt-videoSearch"))
    parser.add_argument("--api-version", default=os.environ.get("CONTENT_UNDERSTANDING_API_VERSION", "2025-11-01"))
    args = parser.parse_args()

    if not args.endpoint:
        raise SystemExit("Set AZURE_CONTENT_UNDERSTANDING_ENDPOINT or pass --endpoint.")

    key = os.environ.get("CONTENT_UNDERSTANDING_KEY")
    credential = AzureKeyCredential(key) if key else DefaultAzureCredential()
    payload, analysis_url = enrich_payload_from_sharepoint_url({"fileUrl": args.file_url}, args.file_url)

    print(f"Resolved file name: {file_name_from_payload(payload)}")
    print(f"Resolved extension: {extension_from_payload(payload)}")
    print(f"Using analysis URL: {analysis_url}")

    client = ContentUnderstandingClient(
        endpoint=args.endpoint,
        credential=credential,
        api_version=args.api_version,
    )

    poller = client.begin_analyze(
        analyzer_id=args.analyzer_id,
        inputs=[AnalysisInput(url=analysis_url)],
    )
    result = poller.result()
    print(json.dumps(result.as_dict(), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
