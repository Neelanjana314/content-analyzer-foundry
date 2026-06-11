import argparse
import json
import os
from pathlib import Path

from azure.ai.contentunderstanding import ContentUnderstandingClient
from azure.ai.contentunderstanding.models import AnalysisInput
from azure.core.credentials import AzureKeyCredential
from azure.identity import DefaultAzureCredential
from dotenv import load_dotenv


load_dotenv(Path(__file__).resolve().parents[1] / ".env")


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

    client = ContentUnderstandingClient(
        endpoint=args.endpoint,
        credential=credential,
        api_version=args.api_version,
    )

    poller = client.begin_analyze(
        analyzer_id=args.analyzer_id,
        inputs=[AnalysisInput(url=args.file_url)],
    )
    result = poller.result()
    print(json.dumps(result.as_dict(), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
