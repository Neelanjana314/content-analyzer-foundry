from __future__ import annotations

import logging
import os

import azure.functions as func
import requests
from azure.core.exceptions import AzureError
from dotenv import load_dotenv

from app.workflow import scan_sharepoint_delta


load_dotenv()

app = func.FunctionApp()


def _env_bool(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.lower() in {"1", "true", "yes"}


@app.function_name(name="ScheduledSharePointMediaScan")
@app.timer_trigger(
    schedule="%SHAREPOINT_SCAN_SCHEDULE%",
    arg_name="timer",
    run_on_startup=_env_bool("SHAREPOINT_RUN_ON_STARTUP"),
    use_monitor=_env_bool("SHAREPOINT_USE_MONITOR", True),
)
def scheduled_sharepoint_media_scan(timer: func.TimerRequest) -> None:
    if timer.past_due:
        logging.warning("Scheduled SharePoint media scan is running later than expected.")

    logging.info("Starting scheduled SharePoint media scan.")
    try:
        summary = scan_sharepoint_delta()
    except (AzureError, requests.HTTPError):
        logging.exception("Scheduled SharePoint media scan failed.")
        raise
    except Exception:
        logging.exception("Unexpected scheduled SharePoint media scan failure.")
        raise

    logging.info("Scheduled SharePoint media scan summary: %s", summary)
