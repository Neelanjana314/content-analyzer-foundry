# SharePoint Media Content Understanding Function

Timer-triggered Azure Function that scans a SharePoint folder for newly changed media files, analyzes each file with Azure Content Understanding, and writes the JSON result as a `.txt` sidecar file to another SharePoint folder.

The app uses Microsoft Graph and `DefaultAzureCredential`. In Azure, that means the Function App's managed identity can authenticate to both Microsoft Graph and Azure Content Understanding without storing client secrets.

## Flow

1. The timer runs on `SHAREPOINT_SCAN_SCHEDULE`.
2. The function resolves the SharePoint site and document library drive.
3. It reads a checkpoint file from the output folder.
4. It lists files in `SHAREPOINT_INPUT_FOLDER_PATH`, or the document library root when blank.
5. It filters to supported media files changed after the checkpoint.
6. For each media file, it asks Graph for `@microsoft.graph.downloadUrl`.
7. It sends that temporary download URL to Content Understanding.
8. It writes `<original-file-stem>.txt` to `SHAREPOINT_OUTPUT_FOLDER_PATH`, or the document library root when blank.
9. It advances the checkpoint only when all candidate files finish successfully.

On the first run, there is no checkpoint. The app scans files changed within `SHAREPOINT_INITIAL_LOOKBACK_MINUTES` so an initial deployment does not accidentally process an entire document library.

## App Settings

Required:

- `AzureWebJobsStorage`: required by Azure Functions timer triggers.
- `FUNCTIONS_WORKER_RUNTIME`: `python`.
- `AZURE_CONTENT_UNDERSTANDING_ENDPOINT`: your Content Understanding endpoint.
SharePoint site selection:

- Prefer `SHAREPOINT_SITE_ID` if you already know it.
- Otherwise set `SHAREPOINT_SITE_HOSTNAME`, for example `<tenant>.sharepoint.com`.
- And set `SHAREPOINT_SITE_PATH`, for example `/sites/Onboard`.

Optional:

- `SHAREPOINT_DRIVE_ID`: use a specific document library. If omitted, the site's default drive is used.
- `SHAREPOINT_OUTPUT_DRIVE_ID`: use this only if the output folder is in a different document library from the input folder.
- `SHAREPOINT_INPUT_FOLDER_PATH`: folder path inside the document library drive. Leave blank to scan the drive root.
- `SHAREPOINT_OUTPUT_FOLDER_PATH`: destination folder path inside the same drive. Leave blank to write to the drive root.
- `SHAREPOINT_SCAN_SCHEDULE`: NCRONTAB schedule. Default local fallback is every 15 minutes: `0 */15 * * * *`.
- `SHAREPOINT_RUN_ON_STARTUP`: set `true` only for local testing when you want one immediate run at host startup.
- `SHAREPOINT_USE_MONITOR`: keep `true` for Azure. You can set `false` for short local schedules such as every 30 seconds.
- `SHAREPOINT_STATE_FILE_PATH`: checkpoint file path. Defaults to `<output-folder>/.content-analyzer-state.json`, or `.content-analyzer-state.json` in the drive root when output is blank.
- `SHAREPOINT_INITIAL_LOOKBACK_MINUTES`: default `60`.
- `SHAREPOINT_RECURSIVE`: set `true` to scan subfolders.
- `CONTENT_UNDERSTANDING_KEY`: local testing convenience. In Azure, prefer managed identity and omit this.
- `CONTENT_UNDERSTANDING_ANALYZER_ID`: default `prebuilt-videoSearch`.
- `CONTENT_UNDERSTANDING_API_VERSION`: default `2025-11-01`.

## Least-Access Graph Permissions

Use the Function App system-assigned managed identity. For Graph, prefer `Sites.Selected` application permission and grant that identity access only to the target SharePoint site.

The identity needs enough site permission to:

- Read the input folder files.
- Read and write files in the output folder.
- Read/write the checkpoint file in the output folder.

In practice this usually means a site-level grant with write permission for the selected site. Avoid broad tenant-wide permissions such as `Sites.ReadWrite.All` unless your tenant process requires them.

For Content Understanding, grant the managed identity access to the Azure AI resource instead of using a key in Azure.

## Local Test

Copy `local.settings.json.sample` to `local.settings.json` and fill in real values. For local SharePoint Graph access, sign in with an identity that has access to the site:

```powershell
az login
# or
Connect-AzAccount
```

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Start the Function host:

```powershell
func start
```

For a quick first test, set `SHAREPOINT_INITIAL_LOOKBACK_MINUTES` high enough to include one known test video, or upload a fresh media file into the input folder after the function host starts. If the input/output folder settings are blank, the function scans and writes in the default document library root. The result should appear as a `.txt` file, and `.content-analyzer-state.json` should appear after a successful scan.

## Publish Test Checklist

1. Publish the Function App.
2. Enable system-assigned managed identity.
3. Add all app settings from `local.settings.json.sample`, except local-only secrets if using managed identity.
4. Grant the identity access to the Content Understanding resource.
5. Grant the identity `Sites.Selected` access to the target SharePoint site with write permission.
6. Upload a small `.mp4` to the configured input folder.
7. Watch Function logs for `ScheduledSharePointMediaScan`.
8. Confirm a matching `.txt` file appears in the output folder.
