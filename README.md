# SharePoint Media Content Understanding Function

HTTP-triggered Azure Function that accepts a SharePoint file event payload from Power Automate, skips non-media files, sends a public file URL to Azure Content Understanding, waits for the analysis result, and optionally writes a `.txt` sidecar file back to SharePoint through Microsoft Graph.

## Request

POST `/api/analyze-sharepoint-media`

```json
{
  "fileUrl": "https://<direct-download-url-to-video>",
  "fileName": "demo.mp4",
  "outputFolderUrl": "https://<sharepoint-folder-sharing-url>"
}
```

The Function also accepts `downloadUrl` or `@microsoft.graph.downloadUrl` instead of `fileUrl`. For SharePoint/OneDrive, prefer the Microsoft Graph temporary download URL because it points to the actual file bytes rather than the browser viewer page.

For the sidecar destination, use one of these options:

- `outputFolderUrl`: SharePoint/OneDrive folder sharing URL.
- `siteId`, `driveId`, and `folderPath`.
- `siteId`, `driveId`, and `parentItemId`.

If none is supplied, the Function returns the Content Understanding result but does not write back to SharePoint.

Content Understanding must be able to read the URL directly. A SharePoint browser/view link usually returns HTML and may produce an empty `contents` array.

## Output

The Function returns JSON and writes the same shape into the SharePoint `.txt` sidecar when sidecar fields are provided:

```json
{
  "status": "analyzed",
  "source": {
    "fileName": "demo.mp4",
    "extension": ".mp4",
    "fileUrl": "https://<direct-download-url-to-video>",
    "siteId": null,
    "driveId": null,
    "folderPath": null
  },
  "contentUnderstanding": {
    "analyzerId": "prebuilt-videoSearch",
    "contents": []
  },
  "sidecarWritten": true,
  "sidecar": {}
}
```

## App Settings

For local development, copy `.env.sample` to `.env` and set:

- `AZURE_CONTENT_UNDERSTANDING_ENDPOINT`
- `CONTENT_UNDERSTANDING_KEY`
- `CONTENT_UNDERSTANDING_ANALYZER_ID`
- `CONTENT_UNDERSTANDING_API_VERSION`

You can also copy `local.settings.json.sample` to `local.settings.json` if you prefer the native Azure Functions local settings style. Do not commit either `.env` or `local.settings.json`.

For local testing, `CONTENT_UNDERSTANDING_KEY` is used when present. If that setting is omitted, the app uses `DefaultAzureCredential`.

For Azure deployment, prefer Managed Identity: remove `CONTENT_UNDERSTANDING_KEY` from the Function App settings and grant the managed identity access to the Content Understanding resource. Microsoft Graph sidecar uploads always use `DefaultAzureCredential`.

In Azure, enable the Function App managed identity and grant it:

- Access to the Content Understanding resource.
- Microsoft Graph application permissions needed to write to SharePoint, such as `Sites.Selected` with site grant or the appropriate broader SharePoint write permission your tenant allows.

## Local Test

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

First test only the Content Understanding call with a real video URL:

```powershell
. .\scripts\Import-DotEnv.ps1
python .\scripts\test_content_understanding.py "https://<public-or-sas-url-to-video>"
```

Start the Functions host:

```powershell
func start
```

Call it with a real video URL:

```powershell
$body = @{
  fileUrl = "https://<public-or-sas-url-to-video>"
  fileName = "sample.mp4"
  outputFolderUrl = "https://<sharepoint-folder-sharing-url>"
  skipSidecar = $true
} | ConvertTo-Json

Invoke-RestMethod `
  -Method Post `
  -Uri "http://localhost:7071/api/analyze-sharepoint-media" `
  -ContentType "application/json" `
  -Body $body
```
