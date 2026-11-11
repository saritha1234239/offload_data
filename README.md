# Offload Performance Dashboard

## Automatically sync the workbook

Render can download the workbook from a SharePoint/OneDrive **Anyone** link,
without Microsoft sign-in. It tries the link as a direct download first, then
the original URL, and only accepts a valid Excel workbook response. It reads the
four dashboard tabs and refreshes `workbook-data.js` every five minutes by
default. Open dashboard tabs reload changed data at most one minute later. A
failed refresh keeps the last successful data.

To configure the public link, copy the workbook's **Anyone with the link** URL
from SharePoint/OneDrive and set it as `WORKBOOK_DOWNLOAD_URL` in the Render
service environment. Use the actual sharing link, not a document path that only
works from a signed-in browser. Check the link in a private window before saving
it in Render. The link can grant access to the entire workbook, including tabs
the dashboard does not use; only enable it if that access is approved.

`/api/health` reports whether sync is configured, when it last succeeded, and a
sanitized download error without exposing the configured URL.

### Microsoft Graph alternative

If public workbook access is not approved, the Render service also supports
Microsoft Graph app-only access to the existing workbook. Graph credentials stay
in Render's environment and are never returned by `/api/health`. A Microsoft
365 administrator must approve an application, and the workbook owner or
administrator must grant that app read access to this file.

### Microsoft 365 administrator setup

1. In Microsoft Entra ID, register a **single-tenant** application for the
   dashboard.
2. Under **API permissions**, add Microsoft Graph **Application** permission
   `Files.SelectedOperations.Selected`, then grant tenant admin consent.
3. Create a client secret (or certificate) for the application. Copy the secret
   value securely; do not put it in this repository or send it in chat.
4. Sign in to [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
   as an account that can access the tracker. Use the `GET` request below to find
   the file's `id` and `parentReference.driveId`:

   ```http
   GET https://graph.microsoft.com/v1.0/me/drive/root:/Desktop/Midmile%20Management/LH%20performance%20tracker.xlsx?$select=id,name,parentReference
   ```

   If Graph Explorer cannot resolve the path, browse to the tracker in OneDrive,
   then resolve it with that account's actual drive-relative path. The file ID and
   drive ID are identifiers, not credentials.
5. Grant the application **read** access to only that file. In Graph Explorer,
   with an account that can manage permissions on the file, submit:

   ```http
   POST https://graph.microsoft.com/v1.0/drives/{drive-id}/items/{item-id}/permissions
   Content-Type: application/json

   {
     "roles": ["read"],
     "grantedToV2": {
       "application": {
         "id": "{dashboard-application-client-id}"
       }
     }
   }
   ```

   The signed-in permission-granting account needs `Files.ReadWrite` delegated
   Graph permission and access to manage the file. The app still has only
   `Files.SelectedOperations.Selected` application permission; the explicit
   item-level grant limits it to this workbook. Microsoft's [selected
   permissions guide](https://learn.microsoft.com/graph/permissions-selected-overview)
   explains that both admin consent and a resource assignment are required.

### Configure Render

In the `offload-dashboard` Render service, set these environment variables, then
redeploy:

| Variable | Value |
| --- | --- |
| `MICROSOFT_TENANT_ID` | Entra tenant ID |
| `MICROSOFT_CLIENT_ID` | Dashboard app's application (client) ID |
| `MICROSOFT_CLIENT_SECRET` | Secret **value**; enter only in Render |
| `MICROSOFT_GRAPH_DRIVE_ID` | `parentReference.driveId` returned above |
| `MICROSOFT_GRAPH_ITEM_ID` | Workbook `id` returned above |
| `WORKBOOK_REFRESH_SECONDS` | Optional; defaults to `300` (minimum `60`) |

When Graph settings are present, they take precedence over
`WORKBOOK_DOWNLOAD_URL`. Do not add a public sharing link for the source workbook.
`/api/health` reports the sync source, last successful refresh, row count, and
configuration/download error without exposing any credentials.

If authentication, permission, download, or workbook validation fails, the service
keeps serving the last successful snapshot. The dashboard itself remains public;
its shipment records can be viewed by anyone with the dashboard URL. Only configure
it if that data is approved for public access.