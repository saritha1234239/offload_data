# Offload Performance Dashboard

## Least-performing vendors and flights

The dashboard ranks vendors and flight numbers from highest bag offload rate
(worst performance) in the rolling 30-day window. An entity is listed when its
rate exceeds the network average in at least three of four periods, with at
least 20 eligible bags in each qualifying period. The watchlist is independent
of the dashboard filters. The dashboard is public, so only publish flight
information approved for public viewing.

## Automatically sync the workbook

Render can download a SharePoint/OneDrive **Anyone** link without Microsoft
sign-in only when the link returns the raw Excel file to server requests. The
downloader tries direct-download mode first, then the original URL, and accepts
only a valid Excel workbook response. Some SharePoint links work in a private
browser window but still redirect server requests to sign-in. In that case, use
the local-computer upload method below. Scheduled server refresh, when available,
runs every five minutes by default, and open dashboard tabs reload changed data
at most one minute later. Failed refreshes keep the last successful data.

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

## Sync from a computer that can access the workbook

If Render cannot download the SharePoint link, a Windows computer with the
workbook available locally can push validated workbook data to Render. For
automatic updates, use a OneDrive-synced copy of the workbook. A one-time
download will not change when the source workbook is edited; it must be replaced
manually unless OneDrive keeps that local copy synchronized. The computer must
be signed in, online, and have the workbook available for scheduled uploads.

1. In Render, set `DASHBOARD_UPLOAD_TOKEN` to a new, random secret of at least
   32 characters. For
   example, generate one locally with:

   ```powershell
   & .\.venv\Scripts\python.exe -c "import secrets; print(secrets.token_urlsafe(48))"
   ```

   Paste the generated value directly into Render's environment setting. Do not
   commit it or send it in chat. Save the setting and wait for deployment.
2. In a terminal on the Windows computer, configure the local workbook path.
   The token is entered with hidden input and saved in the current user's local
   application-data folder, not in this repository:

   ```powershell
   & .\.venv\Scripts\python.exe .\sync_workbook_to_dashboard.py --configure --workbook "C:\path\to\the\synced\workbook.xlsx"
   ```

   Configuration also performs an initial upload. The workbook must contain the
   four expected dashboard sheets. The local configuration file contains the
   upload token; keep the Windows user account private and do not share that file.
3. If the initial upload succeeds, install the five-minute scheduled task:

   ```powershell
   & .\.venv\Scripts\python.exe .\sync_workbook_to_dashboard.py --install-task
   ```

   The task runs every five minutes while that Windows user is signed in. To
   upload once manually, run `sync_workbook_to_dashboard.py --sync`. To stop
   scheduled uploads, run `sync_workbook_to_dashboard.py --remove-task`.
4. Check `https://offload-dashboard.onrender.com/api/health`. A successful push
   updates `workbookAutoSync.lastSuccessAt`, `rows`, and `version`; the
   `workbookUpload.enabled` field indicates whether the protected endpoint is
   configured.

The upload endpoint accepts only HTTPS requests with the configured bearer token,
limits workbook uploads to 30 MB, validates the workbook before replacing the
dashboard data, and keeps the previous data if validation fails. Once local
uploads are working, remove `WORKBOOK_DOWNLOAD_URL` from Render if it is no
longer needed. Anyone with the dashboard URL can see the published shipment data.

## Daily dashboard snapshot

The [GitHub Actions workflow](./.github/workflows/daily-dashboard-snapshot.yml)
runs at `09:00 UTC` (`2:30 PM` in India). GitHub-hosted runners use UTC for
scheduled workflows, and GitHub may delay a scheduled run during periods of
high load; delivery is not guaranteed to start at the exact minute. This avoids
adding a Render Cron service.

The workflow opens the hosted dashboard with Playwright, captures a full-page
PNG, and emails it to BCC recipients using the company SMTP relay. These are
SMTP settings, not Email API credentials. Keep all credentials in GitHub
repository secrets, never in the repository or chat.

In the repository, open **Settings → Secrets and variables → Actions** and add
these **repository secrets**:

| Secret | Value |
| --- | --- |
| `SMTP_HOST` | Company-approved relay hostname |
| `SMTP_USERNAME` | SMTP username |
| `SMTP_PASSWORD` | SMTP password |
| `EMAIL_RECIPIENTS` | Comma-separated authorized email addresses |

Add these **repository variables** on the **Variables** tab:

| Variable | Value |
| --- | --- |
| `DASHBOARD_EMAIL_ENABLED` | `true` after all required secrets are configured |
| `EMAIL_FROM` | An address approved by IT (defaults to `saritha.m@criticalog.com` if unset) |
| `SMTP_PORT` | `587` for STARTTLS, or `465` for implicit TLS; defaults to `587` |
| `SMTP_STARTTLS` | `true` for port 587; defaults to `true` |

The workflow is disabled until `DASHBOARD_EMAIL_ENABLED` is set to `true`.
After adding the settings, go to **Actions → Daily dashboard snapshot →
Run workflow** to test it. The message contains the PNG inline and hides the
recipient list. Check the run log for delivery status.

Ask IT to confirm the relay accepts connections from GitHub-hosted runners.
Hosted runner IP ranges can change, so IP-allowlist-only relays may not work
reliably; do not put SMTP credentials in workflow code.

WhatsApp remains optional and disabled if no WhatsApp settings are added. To
enable it later, configure the previously documented Cloud API values as
repository secrets and set `WHATSAPP_RECIPIENTS_OPTED_IN=true` only after all
recipients have opted in. The Meta-approved template must have an image header.
