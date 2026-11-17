# Offload Performance Dashboard

## Four-week vendor and lane watchlist

The dashboard highlights vendors and lanes with below-average bag offload
rates in at least three of four consecutive seven-day periods in the rolling
28-day window. Each counted period must include at least 20 valid bags. A lower
offload percentage indicates poorer performance, and listed vendors and lanes
are flagged for follow-up. The watchlist is independent of the dashboard
filters.

## LCCS tracker query and mapping

`lccs_tracker.py` contains a parameterized source query using the tables and
join keys from `Multijoins::lccscostQuery()`, plus the 24-column tracker mapping.
It uses `XAWBID` as the source identifier, maps the fields present in that query,
applies the PHP report's flight-number fallback and ETD slot boundaries, and
rejects conflicting duplicate IDs. Manual values are supplied separately so
source refreshes do not replace them. The 37 `NDC_Remarks` dropdown choices
from the supplied screenshot are stored in order and validated on mapping.

`LH Performance Tracker.xlsx` is a blank workbook template with the 24 headers
at D4:AA4 and the NDC dropdown applied to column X for rows 5–5000. It is not
connected to the LCCS pipeline or the public dashboard; share the file through
OneDrive/SharePoint and assign team permissions there.

The tracker covers all hubs (no per-user hub filter) and air MAWBs only
(`XAWBMODE = 'A'`).

To refresh the workbook (defaults: yesterday to today, all hubs):

```powershell
& .\.venv\Scripts\python.exe .\lccs_tracker_job.py --workbook "C:\path\to\LH Performance Tracker.xlsx"
```

For the shared OneDrive copy, save its location once and install the hourly
Windows task (runs while the user is signed in; results go to
`%LOCALAPPDATA%\OffloadDashboard\lccs-tracker.log`):

```powershell
& .\.venv\Scripts\python.exe .\lccs_tracker_job.py --set-workbook "C:\path\to\OneDrive\LH Performance Tracker - LCCS.xlsx"
& .\.venv\Scripts\python.exe .\lccs_tracker_job.py --install-task
```

`--remove-task` stops the hourly refresh. Each run replaces the whole file, so
edits a co-author makes in Excel for the web that have not yet synced to the
running computer can be overwritten; OneDrive version history can restore them.

The job stores each row's LCCS `XAWBID` in hidden column C (`LCCS_ID`),
refreshes the LCCS columns (including ETA from `XAWSARVTIME` and Total Cons,
the count of dockets on the MAWB), keeps team-entered columns (offload
details, NDC_Remarks, LH pickup, Status), adds new MAWBs, and keeps
only the refreshed dates (yesterday and today by default), removing older rows.
Region sheets are protected: editors can type only in the team columns and
cannot insert or delete rows (set `TRACKER_SHEET_PASSWORD` in `.env` to require
a password to unprotect). Close the workbook in Excel before a
refresh. Rows typed without an `LCCS_ID`, or an `NDC_Remarks` value not in the
list, stop the refresh without changing the file.

Rows are split into four sheets, `North`, `West`, `East` and `South`, by the
origin hub's LCCS zone (`mhubmst.HUBZONE`, kept in hidden column B). A hub with
no zone goes to an `Unassigned` sheet, which disappears once it is empty. If a
hub's zone changes, its rows and team entries move to the new sheet. A workbook
in the older single `LH Performance Tracker` sheet layout is split on its next
refresh. The dashboard reads all region sheets.

The query also reproduces `Multijoins::skipTestingAWB()` (MAWBs with dockets for
test client `206324` are excluded) and the report's NFO rule (any non-cancelled
`xdockethdr` docket on the MAWB with `XDKTPROMODE = 'NFO'`). As in the PHP
report, the Prime/Lean slot is set only for air MAWBs (`XAWBMODE = 'A'`).
ETA, Total Cons, and LH pickup date/time remain blank because they are not in
LCCS. The query runs against the LCCS read database (`SITE_DB_M`); connection
details are entered on the running computer with
`python lccs_connection.py --configure`. The tracker still needs durable storage
and access control before any tracker data is served by the public dashboard.

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
   The token is entered with hidden input and saved as `DASHBOARD_UPLOAD_TOKEN`
   in the root `.env` file (see [Credentials](#credentials)):

   ```powershell
   & .\.venv\Scripts\python.exe .\sync_workbook_to_dashboard.py --configure --workbook "C:\path\to\the\synced\workbook.xlsx"
   ```

   Configuration also performs an initial upload. The workbook must contain the
   expected dashboard sheets. The local `sync-config.json` keeps only the
   workbook path and endpoint.
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

## Credentials

On this computer every credential lives in the root `.env` file, which Git
ignores. Names are listed in [.env.example](./.env.example):

- `LCCS_DB_*`: the LCCS read-only database login (`lccs_connection.py --configure` writes them)
- `DASHBOARD_UPLOAD_TOKEN`: the dashboard upload token (`sync_workbook_to_dashboard.py --configure` writes it)
- `SMTP_*`, `EMAIL_FROM`: email settings for local runs of `dashboard_delivery.py`

Values already set in the process environment take precedence. GitHub Actions
cannot read this file, so the daily email keeps using repository secrets, and
Render keeps its own `DASHBOARD_UPLOAD_TOKEN` setting.

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

For local runs of `dashboard_delivery.py`, SMTP settings can instead be placed
in the root `.env` file. The file is ignored by Git; keep it on a trusted
computer and never commit or share it. Environment variables already set by
the process take precedence over values in `.env`. GitHub Actions continues to
use the repository secrets above.

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
