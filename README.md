# Offload Performance Dashboard

## Automatic workbook updates

Once configured, the hosted service polls a sanitized Excel workbook every five
minutes by default. Open dashboard tabs reload the exported data at most one minute
later. Until a source link is configured, the dashboard continues serving its
bundled snapshot. The source must be available as an anonymous, read-only HTTPS
download link; private SharePoint links that redirect to sign-in cannot be fetched
by Render.

### Prepare a sanitized workbook

Keep the original tracker private. Create a separate workbook containing only the
four dashboard tabs, with their original headers and names. The Office Scripts in
`office-scripts/` export those tabs' values from the original and update the
sanitized copy.

In Excel for the web, add `ExportDashboardSheets.ts` to the original tracker and
`UpdateDashboardCopy.ts` to the sanitized workbook. In Power Automate, create a
flow triggered when the original tracker is modified:

1. Run `ExportDashboardSheets` on the original workbook.
2. Run `UpdateDashboardCopy` on the sanitized workbook.
3. Pass the first action's `result` output to the second action's `serializedSheets`
   parameter.

Have the flow watch only the original file so updates to the sanitized copy do not
trigger another run. Create an **Anyone with the link — Can view** link for the
sanitized copy and confirm it downloads without signing in.

### Connect Render

Set `WORKBOOK_DOWNLOAD_URL` in the Render service's environment to the sanitized
workbook's HTTPS download link. Do not put the link in source control. The default
refresh interval is 300 seconds; set `WORKBOOK_REFRESH_SECONDS` to change it (the
minimum is 60 seconds). Redeploy the service after changing environment variables.

`/api/health` reports whether a source is configured, the latest successful refresh,
and any refresh error without returning the source URL. If a download or workbook
validation fails, the service keeps the last successful dashboard data.

The dashboard and linked workbook are public. Include only information approved
for public viewing in the sanitized copy.