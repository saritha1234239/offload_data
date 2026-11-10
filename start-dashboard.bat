@echo off
setlocal
cd /d "%~dp0"

REM Refresh workbook data once before the dashboard opens so the first view is current.
powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File "%~dp0refresh-dashboard-data.ps1" -Once
if errorlevel 1 (
  echo Dashboard data could not be refreshed. Check the workbook path and Excel availability.
  pause
  exit /b 1
)

REM Open the dashboard without requiring a manual refresh step.
start "" "%~dp0dashboard.html"

REM Continue refreshing workbook data in the background while the dashboard is open.
start "" powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File "%~dp0refresh-dashboard-data.ps1" -IntervalSeconds 60
