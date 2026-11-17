"""Refresh the LH Performance Tracker workbook from the LCCS database.

LCCS columns are refreshed on every run. Columns the team fills in (ETA, Total
Cons, offload details, NDC remarks, LH pickup, Status) are read back from the
workbook first and preserved. Rows outside the refreshed date window are left
unchanged, so history is never deleted.
"""
import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

from lccs_connection import CONFIG_DIRECTORY, LccsConfigError, connect
from lccs_tracker import (
    LCCS_TRACKER_QUERY,
    NDC_REMARK_OPTIONS,
    TESTING_CLIENT_CODE,
    TRACKER_COLUMNS,
    TrackerDataError,
    map_lccs_rows,
)


DEFAULT_WORKBOOK = Path(__file__).resolve().parent / "LH Performance Tracker.xlsx"
JOB_CONFIG_PATH = CONFIG_DIRECTORY / "lccs-tracker.json"
LOG_PATH = CONFIG_DIRECTORY / "lccs-tracker.log"
TASK_NAME = "LHPerformanceTrackerRefresh"
SHEET_NAME = "LH Performance Tracker"  # original single-sheet layout; migrated to region sheets
REGION_SHEETS = ("North", "West", "East", "South")
UNASSIGNED_SHEET = "Unassigned"  # only created if LCCS returns a hub outside the four regions
TRACKER_SHEETS = (SHEET_NAME,) + REGION_SHEETS + (UNASSIGNED_SHEET,)
HEADER_ROW = 4
FIRST_DATA_ROW = 5
FIRST_COLUMN = 4  # D
ID_COLUMN = 3  # C, hidden
ID_HEADER = "LCCS_ID"
REGION_COLUMN = 2  # B, hidden
REGION_HEADER = "Org_Region"
MIN_DROPDOWN_ROWS = 5000

# Tracker column index (0-based within D:AA) -> manual key understood by map_lccs_row.
MANUAL_COLUMNS = {
    10: "scheduled_eta_override",  # ETA
    14: "total_cons",
    15: "offloaded_bags",
    16: "impacted_cons",
    17: "next_flight",
    18: "revised_etd",
    19: "revised_eta",
    20: "ndc_remarks",
    21: "lh_pickup_date",
    22: "lh_pickup_time",
    23: "status",
}


class TrackerJobError(RuntimeError):
    pass


def mysql_query(query):
    """Convert the :name placeholders used in LCCS_TRACKER_QUERY to PyMySQL style."""
    return re.sub(r"(?<!:):([a-z_]+)", r"%(\1)s", query.replace("%", "%%"))


def fetch_lccs_rows(from_date, to_date, hub_code=None, connection=None):
    owns_connection = connection is None
    connection = connection or connect()
    try:
        with connection.cursor() as cursor:
            cursor.execute(mysql_query(LCCS_TRACKER_QUERY), {
                "from_date": from_date.isoformat(),
                "to_date": to_date.isoformat(),
                "hub_code": hub_code,
                "testing_client_code": TESTING_CLIENT_CODE,
            })
            return list(cursor.fetchall())
    finally:
        if owns_connection:
            connection.close()


def _check_headers(sheet):
    headers = [sheet.cell(HEADER_ROW, FIRST_COLUMN + i).value for i in range(len(TRACKER_COLUMNS))]
    if headers != list(TRACKER_COLUMNS):
        raise TrackerJobError(f"'{sheet.title}' headers in D4:AA4 do not match the expected 24 columns.")
    return sheet


def tracker_sheets(workbook):
    """Every tracker sheet present: the original single sheet and/or the region sheets."""
    sheets = [_check_headers(workbook[name]) for name in TRACKER_SHEETS if name in workbook.sheetnames]
    if not sheets:
        raise TrackerJobError(
            f"Workbook has no tracker sheet ({', '.join(REGION_SHEETS)} or '{SHEET_NAME}')."
        )
    return sheets


def region_sheet_name(region):
    region = str(region or "").strip().title()
    return region if region in REGION_SHEETS else UNASSIGNED_SHEET


def read_existing(sheet, existing=None):
    """Return {shipment_id: (region, values)} for rows that carry an LCCS_ID."""
    existing = {} if existing is None else existing
    default_region = sheet.title if sheet.title in REGION_SHEETS else ""
    for row in range(FIRST_DATA_ROW, sheet.max_row + 1):
        shipment_id = sheet.cell(row, ID_COLUMN).value
        values = [sheet.cell(row, FIRST_COLUMN + i).value for i in range(len(TRACKER_COLUMNS))]
        if shipment_id in (None, ""):
            if any(value not in (None, "") for value in values):
                raise TrackerJobError(
                    f"'{sheet.title}' row {row} has data but no {ID_HEADER}; "
                    "add rows only through the refresh job."
                )
            continue
        shipment_id = str(shipment_id)
        if shipment_id in existing:
            raise TrackerJobError(
                f"{ID_HEADER} {shipment_id} appears more than once ('{sheet.title}' row {row}); "
                "delete the duplicate row before refreshing."
            )
        existing[shipment_id] = (sheet.cell(row, REGION_COLUMN).value or default_region, values)
    return existing


def manual_values(values):
    manual = {}
    for index, key in MANUAL_COLUMNS.items():
        value = values[index]
        if value not in (None, ""):
            manual[key] = value
    return manual


def merge(existing, source_rows):
    manual_by_id = {shipment_id: manual_values(values) for shipment_id, (_, values) in existing.items()}
    for shipment_id, manual in manual_by_id.items():
        remark = manual.get("ndc_remarks")
        if remark and remark not in NDC_REMARK_OPTIONS:
            raise TrackerJobError(
                f"NDC_Remarks for LCCS_ID {shipment_id} is not one of the dropdown options; "
                "fix it in the workbook before refreshing."
            )
    merged = dict(existing)
    added = updated = 0
    for record in map_lccs_rows(source_rows, manual_by_id):
        if record["shipment_id"] in merged:
            updated += 1
        else:
            added += 1
        previous_region = merged.get(record["shipment_id"], ("", None))[0]
        merged[record["shipment_id"]] = (record.get("region") or previous_region, record["values"])
    return merged, added, updated


def _sort_key(item):
    shipment_id, (_, values) = item
    return (str(values[0] or ""), str(values[9] or ""), str(values[4] or ""), shipment_id)


def write_rows(sheet, rows):
    last_row = max(sheet.max_row, FIRST_DATA_ROW)
    for row in range(FIRST_DATA_ROW, last_row + 1):
        for column in range(REGION_COLUMN, FIRST_COLUMN + len(TRACKER_COLUMNS)):
            sheet.cell(row, column).value = None
    for offset, (shipment_id, (region, values)) in enumerate(sorted(rows.items(), key=_sort_key)):
        row = FIRST_DATA_ROW + offset
        sheet.cell(row, REGION_COLUMN).value = region or None
        sheet.cell(row, ID_COLUMN).value = shipment_id
        for index, value in enumerate(values):
            sheet.cell(row, FIRST_COLUMN + index).value = None if value == "" else value

    sheet.cell(HEADER_ROW, REGION_COLUMN).value = REGION_HEADER
    sheet.cell(HEADER_ROW, ID_COLUMN).value = ID_HEADER
    sheet.column_dimensions["B"].hidden = True
    sheet.column_dimensions["C"].hidden = True
    end_row = max(MIN_DROPDOWN_ROWS, FIRST_DATA_ROW + len(rows) + 500)
    for validation in sheet.data_validations.dataValidation:
        if "X" in str(validation.sqref).upper():
            validation.sqref = openpyxl.worksheet.cell_range.MultiCellRange(f"X{FIRST_DATA_ROW}:X{end_row}")
    sheet.auto_filter.ref = f"D{HEADER_ROW}:AA{end_row}"


def _new_region_sheet(workbook, title, template):
    """Copy the template layout (title, headers, widths, styles) into an empty sheet."""
    sheet = workbook.copy_worksheet(template)
    sheet.title = title
    for row in range(FIRST_DATA_ROW, sheet.max_row + 1):
        for column in range(1, sheet.max_column + 1):
            sheet.cell(row, column).value = None
    sheet.cell(2, FIRST_COLUMN).value = f"LH Performance Tracker - {title}"
    sheet.freeze_panes = template.freeze_panes
    validation = DataValidation(type="list", formula1="=NDC_Remarks_Options", allow_blank=True)
    validation.add(f"X{FIRST_DATA_ROW}:X{MIN_DROPDOWN_ROWS}")
    sheet.add_data_validation(validation)
    return sheet


def write_region_sheets(workbook, rows, sheets):
    """Split rows into North/West/East/South sheets and retire the single-sheet layout."""
    template = sheets[0]
    by_sheet = {name: {} for name in REGION_SHEETS}
    for shipment_id, (region, values) in rows.items():
        by_sheet.setdefault(region_sheet_name(region), {})[shipment_id] = (region, values)
    for name, sheet_rows in by_sheet.items():
        if name == UNASSIGNED_SHEET and not sheet_rows and name not in workbook.sheetnames:
            continue
        if name not in workbook.sheetnames:
            _new_region_sheet(workbook, name, template)
        write_rows(workbook[name], sheet_rows)
    for name in (SHEET_NAME, UNASSIGNED_SHEET):
        if name in workbook.sheetnames and not (name == UNASSIGNED_SHEET and by_sheet.get(name)):
            workbook.remove(workbook[name])
    order = [workbook[name] for name in REGION_SHEETS + (UNASSIGNED_SHEET,) if name in workbook.sheetnames]
    workbook._sheets = order + [sheet for sheet in workbook.worksheets if sheet not in order]
    workbook.active = 0
    for sheet in workbook.worksheets:
        sheet.sheet_view.tabSelected = sheet is order[0]


def save_atomically(workbook, path):
    path = Path(path)
    descriptor, temporary = tempfile.mkstemp(prefix=".lh-tracker-", suffix=".xlsx", dir=path.parent)
    os.close(descriptor)
    try:
        workbook.save(temporary)
        os.replace(temporary, path)
    except PermissionError as exc:
        raise TrackerJobError(
            f"Could not replace {path.name}; close it in Excel and run the refresh again."
        ) from exc
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def refresh(workbook_path, from_date, to_date, hub_code=None, source_rows=None):
    workbook_path = Path(workbook_path)
    if not workbook_path.is_file():
        raise TrackerJobError(f"Tracker workbook not found: {workbook_path}")
    workbook = openpyxl.load_workbook(workbook_path)
    sheets = tracker_sheets(workbook)
    existing = {}
    for sheet in sheets:
        read_existing(sheet, existing)
    if source_rows is None:
        source_rows = fetch_lccs_rows(from_date, to_date, hub_code)
    merged, added, updated = merge(existing, source_rows)
    write_region_sheets(workbook, merged, sheets)
    save_atomically(workbook, workbook_path)
    return {"added": added, "updated": updated, "total": len(merged)}


def saved_workbook():
    try:
        return Path(json.loads(JOB_CONFIG_PATH.read_text(encoding="utf-8"))["workbookPath"])
    except (OSError, KeyError, ValueError):
        return None


def set_workbook(path):
    path = Path(path).expanduser().resolve()
    if not path.is_file():
        raise TrackerJobError(f"Tracker workbook not found: {path}")
    tracker_sheets(openpyxl.load_workbook(path))
    CONFIG_DIRECTORY.mkdir(parents=True, exist_ok=True)
    JOB_CONFIG_PATH.write_text(json.dumps({"workbookPath": str(path)}, indent=2), encoding="utf-8")
    print(f"Tracker workbook saved: {path}")


def allow_task_on_battery(task_name):
    """schtasks cannot set power options; new tasks otherwise skip runs on battery."""
    script = (
        f"$t = Get-ScheduledTask -TaskName '{task_name}'; "
        "$t.Settings.DisallowStartIfOnBatteries = $false; "
        "$t.Settings.StopIfGoingOnBatteries = $false; "
        "$t.Settings.StartWhenAvailable = $true; "
        "$t.Settings.ExecutionTimeLimit = 'PT30M'; "
        "Set-ScheduledTask -InputObject $t | Out-Null"
    )
    subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
        check=True, capture_output=True, text=True,
    )


def install_task():
    if saved_workbook() is None:
        raise TrackerJobError("Run --set-workbook first.")
    interpreter = Path(sys.executable)
    windowless = interpreter.with_name("pythonw.exe")
    if windowless.is_file():
        interpreter = windowless
    action = f'"{interpreter}" "{Path(__file__).resolve()}"'
    if len(action) > 261:
        raise TrackerJobError("The scheduled task command is too long for Windows Task Scheduler.")
    subprocess.run(
        ["schtasks.exe", "/Create", "/SC", "HOURLY", "/MO", "1", "/TN", TASK_NAME,
         "/TR", action, "/F", "/IT", "/RL", "LIMITED"],
        check=True, capture_output=True, text=True,
    )
    allow_task_on_battery(TASK_NAME)
    print(f"Scheduled '{TASK_NAME}' every hour while you are signed in. Log: {LOG_PATH}")


def remove_task():
    subprocess.run(["schtasks.exe", "/Delete", "/TN", TASK_NAME, "/F"],
                   check=True, capture_output=True, text=True)
    print(f"Scheduled task '{TASK_NAME}' removed.")


def _log(message):
    try:
        CONFIG_DIRECTORY.mkdir(parents=True, exist_ok=True)
        with LOG_PATH.open("a", encoding="utf-8") as log:
            log.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {message}\n")
    except OSError:
        pass


def _report(message, error=False):
    _log(("ERROR " if error else "") + message)
    if sys.stdout is not None:
        print(message, file=sys.stderr if error else sys.stdout)


def main():
    parser = argparse.ArgumentParser(description="Refresh the LH Performance Tracker from LCCS.")
    parser.add_argument("--workbook", help="Tracker workbook to update (default: the saved one).")
    parser.add_argument("--from-date", type=date.fromisoformat, help="YYYY-MM-DD (default: yesterday).")
    parser.add_argument("--to-date", type=date.fromisoformat, help="YYYY-MM-DD (default: today).")
    parser.add_argument("--hub", help="Optional origin hub code; default is all hubs.")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--set-workbook", metavar="PATH", help="Save the shared tracker workbook location.")
    mode.add_argument("--install-task", action="store_true", help="Refresh automatically every hour.")
    mode.add_argument("--remove-task", action="store_true", help="Stop the hourly refresh.")
    args = parser.parse_args()
    try:
        if args.set_workbook:
            set_workbook(args.set_workbook)
            return 0
        if args.install_task:
            install_task()
            return 0
        if args.remove_task:
            remove_task()
            return 0
    except (OSError, subprocess.CalledProcessError, TrackerJobError) as exc:
        _report(f"Tracker setup failed: {exc}", error=True)
        return 1

    to_date = args.to_date or date.today()
    from_date = args.from_date or (to_date - timedelta(days=1))
    if from_date > to_date:
        parser.error("--from-date must be on or before --to-date")
    workbook = args.workbook or saved_workbook() or DEFAULT_WORKBOOK
    try:
        result = refresh(workbook, from_date, to_date, args.hub)
    except (OSError, LccsConfigError, TrackerDataError, TrackerJobError) as exc:
        _report(f"Tracker refresh failed: {exc}", error=True)
        return 1
    except Exception as exc:  # database driver errors
        _report(f"Tracker refresh failed: {type(exc).__name__}: {exc}", error=True)
        return 1
    _report(
        f"Tracker updated for {from_date} to {to_date}: "
        f"{result['added']} added, {result['updated']} refreshed, {result['total']} rows total."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
