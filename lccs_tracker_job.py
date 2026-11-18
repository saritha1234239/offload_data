"""Refresh the LH Performance Tracker workbook from the LCCS database.

LCCS columns are refreshed on every run. Columns the team fills in (offload
details, NDC remarks, LH pickup, Status) are read back from the workbook first
and preserved. The tracker keeps every MAWB departing on or after
TRACKER_START_DATE (16 Nov 2026); each run re-reads that date to today.
"""
import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from env_file import env_value

from lccs_connection import CONFIG_DIRECTORY, LccsConfigError, connect
from lccs_tracker import (
    DEST_REGION_INDEX,
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
TRACKER_START_DATE = date(2026, 11, 16)  # first departure date the tracker holds
SHEET_PASSWORD_ENV = "TRACKER_SHEET_PASSWORD"  # optional, in .env
STATUS_OPTIONS = ("Ontime", "Delayed")
STATUS_FORMULA = '"' + ",".join(STATUS_OPTIONS) + '"'

# Tracker column index (0-based within D:AA) -> manual key understood by map_lccs_row.
# ETA (XAWSARVTIME), Total Cons (docket count) and LH pickup date/time (latest
# linehaul pickup scan) come from LCCS on every refresh.
MANUAL_COLUMNS = {
    15: "offloaded_bags",
    16: "impacted_cons",
    17: "next_flight",
    18: "revised_etd",
    19: "revised_eta",
    20: "ndc_remarks",
    23: "status",
}


class TrackerJobError(RuntimeError):
    pass


class WorkbookLockedError(TrackerJobError):
    """The workbook is open in desktop Excel (or mid-sync), so it cannot be read or replaced."""


# The hourly task is stopped after 30 minutes, so retry for at most 25.
LOCK_RETRY_SECONDS = 300
LOCK_RETRY_ATTEMPTS = 6


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


HEADER_SEARCH_ROWS = 10


OPTIONAL_COLUMNS = {DEST_REGION_INDEX}  # added later; inserted after "Dest" if missing


def _header_key(value):
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def layout(sheet):
    """Return (header_row, columns): the sheet column of each TRACKER_COLUMNS entry.

    Columns are matched by header name (case, spaces and punctuation ignored), so
    editors may move columns or rows above the headers. Repeated names (ETD, ETA)
    are matched in left-to-right order. Optional columns may be None.
    """
    wanted = [_header_key(name) for name in TRACKER_COLUMNS]
    for row in range(1, HEADER_SEARCH_ROWS + 1):
        cells = [(column, _header_key(sheet.cell(row, column).value))
                 for column in range(FIRST_COLUMN, sheet.max_column + 1)]
        used = set()
        columns = []
        for key in wanted:
            match = next((c for c, k in cells if k == key and c not in used), None)
            columns.append(match)
            if match:
                used.add(match)
        if all(c for i, c in enumerate(columns) if i not in OPTIONAL_COLUMNS):
            return row, columns
    raise TrackerJobError(
        f"'{sheet.title}' has no row with the tracker headers (Departure_Date ... Status) "
        f"within the first {HEADER_SEARCH_ROWS} rows."
    )


def header_row(sheet):
    return layout(sheet)[0]


def ensure_columns(sheet):
    """Insert Dest_Region right after Dest on sheets that do not have it yet."""
    row, columns = layout(sheet)
    if columns[DEST_REGION_INDEX] is None:
        position = columns[TRACKER_COLUMNS.index("Dest")] + 1
        sheet.insert_cols(position)
        sheet.cell(row, position).value = TRACKER_COLUMNS[DEST_REGION_INDEX]
        source = sheet.cell(row, position - 1)
        sheet.cell(row, position)._style = source._style


def _check_headers(sheet):
    layout(sheet)
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


def _text_id(source_row):
    value = source_row.get("XAWBID")
    return "" if value is None else str(value).strip()


ROW_TOKEN = "{row}"


def _is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def _formula_template(value, row):
    """Turn same-row references (C56, $U56) into {row} so the formula can move with its row."""
    return re.sub(rf"(\$?[A-Z]{{1,3}}\$?){row}(?!\d)", rf"\g<1>{ROW_TOKEN}", value)


def _formula_for_row(value, row):
    return value.replace(ROW_TOKEN, str(row)) if _is_formula(value) else value


def read_existing(sheet, existing=None):
    """Return {shipment_id: (region, values)} for rows that carry an LCCS_ID."""
    existing = {} if existing is None else existing
    default_region = sheet.title if sheet.title in REGION_SHEETS else ""
    header, columns = layout(sheet)
    for row in range(header + 1, sheet.max_row + 1):
        shipment_id = sheet.cell(row, ID_COLUMN).value
        values = [sheet.cell(row, column).value if column else None for column in columns]
        values = [_formula_template(v, row) if _is_formula(v) else v for v in values]
        if shipment_id in (None, ""):
            # Formulas filled down below the data (e.g. Status) are not typed rows.
            if any(value not in (None, "") and not _is_formula(value) for value in values):
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


NDC_FORMULA = "NDC_Remarks_Options"
NDC_INDEX = TRACKER_COLUMNS.index("NDC_Remarks")
STATUS_INDEX = TRACKER_COLUMNS.index("Status")


def _set_dropdown(sheet, formula, column, first_row, last_row):
    """Point the list dropdown with this formula at one column, creating it if missing."""
    matches = [v for v in sheet.data_validations.dataValidation
               if str(v.formula1 or "").lstrip("=") == formula.lstrip("=")]
    validation = matches[0] if matches else None
    for extra in matches[1:]:
        sheet.data_validations.dataValidation.remove(extra)
    if validation is None:
        validation = DataValidation(type="list", formula1=formula, allow_blank=True)
        sheet.add_data_validation(validation)
    letter = get_column_letter(column)
    validation.sqref = openpyxl.worksheet.cell_range.MultiCellRange(f"{letter}{first_row}:{letter}{last_row}")


def write_rows(sheet, rows):
    header, columns = layout(sheet)
    first_data_row = header + 1
    last_row = max(sheet.max_row, first_data_row)
    # Clear only the tracker columns (and hidden B:C); other columns are left alone.
    owned = [REGION_COLUMN, ID_COLUMN] + [column for column in columns if column]
    for row in range(first_data_row, last_row + 1):
        for column in owned:
            sheet.cell(row, column).value = None
    for offset, (shipment_id, (region, values)) in enumerate(sorted(rows.items(), key=_sort_key)):
        row = first_data_row + offset
        sheet.cell(row, REGION_COLUMN).value = region or None
        sheet.cell(row, ID_COLUMN).value = shipment_id
        for column, value in zip(columns, values):
            if column:
                sheet.cell(row, column).value = None if value in ("", None) else _formula_for_row(value, row)

    sheet.cell(header, REGION_COLUMN).value = REGION_HEADER
    sheet.cell(header, ID_COLUMN).value = ID_HEADER
    sheet.column_dimensions["B"].hidden = True
    sheet.column_dimensions["C"].hidden = True
    end_row = max(MIN_DROPDOWN_ROWS, first_data_row + len(rows) + 500)
    _set_dropdown(sheet, "=" + NDC_FORMULA, columns[NDC_INDEX], first_data_row, end_row)
    _set_dropdown(sheet, STATUS_FORMULA, columns[STATUS_INDEX], first_data_row, end_row)
    present = [column for column in columns if column]
    sheet.auto_filter.ref = (
        f"{get_column_letter(min(present))}{header}:{get_column_letter(max(present))}{end_row}"
    )
    protect_sheet(sheet, len(rows))


def protect_sheet(sheet, row_count):
    """Editors may type only in the team columns; LCCS columns and rows cannot be deleted."""
    header, columns = layout(sheet)
    first_data_row = header + 1
    unlocked = Protection(locked=False)
    for index in MANUAL_COLUMNS:
        column = columns[index]
        sheet.column_dimensions[get_column_letter(column)].protection = unlocked  # empty cells below the data
        for row in range(first_data_row, first_data_row + row_count):
            sheet.cell(row, column).protection = unlocked
    protection = sheet.protection
    protection.sheet = True
    protection.autoFilter = False  # False = allowed
    protection.sort = False
    protection.formatColumns = False
    protection.formatRows = False
    protection.deleteRows = True
    protection.insertRows = True
    protection.deleteColumns = True
    protection.insertColumns = True
    password = env_value(SHEET_PASSWORD_ENV)
    if password:
        protection.password = password


def _new_region_sheet(workbook, title, template):
    """Copy the template layout (title, headers, widths, styles) into an empty sheet."""
    sheet = workbook.copy_worksheet(template)
    sheet.title = title
    header = header_row(template)
    for row in range(header + 1, sheet.max_row + 1):
        for column in range(1, sheet.max_column + 1):
            sheet.cell(row, column).value = None
    return sheet  # dropdowns, filter and protection are applied by write_rows


def remove_title_rows(sheet):
    """Delete everything above the header row so the headers sit in row 1."""
    header = header_row(sheet)
    if header == 1:
        return
    for merged in list(sheet.merged_cells.ranges):
        if merged.min_row < header:
            sheet.unmerge_cells(str(merged))
    sheet.delete_rows(1, header - 1)


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
        ensure_columns(workbook[name])
        remove_title_rows(workbook[name])
        workbook[name].freeze_panes = "D2"  # headers and hidden B:C stay in view
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
        raise WorkbookLockedError(
            f"Could not replace {path.name}; it is open in desktop Excel. Close it, or use Excel for the web."
        ) from exc
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def refresh(workbook_path, from_date, to_date, hub_code=None, source_rows=None):
    workbook_path = Path(workbook_path)
    if not workbook_path.is_file():
        raise TrackerJobError(f"Tracker workbook not found: {workbook_path}")
    try:
        workbook = openpyxl.load_workbook(workbook_path)
    except PermissionError as exc:
        raise WorkbookLockedError(
            f"Could not open {workbook_path.name}; it is open in desktop Excel. Close it, or use Excel for the web."
        ) from exc
    sheets = tracker_sheets(workbook)
    existing = {}
    for sheet in sheets:
        read_existing(sheet, existing)
    if source_rows is None:
        source_rows = fetch_lccs_rows(from_date, to_date, hub_code)
    merged, added, updated = merge(existing, source_rows)
    # If the team uses a Status formula, give new rows the same formula.
    status_formula = next((values[STATUS_INDEX] for _, values in existing.values()
                           if _is_formula(values[STATUS_INDEX])), None)
    if status_formula:
        for _, values in merged.values():
            if values[STATUS_INDEX] in (None, ""):
                values[STATUS_INDEX] = status_formula
    # The tracker keeps every MAWB departing on or after TRACKER_START_DATE. Within
    # the refreshed dates, a full (all-hub) refresh also drops MAWBs LCCS no longer
    # returns (e.g. changed to surface, deleted). Days outside the refreshed dates
    # are left as they are.
    start = TRACKER_START_DATE.isoformat()
    window = (from_date.isoformat(), to_date.isoformat())
    returned = {str(_text_id(row)) for row in source_rows}

    def keep(shipment_id, row):
        day = str(row[1][0] or "")[:10]
        if day < start:
            return False
        if window[0] <= day <= window[1] and hub_code is None and returned:
            return shipment_id in returned  # an empty LCCS result is a fault: keep everything
        return True

    kept = {shipment_id: row for shipment_id, row in merged.items() if keep(shipment_id, row)}
    write_region_sheets(workbook, kept, sheets)
    save_atomically(workbook, workbook_path)
    return {"added": added, "updated": updated, "removed": len(merged) - len(kept), "total": len(kept)}


def refresh_with_retry(workbook_path, from_date, to_date, hub_code=None,
                       attempts=LOCK_RETRY_ATTEMPTS, wait_seconds=LOCK_RETRY_SECONDS, sleep=time.sleep):
    """Refresh, waiting and retrying while the workbook is open in desktop Excel."""
    for attempt in range(1, attempts + 1):
        try:
            return refresh(workbook_path, from_date, to_date, hub_code)
        except WorkbookLockedError as exc:
            if attempt == attempts:
                raise
            _log(f"Workbook is locked (attempt {attempt} of {attempts}); retrying in "
                 f"{wait_seconds // 60} minutes. {exc}")
            sleep(wait_seconds)


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
    parser.add_argument("--from-date", type=date.fromisoformat,
                        help=f"YYYY-MM-DD (default: {TRACKER_START_DATE}, the tracker start date).")
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
    from_date = args.from_date or TRACKER_START_DATE
    if from_date > to_date:
        parser.error("--from-date must be on or before --to-date")
    workbook = args.workbook or saved_workbook() or DEFAULT_WORKBOOK
    try:
        result = refresh_with_retry(workbook, from_date, to_date, args.hub)
    except (OSError, LccsConfigError, TrackerDataError, TrackerJobError) as exc:
        _report(f"Tracker refresh failed: {exc}", error=True)
        return 1
    except Exception as exc:  # database driver errors
        _report(f"Tracker refresh failed: {type(exc).__name__}: {exc}", error=True)
        return 1
    _report(
        f"Tracker updated for {from_date} to {to_date}: "
        f"{result['added']} added, {result['updated']} refreshed, "
        f"{result['removed']} rows removed, {result['total']} rows total."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
