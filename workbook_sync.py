import hashlib
import json
import math
import os
import re
import tempfile
import threading
from collections import OrderedDict
from datetime import date, datetime, time
from io import BytesIO
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlparse, urlsplit, urlunsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener, urlopen
from zipfile import is_zipfile

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel, to_excel


EXPECTED_SHEETS = ("Daily-North-new", "Daily-South-new", "Daily-West-new", "Daily-East-new")
MAX_WORKBOOK_BYTES = 30 * 1024 * 1024
DATE_FORMATS = (
    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
    "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y",
    "%d-%m-%y", "%d/%m/%y", "%d.%m.%y",
    "%m-%d-%Y", "%m/%d/%Y", "%m.%d.%Y",
    "%m-%d-%y", "%m/%d/%y", "%m.%d.%y",
)
HEADER_TOKENS = {
    "departure_date", "date", "org", "dest", "mawb", "vendor_name", "flight_number",
    "etd", "eta", "bags", "cd_weight", "total_cons", "offloaded_bags",
    "impacted_cons", "next_flight", "remarks", "ndc_remark", "ndc_remarks",
}


class WorkbookSyncError(RuntimeError):
    pass


def _normalize_name(value):
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def _header_token(value, column):
    token = re.sub(r"[^\w]+", "_", str(value or "").strip(), flags=re.UNICODE).strip("_")
    return token or f"Column_{column}"


def _normalize_header(value):
    return re.sub(r"[^a-z0-9]+", "_", str(value or "").strip().lower()).strip("_")


def _date_variants(value, region):
    if value is None or isinstance(value, bool):
        return []
    if isinstance(value, datetime):
        return [value.date().isoformat()]
    if isinstance(value, date):
        return [value.isoformat()]
    if isinstance(value, (int, float)):
        if 1000 <= value <= 2958465:
            try:
                parsed = from_excel(value)
                if isinstance(parsed, datetime):
                    parsed = parsed.date()
                if isinstance(parsed, date):
                    return [parsed.isoformat()]
            except (OverflowError, TypeError, ValueError):
                pass
        return []

    text = str(value).strip().replace("\u200b", "").replace("\ufeff", "")
    if not text or re.fullmatch(r"[A-Za-z\s_/-]+", text):
        return []
    if _normalize_header(text) in HEADER_TOKENS:
        return []

    iso_datetime = re.fullmatch(r"(\d{4}-\d{2}-\d{2})[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?", text)
    if iso_datetime:
        text = iso_datetime.group(1)
    candidates = [text]
    ambiguous = re.fullmatch(r"(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})", text)
    if ambiguous:
        first, second, year = ambiguous.groups()
        candidates.extend((f"{first}/{second}/{year}", f"{second}/{first}/{year}"))

    month_first = "north" in region.lower()
    mdy = [fmt for fmt in DATE_FORMATS if fmt.startswith("%m")]
    dmy = [fmt for fmt in DATE_FORMATS if fmt.startswith("%d")]
    ordered_formats = (
        ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"] + (mdy + dmy if month_first else dmy + mdy)
    )

    variants = []
    for candidate in candidates:
        for fmt in ordered_formats:
            try:
                parsed = datetime.strptime(candidate, fmt).date().isoformat()
            except ValueError:
                continue
            if parsed not in variants:
                variants.append(parsed)
    return variants


def _text(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value).strip()


def _number(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, str):
        value = value.strip()
        if not value or value.upper() == "NA" or value == "�":
            return None
        value = value.replace(",", "")
        try:
            value = float(value)
        except ValueError:
            return None
    if isinstance(value, (int, float)) and math.isfinite(value):
        return value
    return None


def _time_value(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)) and math.isfinite(value):
        return float(value)
    if isinstance(value, time):
        return (value.hour * 3600 + value.minute * 60 + value.second) / 86400
    if isinstance(value, datetime):
        return (value.hour * 3600 + value.minute * 60 + value.second) / 86400
    parsed = None
    try:
        parsed = datetime.fromisoformat(str(value).strip())
    except ValueError:
        try:
            parsed = time.fromisoformat(str(value).strip())
        except ValueError:
            return None
    return (parsed.hour * 3600 + parsed.minute * 60 + parsed.second) / 86400


def _source_value(value):
    if isinstance(value, datetime):
        return to_excel(value)
    if isinstance(value, date):
        return to_excel(value)
    if isinstance(value, time):
        return (value.hour * 3600 + value.minute * 60 + value.second) / 86400
    if isinstance(value, float) and not math.isfinite(value):
        return None
    return value


def _first_column(headers, aliases):
    for alias in aliases:
        if alias in headers:
            return headers[alias][0]
    return None


def _is_header_row(values):
    matches = 0
    for value in values[:20]:
        token = _normalize_header(value)
        if token in HEADER_TOKENS:
            matches += 1
            if matches >= 3:
                return True
    return False


TRACKER_SHEET = "LH Performance Tracker"
TRACKER_REGION_SHEETS = ("North", "West", "East", "South", "Unassigned")
TRACKER_HEADER_ROW = 4
TRACKER_HEADER_SEARCH_ROWS = 10
TRACKER_FIRST_COLUMN = 4  # D
TRACKER_HEADERS = (
    "Departure_Date", "Org", "Dest", "Lane", "Mawb(with prefix)", "Vendor_name",
    "NFO_Flag", "Direct / Console", "Flight_Number", "ETD", "ETA", "Type as per ETD",
    "Bags", "CD weight", "Total Cons", "Offloaded bags", "Impacted cons",
    "Next_flight", "ETD", "ETA", "NDC_Remarks", "LH_Pickup_date", "LH_Pickup_time",
    "Status",
)


def _column_letter(number):
    letter = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter


def _parse_lccs_tracker(sheet):
    """Parse the LCCS-built LH Performance Tracker (headers in D4:AA4).

    Every tracker row is a dispatched air MAWB from LCCS, and the team records
    offloads as exceptions, so a blank Offloaded bags cell counts as 0.
    """
    # Headers are in row 4 of the template, but editors may delete rows above them.
    width = TRACKER_FIRST_COLUMN - 1 + len(TRACKER_HEADERS)
    rows = sheet.iter_rows(min_row=1, values_only=True)
    header_number = None
    for number, row in enumerate(rows, start=1):
        row = tuple(row) + (None,) * max(0, width - len(row))
        if tuple(_text(value) for value in row[TRACKER_FIRST_COLUMN - 1:width]) == TRACKER_HEADERS:
            header_number = number
            break
        if number >= TRACKER_HEADER_SEARCH_ROWS:
            break
    if header_number is None:
        raise WorkbookSyncError(
            f"'{sheet.title}' has no row with the LH Performance Tracker headers in columns D:AA."
        )

    columns = []
    for index, header in enumerate(TRACKER_HEADERS):
        letter = _column_letter(TRACKER_FIRST_COLUMN + index)
        key = f"tracker_{letter}_{_header_token(header, letter)}"
        columns.append({"key": key, "label": f"{header} ({letter})"})

    default_region = sheet.title if sheet.title in TRACKER_REGION_SHEETS[:4] else "Unspecified"
    records = []
    for row_number, row in enumerate(rows, start=header_number + 1):
        row = tuple(row) + (None,) * max(0, width - len(row))
        values = row[TRACKER_FIRST_COLUMN - 1:width]
        if not any(_text(value) for value in values):
            continue
        # hidden column B, Org_Region; region sheets fall back to their own name
        region = _text(row[1]).title() or default_region
        date_variants = _date_variants(values[0], region)
        origin, destination = _text(values[1]), _text(values[2])
        bags = _number(values[12])
        offloaded = _number(values[15])
        if offloaded is None and bags is not None:
            offloaded = 0
        records.append({
            "sourceSheet": sheet.title,
            "sourceRow": row_number,
            "recordType": "Shipment",
            "date": date_variants[0] if date_variants else "",
            "dateVariants": date_variants,
            "region": region,
            "origin": origin,
            "destination": destination,
            "lane": _text(values[3]) or (f"{origin}-{destination}" if origin and destination else ""),
            "mawb": _text(values[4]),
            "vendor": _text(values[5]),
            "flight": _text(values[8]),
            "etd": _time_value(values[9]),
            "eta": _time_value(values[10]),
            "bags": bags,
            "cdWeight": _number(values[13]),
            "totalCons": _number(values[14]),
            "offloadedBags": offloaded,
            "impactedCons": _number(values[16]),
            "nextFlight": _text(values[17]),
            "nextEtd": _time_value(values[18]),
            "nextEta": _time_value(values[19]),
            "remarks": _text(values[23]),
            "ndcRemarks": _text(values[20]),
            "sourceValues": {
                column["key"]: _source_value(value) for column, value in zip(columns, values)
            },
        })
    dated = [record["date"] for record in records if record["date"]]
    source = {
        "name": sheet.title,
        "rows": len(records),
        "shipments": len(records),
        "worksheetDetails": 0,
        "undated": len(records) - len(dated),
        "minDate": min(dated) if dated else "",
        "maxDate": max(dated) if dated else "",
    }
    return {"columns": columns, "records": records, "sources": [source]}


def parse_workbook(workbook_bytes):
    try:
        workbook = load_workbook(BytesIO(workbook_bytes), read_only=True, data_only=True)
    except Exception as exc:
        raise WorkbookSyncError("The download is not a readable Excel workbook.") from exc

    sheets = {_normalize_name(sheet.title): sheet for sheet in workbook.worksheets}
    trackers = [
        sheets[_normalize_name(name)]
        for name in (TRACKER_SHEET,) + TRACKER_REGION_SHEETS
        if _normalize_name(name) in sheets
    ]
    if trackers and not any(_normalize_name(name) in sheets for name in EXPECTED_SHEETS):
        data = {"columns": [], "records": [], "sources": []}
        try:
            for tracker in trackers:
                parsed = _parse_lccs_tracker(tracker)
                data["columns"] = parsed["columns"]
                data["records"].extend(parsed["records"])
                data["sources"].extend(parsed["sources"])
        finally:
            workbook.close()
        if not data["records"]:
            raise WorkbookSyncError("The tracker workbook contains no dashboard rows.")
        return data
    records = []
    columns = OrderedDict()
    sources = []
    try:
        for expected in EXPECTED_SHEETS:
            sheet = sheets.get(_normalize_name(expected))
            if sheet is None:
                raise WorkbookSyncError(f"Required source sheet '{expected}' was not found.")

            rows = sheet.iter_rows(values_only=True)
            raw_headers = next(rows, None)
            if not raw_headers:
                raise WorkbookSyncError(f"Source sheet '{sheet.title}' is empty.")
            headers = [_text(value) for value in raw_headers]
            header_columns = {}
            source_columns = []
            for index, header in enumerate(headers, 1):
                letter = ""
                column_number = index
                while column_number:
                    column_number, remainder = divmod(column_number - 1, 26)
                    letter = chr(65 + remainder) + letter
                token = _header_token(header, letter)
                key = f"raw_{index}_{token}"
                label = f"{header} ({letter})" if header else f"Column {letter} ({letter})"
                columns.setdefault(key, {"key": key, "label": label})
                source_columns.append(key)
                normalized = _normalize_header(header)
                if normalized:
                    header_columns.setdefault(normalized, []).append(index - 1)

            date_column = _first_column(header_columns, ("date", "departure_date"))
            origin_column = _first_column(header_columns, ("org", "origin"))
            destination_column = _first_column(header_columns, ("dest", "destination"))
            lane_column = _first_column(header_columns, ("lane",))
            mawb_column = _first_column(header_columns, ("mawb",))
            vendor_column = _first_column(header_columns, ("vendor_name", "vendor"))
            flight_column = _first_column(header_columns, ("flight_number", "flight"))
            bags_column = _first_column(header_columns, ("bags",))
            weight_column = _first_column(header_columns, ("cd_weight",))
            total_cons_column = _first_column(header_columns, ("total_cons",))
            offloaded_column = _first_column(header_columns, ("offloaded_bags",))
            impacted_column = _first_column(header_columns, ("impacted_cons",))
            next_flight_column = _first_column(header_columns, ("next_flight",))
            remarks_column = _first_column(header_columns, ("remarks",))
            ndc_remarks_column = _first_column(header_columns, ("ndc_remark", "ndc_remarks"))
            etd_columns = header_columns.get("etd", [])
            eta_columns = header_columns.get("eta", [])
            required = (date_column, origin_column, destination_column, bags_column, offloaded_column)
            if any(column is None for column in required):
                raise WorkbookSyncError(f"Required shipment columns are missing from '{sheet.title}'.")

            region = next((name for name in ("North", "South", "West", "East") if name.lower() in sheet.title.lower()), "Unspecified")
            sheet_records = []
            dated = []
            for row_number, row in enumerate(rows, start=2):
                if not any(value is not None and _text(value) for value in row):
                    continue
                if _is_header_row(row):
                    continue

                def cell(column):
                    return row[column] if column is not None and column < len(row) else None

                date_variants = _date_variants(cell(date_column), region)
                source_values = {
                    key: _source_value(row[index] if index < len(row) else None)
                    for index, key in enumerate(source_columns)
                }
                origin = _text(cell(origin_column))
                destination = _text(cell(destination_column))
                mawb = _text(cell(mawb_column))
                vendor = _text(cell(vendor_column))
                flight = _text(cell(flight_column))
                bags = _number(cell(bags_column))
                cd_weight = _number(cell(weight_column))
                total_cons = _number(cell(total_cons_column))
                offloaded_bags = _number(cell(offloaded_column))
                impacted_cons = _number(cell(impacted_column))
                next_flight = _text(cell(next_flight_column))
                record_type = (
                    "Shipment"
                    if mawb or vendor or flight or any(
                        value is not None
                        for value in (bags, cd_weight, total_cons, offloaded_bags)
                    )
                    else "Worksheet detail"
                )
                lane = (
                    _text(cell(lane_column))
                    if lane_column is not None
                    else f"{origin}-{destination}" if origin and destination else ""
                )
                record = {
                    "sourceSheet": sheet.title,
                    "sourceRow": row_number,
                    "recordType": record_type,
                    "date": date_variants[0] if date_variants else "",
                    "dateVariants": date_variants,
                    "region": region,
                    "origin": origin,
                    "destination": destination,
                    "lane": lane,
                    "mawb": mawb,
                    "vendor": vendor,
                    "flight": flight,
                    "etd": _time_value(cell(etd_columns[0] if etd_columns else None)),
                    "eta": _time_value(cell(eta_columns[0] if eta_columns else None)),
                    "bags": bags,
                    "cdWeight": cd_weight,
                    "totalCons": total_cons,
                    "offloadedBags": offloaded_bags,
                    "impactedCons": impacted_cons,
                    "nextFlight": next_flight,
                    "nextEtd": _time_value(cell(etd_columns[1] if len(etd_columns) > 1 else None)),
                    "nextEta": _time_value(cell(eta_columns[1] if len(eta_columns) > 1 else None)),
                    "remarks": _text(cell(remarks_column)),
                    "ndcRemarks": _text(cell(ndc_remarks_column)),
                    "sourceValues": source_values,
                }
                records.append(record)
                sheet_records.append(record)
                if record["date"]:
                    dated.append(record["date"])

            sources.append({
                "name": sheet.title,
                "rows": len(sheet_records),
                "shipments": sum(record["recordType"] == "Shipment" for record in sheet_records),
                "worksheetDetails": sum(record["recordType"] != "Shipment" for record in sheet_records),
                "undated": sum(not record["date"] for record in sheet_records),
                "minDate": min(dated) if dated else "",
                "maxDate": max(dated) if dated else "",
            })
    finally:
        workbook.close()

    if not records:
        raise WorkbookSyncError("The source workbook contains no dashboard rows.")
    return {"columns": list(columns.values()), "records": records, "sources": sources}


def workbook_data_script(data):
    records_json = json.dumps(data["records"], ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    version = hashlib.sha256(records_json.encode("utf-8")).hexdigest()
    columns_json = json.dumps(data["columns"], ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    sources_json = json.dumps(data["sources"], ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    content = "\n".join((
        f'window.offloadWorkbookUpdatedAt = "{version}";',
        f"window.offloadWorkbookColumns = {columns_json};",
        f"window.offloadWorkbookRecords = {records_json};",
        f"window.offloadWorkbookSources = {sources_json};",
        "window.offloadWorkbookLists = {};",
        "",
    ))
    return version, content


def fetch_workbook(url):
    parsed = urlparse(url)
    if parsed.scheme != "https" or not parsed.netloc:
        raise WorkbookSyncError("WORKBOOK_DOWNLOAD_URL must be an HTTPS download link.")
    parts = urlsplit(url)
    download_match = re.search(r"(^|&)download=[^&]*", parts.query, flags=re.IGNORECASE)
    if download_match:
        download_query = (
            f"{parts.query[:download_match.start()]}"
            f"{download_match.group(1)}download=1"
            f"{parts.query[download_match.end():]}"
        )
    else:
        separator = "&" if parts.query else ""
        download_query = f"{parts.query}{separator}download=1"
    download_url = urlunsplit((
        parts.scheme, parts.netloc, parts.path, download_query, parts.fragment,
    ))
    candidates = list(dict.fromkeys((download_url, url)))
    login_hosts = {"login.microsoftonline.com", "login.windows.net", "login.live.com"}
    failures = []

    for candidate in candidates:
        request = Request(candidate, headers={
            "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/octet-stream, */*",
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0.0.0 Safari/537.36"
            ),
        })
        try:
            with urlopen(request, timeout=60) as response:
                final_url = urlparse(response.geturl())
                if final_url.scheme != "https" or (final_url.hostname or "").lower() in login_hosts:
                    failures.append("sign-in redirect")
                    continue
                content_length = response.headers.get("Content-Length")
                if content_length and int(content_length) > MAX_WORKBOOK_BYTES:
                    raise WorkbookSyncError("The source workbook exceeds the 30 MB download limit.")
                content = response.read(MAX_WORKBOOK_BYTES + 1)
                if len(content) > MAX_WORKBOOK_BYTES:
                    raise WorkbookSyncError("The source workbook exceeds the 30 MB download limit.")
                if is_zipfile(BytesIO(content)):
                    return content
                failures.append("non-Excel response")
        except WorkbookSyncError:
            raise
        except HTTPError as exc:
            failures.append(f"HTTP {exc.code}")
        except (OSError, URLError, ValueError):
            failures.append("connection failure")

    if "sign-in redirect" in failures:
        raise WorkbookSyncError(
            "The workbook link redirected the dashboard to Microsoft sign-in. "
            "Set WORKBOOK_DOWNLOAD_URL to the exact Anyone link that downloads the file in a private window."
        )
    http_errors = sorted({failure for failure in failures if failure.startswith("HTTP ")})
    if http_errors:
        raise WorkbookSyncError(
            f"The workbook server rejected the dashboard request ({', '.join(http_errors)}). "
            "Verify that WORKBOOK_DOWNLOAD_URL is the Anyone link and that its access has not expired."
        )
    if "non-Excel response" in failures:
        raise WorkbookSyncError(
            "The workbook link returned a web page instead of an Excel file. "
            "Verify that WORKBOOK_DOWNLOAD_URL is a direct download link."
        )
    raise WorkbookSyncError("Could not connect to the workbook download link.")


class _GraphRedirectHandler(HTTPRedirectHandler):
    def redirect_request(self, request, response, code, message, headers, new_url):
        redirected = super().redirect_request(
            request, response, code, message, headers, new_url
        )
        if redirected is None:
            return None
        old_host = (urlparse(request.full_url).hostname or "").lower()
        new_host = (urlparse(new_url).hostname or "").lower()
        if old_host != new_host:
            for name in tuple(redirected.headers):
                if name.lower() == "authorization":
                    del redirected.headers[name]
            for name in tuple(redirected.unredirected_hdrs):
                if name.lower() == "authorization":
                    del redirected.unredirected_hdrs[name]
        return redirected


def fetch_graph_workbook(tenant_id, client_id, client_secret, drive_id, item_id):
    try:
        import msal
    except ImportError as exc:
        raise WorkbookSyncError("Microsoft Graph sync requires the msal package.") from exc

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    application = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )
    token_result = application.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    access_token = token_result.get("access_token")
    if not access_token:
        raise WorkbookSyncError(
            "Microsoft Graph authentication failed; verify the app credentials and admin consent."
        )

    graph_url = (
        "https://graph.microsoft.com/v1.0/drives/"
        f"{quote(drive_id, safe='')}/items/{quote(item_id, safe='')}/content"
    )
    request = Request(
        graph_url,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/octet-stream",
            "User-Agent": "OffloadDashboard/1.0",
        },
    )
    try:
        opener = build_opener(_GraphRedirectHandler)
        with opener.open(request, timeout=60) as response:
            content_length = response.headers.get("Content-Length")
            if content_length and int(content_length) > MAX_WORKBOOK_BYTES:
                raise WorkbookSyncError("The source workbook exceeds the 30 MB download limit.")
            content = response.read(MAX_WORKBOOK_BYTES + 1)
    except WorkbookSyncError:
        raise
    except (OSError, URLError, ValueError) as exc:
        raise WorkbookSyncError(
            "Microsoft Graph could not download the configured file; verify the file IDs and its read permission."
        ) from exc
    if len(content) > MAX_WORKBOOK_BYTES:
        raise WorkbookSyncError("The source workbook exceeds the 30 MB download limit.")
    return content


def refresh_from_bytes(workbook_bytes, data_path):
    data = parse_workbook(workbook_bytes)
    version, content = workbook_data_script(data)
    destination = os.path.abspath(data_path)
    directory = os.path.dirname(destination)
    os.makedirs(directory, exist_ok=True)
    file_descriptor, temporary_path = tempfile.mkstemp(prefix=".workbook-data-", suffix=".tmp", dir=directory)
    try:
        with os.fdopen(file_descriptor, "w", encoding="utf-8", newline="\n") as output:
            output.write(content)
        os.replace(temporary_path, destination)
    except Exception:
        try:
            os.unlink(temporary_path)
        except OSError:
            pass
        raise
    return {"version": version, "rows": len(data["records"])}


def refresh_from_url(url, data_path):
    return refresh_from_bytes(fetch_workbook(url), data_path)


class WorkbookAutoSync:
    def __init__(
        self,
        url,
        data_path,
        interval_seconds,
        fetcher=None,
        source=None,
        configuration_error=None,
    ):
        if interval_seconds < 60:
            raise ValueError("WORKBOOK_REFRESH_SECONDS must be at least 60.")
        self.url = url.strip()
        self.fetcher = fetcher
        self.source = source or ("anonymous_link" if self.url else None)
        self.configuration_error = configuration_error
        self.data_path = data_path
        self.interval_seconds = interval_seconds
        self._lock = threading.Lock()
        self._status_lock = threading.Lock()
        self._stop = threading.Event()
        self._thread = None
        self._last_success_at = None
        self._last_error = None
        self._last_version = None
        self._last_rows = None

    @property
    def configured(self):
        return bool(self.url or self.fetcher)

    def status(self):
        with self._status_lock:
            return {
                "configured": self.configured,
                "source": self.source,
                "configurationError": self.configuration_error,
                "refreshSeconds": self.interval_seconds,
                "lastSuccessAt": self._last_success_at,
                "lastError": self._last_error,
                "version": self._last_version,
                "rows": self._last_rows,
            }

    def record_external_refresh(self, result):
        with self._status_lock:
            self._last_success_at = datetime.now().astimezone().isoformat(timespec="seconds")
            self._last_error = None
            self._last_version = result["version"]
            self._last_rows = result["rows"]

    def refresh_once(self):
        if not self.configured:
            raise WorkbookSyncError(
                self.configuration_error or "Workbook auto-sync is not configured."
            )
        if not self._lock.acquire(blocking=False):
            raise WorkbookSyncError("A workbook refresh is already in progress.")
        try:
            workbook_bytes = self.fetcher() if self.fetcher else fetch_workbook(self.url)
            result = refresh_from_bytes(workbook_bytes, self.data_path)
            with self._status_lock:
                self._last_success_at = datetime.now().astimezone().isoformat(timespec="seconds")
                self._last_error = None
                self._last_version = result["version"]
                self._last_rows = result["rows"]
            return result
        except WorkbookSyncError as exc:
            with self._status_lock:
                self._last_error = str(exc)
            raise
        except Exception as exc:
            with self._status_lock:
                self._last_error = "Workbook refresh failed unexpectedly; the last good dashboard data was kept."
            raise WorkbookSyncError(self._last_error) from exc
        finally:
            self._lock.release()

    def _run(self):
        while not self._stop.is_set():
            try:
                result = self.refresh_once()
                print(f"Workbook auto-sync refreshed {result['rows']} rows; version {result['version'][:12]}.")
            except WorkbookSyncError as exc:
                print(f"Workbook auto-sync failed; keeping the last good data. {exc}")
            self._stop.wait(self.interval_seconds)

    def start(self):
        if not self.configured:
            if self.configuration_error:
                print(f"Workbook auto-sync is disabled. {self.configuration_error}")
            elif self.source == "local_upload":
                print("Workbook auto-sync is waiting for authenticated local uploads.")
            else:
                print("Workbook auto-sync is disabled; configure Microsoft Graph or WORKBOOK_DOWNLOAD_URL.")
            return
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._run, name="workbook-auto-sync", daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=5)
