import csv
from io import StringIO
import re
from datetime import date, datetime, time


TRACKER_COLUMNS = (
    "Departure_Date",
    "Org",
    "Dest",
    "Lane",
    "Mawb(with prefix)",
    "Vendor_name",
    "NFO_Flag",
    "Direct / Console",
    "Flight_Number",
    "ETD",
    "ETA",
    "Type as per ETD",
    "Bags",
    "CD weight",
    "Total Cons",
    "Offloaded bags",
    "Impacted cons",
    "Next_flight",
    "ETD",
    "ETA",
    "NDC_Remarks",
    "LH_Pickup_date",
    "LH_Pickup_time",
    "Status",
)


NDC_REMARK_OPTIONS = (
    "PUD Vehicle / GSA arrived in the Hub / Location after Airport Cut Off",
    "Missed connections due to lodgement failure",
    "Offloaded due to security reasons / prohibited contain by airlines",
    "Missed connection due to non-availability of vendor",
    "Holded At Hub due to no connection in same day",
    "Connected in later flight due to flight cancellation",
    "Missed Connection due to improper packaging",
    "Missed Connection due to Odd size / over size shipment",
    "Delay in connection due to Terminal issue ( Manpower / Infra / Backlog/Embargo / strike)",
    "Missed connections due to improper documentation",
    "Offloaded by airlines due to passenger baggage / Fuel restriction",
    "Offloaded by Airlines due to priority shipments - Like Perishables , HUM , UPSC mail etc",
    "Offloaded by Airlines due to space unavailability - Completely",
    "Offloaded by Airlines due to space unavailability - Partially",
    "Flight delayed due to weather condition / runway closure",
    "Flight delayed due to flight diversion",
    "Delay in Retrieval process - Airline ( manpower / infra / terminal etc)",
    "Delay in Retrieval process - Vehicle issue",
    "Delay in Retrieval Process - Vendor delayed",
    "Delay in Retrieval due sector booking (sent manually) not done by previous node",
    "Delay in Retrieval due to wrong marking by origin airport.",
    "Reached but not retrieved at destination on time by own staff",
    "Delayed due to misrouting by Vendor / Airline",
    "Delayed due to misrouting by own staff",
    "Shipment holded at hub due to wrong marking at Hub",
    "Dedicated vehicle reached late Breakdown / traffic / driver issue etc",
    "Holded At Hub due to partially receiving",
    "Holded At Hub due to received damaged",
    "Vendor breached the TAT",
    "Air moved in Surface",
    "Surface moved in Air",
    "Rerouted through alternate airport due to origin airport constraints",
    "Ewaybill expired in 3rd party network, hence cargo not accepted from the Vendor",
    "eWaybill Non-compliance - expired / not-reassigned",
    "eWaybill not re-assigned to CriticaLog, hence cargo not accepted from the Vendor",
    "Ewaybill part B not re-assigned to CriticaLog, hence OFD delayed",
    "Vehicle caught by sales tax team",
)


TESTING_CLIENT_CODE = "206324"

# Mirrors Multijoins::lccscostQuery(), Multijoins::skipTestingAWB() and the
# report's Docket (xdockethdr) NFO lookup in lccsCostNew2db.php, but for all
# hubs (no per-user hub filter) and air MAWBs only.
LCCS_TRACKER_QUERY = """
SELECT
    awb.XAWBID AS XAWBID,
    awb.XAWSDPRDT AS XAWSDPRDT,
    awb.XAWBOOKLOC AS XAWBOOKLOC,
    awb.XAWPKUPLOC AS XAWPKUPLOC,
    awb.XAWBILLNO AS mawb_no,
    awb.VENDORNM AS VENDORNM,
    awb.SECTORMODE AS SECTORMODE,
    awb.XAWBMODE AS XAWBMODE,
    awb.XAWFLTNO AS XAWFLTNO,
    awb.XAWFLTNO2 AS XAWFLTNO2,
    awb.XAWSDPRTIME AS XAWSDPRTIME,
    awb.XAWSARVTIME AS XAWSARVTIME,
    awb.XAWPACKEGE AS XAWPACKEGE,
    awb.XAWBCDWGT AS XAWBCDWGT,
    (
        SELECT COUNT(DISTINCT cons.XAWDDKTID) FROM xawbdtl cons
        WHERE cons.XAIRWAYBILLNO = awb.XAWBILLNO
    ) AS total_cons,
    (
        SELECT region_hub.HUBZONE FROM mhubmst region_hub
        WHERE region_hub.MHBHUBCD = awb.XAWBOOKLOC
        LIMIT 1
    ) AS org_region,
    CASE WHEN EXISTS (
        SELECT 1
        FROM xawbdtl nfo_detail
        JOIN xdockethdr nfo_docket
            ON nfo_docket.XDKTID = nfo_detail.XAWDDKTID
        WHERE nfo_detail.XAIRWAYBILLNO = awb.XAWBILLNO
          AND UPPER(nfo_docket.XDKTPROMODE) = 'NFO'
          AND (nfo_docket.XLASTNEGATIVECD != 'AA' OR nfo_docket.XLASTNEGATIVECD IS NULL)
    ) THEN 'NFO' ELSE '' END AS is_nfo
FROM xawbhdr awb
WHERE awb.XAWBMODE = 'A'
  AND DATE(awb.XAWSDPRDT) BETWEEN :from_date AND :to_date
  AND (:hub_code IS NULL OR awb.XAWBOOKLOC = :hub_code)
  AND EXISTS (
      SELECT 1 FROM mhubmst hub
      WHERE hub.MHBHUBCD = awb.XAWBOOKLOC
  )
  AND EXISTS (
      SELECT 1 FROM xawbdtl detail
      WHERE detail.XAIRWAYBILLNO = awb.XAWBILLNO
  )
  AND NOT EXISTS (
      SELECT 1
      FROM xawbdtl test_detail
      JOIN xdockethdr test_docket
          ON test_docket.XDKTID = test_detail.XAWDDKTID
      WHERE test_detail.XAIRWAYBILLNO = awb.XAWBILLNO
        AND test_docket.XCLNTCD = :testing_client_code
  )
ORDER BY awb.XAWBID DESC
""".strip()


class TrackerDataError(ValueError):
    pass


def _value(row, *names):
    for name in names:
        if name in row and row[name] is not None:
            return row[name]
    return ""


def _override(manual, name, source_value):
    if name in manual and manual[name] is not None:
        return _text(manual[name])
    return source_value


def _text(value):
    if value is None:
        return ""
    if isinstance(value, (date, datetime, time)):
        return value.isoformat()
    return str(value).strip()


def _date_text(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    text = _text(value)
    match = re.match(r"(\d{4}-\d{2}-\d{2})[T ]\d{2}:\d{2}", text)
    return match.group(1) if match else text


def _sector(value):
    return _text(value)[:3]


def _dispatch_type(value):
    if value is None or not _text(value):
        return ""
    if isinstance(value, datetime):
        value = value.time()
    if isinstance(value, time):
        parsed = value
    else:
        match = re.fullmatch(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", _text(value))
        if not match:
            return ""
        hour, minute, second = match.groups()
        try:
            parsed = time(int(hour), int(minute), int(second or 0))
        except ValueError:
            return ""
    if parsed <= time(9, 0):
        return "Prime"
    if parsed <= time(12, 0):
        return "Lean 1"
    return "Lean 2"


def map_lccs_row(source, manual=None):
    shipment_id = _text(_value(source, "XAWBID"))
    if not shipment_id:
        raise TrackerDataError("LCCS row is missing XAWBID; no stable record key is available.")

    manual = manual or {}
    org = _sector(_value(source, "XAWBOOKLOC"))
    dest = _sector(_value(source, "XAWPKUPLOC"))
    raw_flight = _text(_value(source, "XAWFLTNO"))
    flight = (
        _text(_value(source, "XAWFLTNO2"))
        if raw_flight.upper() == "NA"
        else raw_flight
    )
    flight = _override(manual, "flight_number_override", flight)
    sector_mode = _text(_value(source, "SECTORMODE"))
    if sector_mode:
        sector_mode = "Direct" if sector_mode.upper() == "D" else "Console"

    # The PHP report assigns Prime/Lean slots only to air MAWBs (XAWBMODE = "A").
    is_air = _text(_value(source, "XAWBMODE")).upper() == "A"
    etd = _text(_value(source, "XAWSDPRTIME"))
    eta = _text(_value(source, "XAWSARVTIME"))  # scheduled arrival time
    if eta in ("00:00:00", "00:00"):
        eta = ""
    total_cons = _value(source, "total_cons")
    if total_cons in ("", None, 0):
        total_cons = _value(manual, "total_cons")
    scheduled_etd = _override(manual, "scheduled_etd_override", etd)
    scheduled_eta = _override(manual, "scheduled_eta_override", eta)
    impacted_cons = _value(manual, "impacted_cons")
    if isinstance(impacted_cons, str):
        impacted_cons = impacted_cons.rstrip()
    ndc_remarks = _value(manual, "ndc_remarks")
    if ndc_remarks and ndc_remarks not in NDC_REMARK_OPTIONS:
        raise TrackerDataError("NDC_Remarks must match one of the supplied dropdown options.")

    values = [
        _date_text(_value(source, "XAWSDPRDT")),
        org,
        dest,
        f"{org}-{dest}" if org and dest else "",
        _text(_value(source, "mawb_no", "XAWBILLNO")),
        _text(_value(source, "VENDORNM")),
        _text(_value(source, "is_nfo")),
        sector_mode,
        flight,
        scheduled_etd,
        scheduled_eta,
        _dispatch_type(_value(source, "XAWSDPRTIME")) if is_air else "",
        _value(source, "XAWPACKEGE"),
        _value(source, "XAWBCDWGT"),
        total_cons,
        _value(manual, "offloaded_bags"),
        impacted_cons,
        _value(manual, "next_flight"),
        _value(manual, "revised_etd"),
        _value(manual, "revised_eta"),
        ndc_remarks,
        _value(manual, "lh_pickup_date"),
        _value(manual, "lh_pickup_time"),
        _value(manual, "status"),
    ]
    if len(values) != len(TRACKER_COLUMNS):
        raise TrackerDataError("Mapped tracker row does not match the 24-column report.")
    region = _text(_value(source, "org_region", "HUBZONE")).title()
    return {"shipment_id": shipment_id, "region": region, "values": values}


def map_lccs_rows(source_rows, manual_by_shipment=None):
    manual_by_shipment = manual_by_shipment or {}
    mapped = {}
    for source in source_rows:
        shipment_id = _text(_value(source, "XAWBID"))
        record = map_lccs_row(source, manual_by_shipment.get(shipment_id))
        previous = mapped.get(shipment_id)
        if previous is not None and previous != record:
            raise TrackerDataError(
                f"LCCS returned conflicting rows for XAWBID {shipment_id}."
            )
        mapped[shipment_id] = record
    return list(mapped.values())


def tracker_csv(records):
    output = StringIO(newline="")
    writer = csv.writer(output)
    writer.writerow(TRACKER_COLUMNS)
    writer.writerows(record["values"] for record in records)
    return output.getvalue()
