let records = window.offloadWorkbookRecords;
const $ = id => document.getElementById(id);
let loadedWorkbookVersion = window.offloadWorkbookUpdatedAt || "";
let lastRefreshAt = null;
let activeRangeMode = "all";
let watchlistDate = toLocalDateString(new Date());

function updateLastRefreshStamp() {
  const stamp = new Date();
  lastRefreshAt = stamp;
  const refreshText = stamp.toLocaleString(undefined, {
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  });
  const target = $("lastRefreshBadge");
  if (target) {
    target.textContent = `Last refreshed: ${refreshText}`;
  }
}

if (!Array.isArray(records) || records.length === 0) {
  const error = document.createElement("p");
  error.className = "load-error";
  error.textContent = "Could not load workbook-data.js. Keep it in the same folder as dashboard.html and reload.";
  document.querySelector(".filters").after(error);
  throw new Error("Workbook data is unavailable.");
}

const escapeHtml = value => String(value ?? "").replace(/[&<>"']/g, char => ({
  "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
}[char]));
const numberText = value => Number.isFinite(value)
  ? value.toLocaleString(undefined, { maximumFractionDigits: 2 })
  : "—";
const formatPct = value => `${value.toFixed(1)}%`;
const sum = (rows, key) => rows.reduce((total, row) => total + (Number.isFinite(row[key]) ? row[key] : 0), 0);
const count = (rows, predicate) => rows.filter(predicate).length;
const watchlistPeriodCount = 4;
const watchlistWindowDays = 28;
const watchlistMinimumBags = 20;
const isShipment = row => row.recordType === "Shipment";
const hasBagValues = row => Number.isFinite(row.bags) && Number.isFinite(row.offloadedBags);
const bagsKnown = row => isShipment(row) && hasBagValues(row) &&
  Number.isInteger(row.bags) && Number.isInteger(row.offloadedBags) &&
  row.bags >= 0 && row.offloadedBags >= 0 && row.offloadedBags <= row.bags;
const invalidBagPair = row => isShipment(row) && hasBagValues(row) && !bagsKnown(row);
const offloaded = row => bagsKnown(row) && row.offloadedBags > 0;
const hasRemark = row => Boolean((row.remarks && row.remarks.trim()) || (row.ndcRemarks && row.ndcRemarks.trim()));
const rateFor = rows => {
  const known = rows.filter(bagsKnown);
  const totalBags = sum(known, "bags");
  return totalBags ? sum(known, "offloadedBags") / totalBags * 100 : 0;
};
const unique = key => [...new Set(records.map(row => row[key]).filter(value => typeof value === "string" && value.trim()))]
  .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
const groupBy = (rows, key) => {
  const groups = new Map();
  rows.forEach(row => {
    const raw = row[key];
    const name = typeof raw === "string" ? raw.trim() : (raw == null ? "" : String(raw).trim());
    if (!name || name === "Unspecified") {
      return;
    }
    if (!groups.has(name)) groups.set(name, []);
    groups.get(name).push(row);
  });
  return [...groups.entries()].map(([name, items]) => ({ name, items }));
};
let maxDate;
let minDate;
const standardDetailColumns = [
  ["Type", "recordType"], ["Source sheet", "sourceSheet"], ["Source row", "sourceRow"],
  ["Date", "date"], ["Region", "region"], ["Origin", "origin"], ["Destination", "destination"],
  ["Lane", "lane"], ["MAWB", "mawb"], ["Vendor Name", "vendor"], ["Flight Number", "flight"],
  ["Bags", "bags"], ["CD Weight", "cdWeight"], ["Total Cons", "totalCons"],
  ["Offloaded Bags", "offloadedBags"], ["Impacted Cons", "impactedCons"],
  ["Next Flight", "nextFlight"], ["ETD", "etd"], ["ETA", "eta"],
  ["Next ETD", "nextEtd"], ["Next ETA", "nextEta"], ["Remarks", "remarks"], ["NDC Remarks", "ndcRemarks"]
];
let detailColumns = standardDetailColumns.concat(
  (window.offloadWorkbookColumns || []).map(column => [column.label, column.key])
);
const tablePageSize = 100;
const dialogColumns = detailColumns.slice(0, 23);
let dialogPage = 0;

function setOptions(id, label, values, preserveSelection = false) {
  const selected = preserveSelection ? $(id).value : "";
  $(id).innerHTML = `<option value="">All ${label}</option>` +
    values.map(value => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("");
  if (preserveSelection && [...$(id).options].some(option => option.value === selected)) {
    $(id).value = selected;
  }
}

function populateFilters(preserveSelection = false) {
  setOptions("regionFilter", "regions", unique("region"), preserveSelection);
  setOptions("vendorFilter", "vendors", unique("vendor"), preserveSelection);
  setOptions("laneFilter", "lanes", unique("lane"), preserveSelection);
  setOptions("originFilter", "origins", unique("origin"), preserveSelection);
  setOptions("flightFilter", "flights", unique("flight"), preserveSelection);
}

function updateDateBounds() {
  const dates = records.map(dateForRow).filter(Boolean);
  maxDate = dates.reduce((latest, value) => value > latest ? value : latest, "");
  minDate = dates.reduce((earliest, value) => value < earliest ? value : earliest, maxDate);
  ["startDate", "endDate"].forEach(id => {
    $(id).min = minDate;
    $(id).max = maxDate;
  });
}

function resetToAllDates() {
  activeRangeMode = "all";
  $("startDate").value = minDate || "";
  $("endDate").value = maxDate || "";
  document.querySelectorAll(".quick-btn[data-range]").forEach(button =>
    button.classList.toggle("active", button.dataset.range === "all"));
  render();
}

populateFilters();
updateDateBounds();
resetToAllDates();

function rowMatchesDateRange(row, start, end) {
  const date = dateForRow(row);
  if (!date) {
    return !start && !end;
  }
  if (!start && !end) {
    return true;
  }
  return (!start || date >= start) && (!end || date <= end);
}

function isIsoDate(value) {
  if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return false;
  }
  const [year, month, day] = value.split("-").map(Number);
  if (year < 2000) {
    return false;
  }
  const parsed = new Date(Date.UTC(year, month - 1, day));
  return parsed.getUTCFullYear() === year &&
    parsed.getUTCMonth() === month - 1 &&
    parsed.getUTCDate() === day;
}

function dateForRow(row) {
  if (isIsoDate(row.date)) {
    return row.date;
  }
  if (!Array.isArray(row.dateVariants)) {
    return "";
  }
  return row.dateVariants.find(isIsoDate) || "";
}

function currentRows() {
  const start = $("startDate").value;
  const end = $("endDate").value;
  const includeUndated = start === minDate && end === maxDate;
  const filters = {
    region: $("regionFilter").value,
    vendor: $("vendorFilter").value,
    lane: $("laneFilter").value,
    origin: $("originFilter").value,
    flight: $("flightFilter").value
  };
  return records.filter(row =>
    rowMatchesDateRange(row, start, end) &&
    Object.entries(filters).every(([key, value]) => !value || row[key] === value)
  );
}

function drawMetrics(rows) {
  const shipmentRows = rows.filter(isShipment);
  const known = shipmentRows.filter(bagsKnown);
  const offloadedBags = sum(shipmentRows, "offloadedBags");
  const invalidCounts = count(shipmentRows, invalidBagPair);
  const missingBagData = count(shipmentRows, row => !hasBagValues(row));
  const metrics = [
    ["Bag offload %", "Valid offloaded bags / valid total bags", formatPct(rateFor(shipmentRows)), `${known.length.toLocaleString()} valid rows · ${invalidCounts.toLocaleString()} invalid rows excluded from rate`],
    ["Offloaded bags", "Sum of recorded offloaded-bag values", numberText(offloadedBags), `${missingBagData.toLocaleString()} shipments have no recorded offloaded-bag value`],
    ["CD weight", "Total recorded CD weight", numberText(sum(shipmentRows, "cdWeight")), "All filtered shipment rows"],
    ["Impacted consignments", "Sum of recorded impacted-consignment values", numberText(sum(shipmentRows, "impactedCons")), "All filtered shipment rows"]
  ];
  $("metrics").innerHTML = metrics.map(([title, caption, value, note]) =>
    `<article class="metric-card"><div class="metric-heading">${escapeHtml(title)}</div><div class="metric-caption">${escapeHtml(caption)}</div><div class="metric-value">${escapeHtml(value)}</div><div class="metric-delta">${escapeHtml(note)}</div></article>`
  ).join("");
}

function makeBarInteractive(element, name, rows) {
  element.classList.add("selectable");
  element.tabIndex = 0;
  element.setAttribute("role", "button");
  element.setAttribute("aria-label", `View and download shipment details for ${name}`);
  const open = () => showBarDetails(name, rows);
  element.addEventListener("click", open);
  element.addEventListener("keydown", event => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      open();
    }
  });
}

function drawRegion(rows) {
  const shipmentRows = rows.filter(isShipment);
  const groups = groupBy(shipmentRows, "region").map(group => ({
    ...group,
    totalBags: sum(group.items.filter(row => Number.isInteger(row.bags) && row.bags >= 0), "bags"),
    eligibleBags: sum(group.items.filter(bagsKnown), "bags"),
    offloadedBags: sum(group.items.filter(bagsKnown), "offloadedBags"),
    impactedCons: sum(group.items, "impactedCons"),
    rate: group.items.some(bagsKnown) ? rateFor(group.items) : null
  })).sort((a, b) => a.name.localeCompare(b.name));
  const network = rateFor(shipmentRows);
  const maxBags = Math.max(1, ...groups.map(group => group.rate === null ? group.totalBags : group.eligibleBags));
  $("regionChart").innerHTML = groups.length
    ? `<div class="table-head region-layout"><span>Origin region</span><span></span><span>Rate</span><span class="count">Offloaded / bags</span><span class="count">Impacted cons.</span></div>` +
      groups.map(group => `<div class="bar-row region-layout" title="${escapeHtml(group.name)}: ${group.rate === null ? `${numberText(group.totalBags)} total bags; no valid offload data` : formatPct(group.rate)}">
        <span class="bar-label">${escapeHtml(group.name)}</span><div class="track-cell"><div class="bar-track threshold" style="--threshold:${Math.min(network,100)}%"><div class="bar-fill${group.rate === null ? " bar-no-data" : ""}" style="width:${(group.rate === null ? group.totalBags : group.eligibleBags) / maxBags * 100}%"><span class="bar-before" style="width:${group.rate === null ? 0 : Math.max(0,100-group.rate)}%"></span><span class="bar-after" style="width:${group.rate === null ? 0 : Math.min(group.rate,100)}%"></span></div></div></div>
        <span class="value">${group.rate === null ? "—" : formatPct(group.rate)}</span><span class="count">${group.rate === null ? (group.totalBags ? `0 / ${numberText(group.totalBags)}` : "No valid bag data") : `${numberText(group.offloadedBags)} / ${numberText(group.eligibleBags)}`}</span><span class="count">${numberText(group.impactedCons)}</span></div>`).join("")
    : `<div class="empty-chart">No regional data.</div>`;
  $("regionChart").querySelectorAll(".bar-row").forEach((element, index) =>
    makeBarInteractive(element, groups[index].name, groups[index].items));
  $("regionFooter").textContent = `Network average ${formatPct(network)}`;
}

function drawRemarks(rows) {
  const offloads = rows.filter(offloaded);
  const groups = groupBy(offloads, "region").map(group => {
    const withRemark = sum(group.items.filter(hasRemark), "offloadedBags");
    const withoutRemark = sum(group.items.filter(row => !hasRemark(row)), "offloadedBags");
    const totalOffloadedBags = sum(group.items, "offloadedBags");
    const impactedCons = sum(group.items, "impactedCons");
    return { ...group, withRemark, withoutRemark, totalOffloadedBags, impactedCons, rate: totalOffloadedBags ? withRemark / totalOffloadedBags * 100 : 0 };
  }).sort((a, b) => a.name.localeCompare(b.name));

  $("reasonChart").innerHTML = groups.length
    ? `<div class="table-head reason-layout"><span>Origin region</span><span></span><span>With remark</span><span class="count">Offloaded bags</span><span class="count">Impacted cons.</span></div>` +
      groups.map(group => `<div class="bar-row reason-layout" title="${escapeHtml(group.name)}: ${numberText(group.withRemark)} bags with a remark, ${numberText(group.withoutRemark)} bags without a remark">
        <span class="bar-label">${escapeHtml(group.name)}</span><div class="track-cell"><div class="bar-track"><div class="bar-fill" style="width:100%"><span class="bar-reason" style="width:${group.rate}%"></span><span class="bar-missing" style="width:${100 - group.rate}%"></span></div></div></div>
        <span class="value">${formatPct(group.rate)} (${numberText(group.withRemark)})</span><span class="count">${numberText(group.totalOffloadedBags)}</span><span class="count">${numberText(group.impactedCons)}</span></div>`).join("")
    : `<div class="empty-chart">No recorded offloads in this selection.</div>`;
  $("reasonChart").querySelectorAll(".bar-row").forEach((element, index) =>
    makeBarInteractive(element, `${groups[index].name} — remarks`, groups[index].items));

  const totalWithRemark = sum(offloads.filter(hasRemark), "offloadedBags");
  const totalWithoutRemark = sum(offloads.filter(row => !hasRemark(row)), "offloadedBags");
  const totalOffloadedBags = sum(offloads, "offloadedBags");
  $("reasonSummary").textContent = `${numberText(totalWithRemark)} of ${numberText(totalOffloadedBags)} offloaded bags have a remark; ${numberText(totalWithoutRemark)} have no remark.`;
}

function drawRanked(rows, key, target, label, layout) {
  const shipmentRows = rows.filter(isShipment);
  const groups = groupBy(shipmentRows, key).map(group => ({
    ...group,
    eligibleBags: sum(group.items.filter(bagsKnown), "bags"),
    offloadedBags: sum(group.items.filter(bagsKnown), "offloadedBags"),
    impactedCons: sum(group.items, "impactedCons"),
    rate: rateFor(group.items)
  })).filter(group => group.eligibleBags > 0)
    .sort((a, b) => b.rate - a.rate || b.eligibleBags - a.eligibleBags || a.name.localeCompare(b.name))
    .slice(0, 10);
  const maxBags = Math.max(1, ...groups.map(group => group.eligibleBags));
  $(target).innerHTML = groups.length
    ? `<div class="table-head ${layout}"><span>${label}</span><span></span><span>Rate</span><span class="count">Offloaded / bags</span><span class="count">Impacted cons.</span></div>` +
      groups.map(group => `<div class="bar-row ${layout}" title="${escapeHtml(group.name)}: ${formatPct(group.rate)}">
        <span class="bar-label">${escapeHtml(group.name)}</span><div class="track-cell"><div class="bar-track"><div class="bar-fill" style="width:${group.eligibleBags / maxBags * 100}%"><span class="bar-before" style="width:${Math.max(0,100-group.rate)}%"></span><span class="bar-after" style="width:${Math.min(group.rate,100)}%"></span></div></div></div>
        <span class="value">${formatPct(group.rate)}</span><span class="count">${numberText(group.offloadedBags)} / ${numberText(group.eligibleBags)}</span><span class="count">${numberText(group.impactedCons)}</span></div>`).join("")
    : `<div class="empty-chart">No data with recorded bag counts.</div>`;
  $(target).querySelectorAll(".bar-row").forEach((element, index) =>
    makeBarInteractive(element, groups[index].name, groups[index].items));
}

function drawConsistentUnderperformers(key, target, label, start, end) {
  const startTime = Date.parse(`${start}T00:00:00Z`);
  const network = Array.from({ length: watchlistPeriodCount }, () => ({
    eligibleBags: 0,
    offloadedBags: 0
  }));
  const groups = new Map();

  records.forEach(row => {
    if (!bagsKnown(row)) {
      return;
    }

    const date = dateForRow(row);
    if (!date || date < start || date > end) {
      return;
    }

    const dayIndex = Math.floor((Date.parse(`${date}T00:00:00Z`) - startTime) / 86400000);
    const period = Math.min(watchlistPeriodCount - 1, Math.floor(dayIndex * watchlistPeriodCount / watchlistWindowDays));
    network[period].eligibleBags += row.bags;
    network[period].offloadedBags += row.offloadedBags;

    const rawName = row[key];
    const name = typeof rawName === "string" ? rawName.trim() : "";
    if (!name || name === "Unspecified") {
      return;
    }

    if (!groups.has(name)) {
      groups.set(name, {
        name,
        items: [],
        eligibleBags: 0,
        offloadedBags: 0,
        periods: Array.from({ length: watchlistPeriodCount }, () => ({
          eligibleBags: 0,
          offloadedBags: 0
        }))
      });
    }

    const group = groups.get(name);
    group.items.push(row);
    group.eligibleBags += row.bags;
    group.offloadedBags += row.offloadedBags;
    group.periods[period].eligibleBags += row.bags;
    group.periods[period].offloadedBags += row.offloadedBags;
  });

  const networkRates = network.map(week =>
    week.eligibleBags ? week.offloadedBags / week.eligibleBags * 100 : null
  );
  const qualifying = [...groups.values()].map(group => {
    const poorPeriods = group.periods.filter((period, index) =>
      period.eligibleBags >= watchlistMinimumBags &&
      networkRates[index] !== null &&
      // A higher offload % means more bags were offloaded: that is the poor performance.
      period.offloadedBags / period.eligibleBags * 100 > networkRates[index]
    ).length;
    return {
      ...group,
      poorPeriods,
      rate: group.eligibleBags ? group.offloadedBags / group.eligibleBags * 100 : 0
    };
  }).filter(group => group.poorPeriods >= 3)
    .sort((a, b) => b.rate - a.rate || b.poorPeriods - a.poorPeriods || a.name.localeCompare(b.name));

  const element = $(target);
  element.innerHTML = qualifying.length
    ? `<div class="table-head watchlist-layout"><span>${label}</span><span class="count" title="Offload percentage over the last 4 weeks">Offload %</span><span class="count" title="Seven-day periods above the network average">Weeks above avg.</span><span class="count" title="Valid bags in the four-week period">Bags</span></div>` +
      qualifying.map(group => `<div class="bar-row watchlist-layout" title="${escapeHtml(group.name)}: ${formatPct(group.rate)} over the last 4 weeks; above the network average in ${group.poorPeriods} of 4 weeks">
        <span class="bar-label">${escapeHtml(group.name)}</span><span class="count">${formatPct(group.rate)}</span><span class="count">${group.poorPeriods}/4</span><span class="count">${numberText(group.eligibleBags)}</span></div>`).join("")
    : `<div class="empty-chart">No ${label.toLowerCase()} met the consistency and volume criteria.</div>`;
  element.querySelectorAll(".bar-row").forEach((row, index) =>
    makeBarInteractive(row, `${label}: ${qualifying[index].name}`, qualifying[index].items));
}

function drawConsistentPerformance() {
  const today = toLocalDateString(new Date());
  const startDate = new Date(`${today}T00:00:00`);
  startDate.setDate(startDate.getDate() - watchlistWindowDays + 1);
  const start = toLocalDateString(startDate);
  drawConsistentUnderperformers("vendor", "consistentVendorChart", "Vendor", start, today);
  drawConsistentUnderperformers("lane", "consistentLaneChart", "Lane", start, today);
}

function displayField(row, key) {
  if (key.startsWith("raw_")) return row.sourceValues?.[key] ?? "";
  if (key === "date") return dateForRow(row);
  if (key === "recordType" && !row.mawb && Number.isFinite(row.bags) && Number.isFinite(row.offloadedBags)) {
    return "Shipment data (no MAWB)";
  }
  if (["etd", "eta", "nextEtd", "nextEta"].includes(key) && Number.isFinite(row[key])) {
    const minutes = Math.round(row[key] * 1440) % 1440;
    return `${String(Math.floor(minutes / 60)).padStart(2, "0")}:${String(minutes % 60).padStart(2, "0")}`;
  }
  return row[key] ?? "";
}

let activeBarRows = [];
let activeBarName = "";

function showBarDetails(name, rows) {
  activeBarName = name;
  activeBarRows = rows.slice().sort((a, b) =>
    dateForRow(a).localeCompare(dateForRow(b)) || a.sourceSheet.localeCompare(b.sourceSheet) || a.sourceRow - b.sourceRow
  );
  $("barDialogTitle").textContent = name;
  $("barDialogSummary").textContent =
    `${activeBarRows.length.toLocaleString()} matching workbook rows · ${numberText(sum(activeBarRows.filter(bagsKnown), "offloadedBags"))} offloaded bags from valid rows · ${count(activeBarRows, invalidBagPair).toLocaleString()} invalid bag-count rows · ${numberText(sum(activeBarRows, "cdWeight"))} CD weight`;
  $("downloadStatus").textContent = "";
  dialogPage = 0;
  renderDialogPage();
  $("barDialog").showModal();
}

function renderDialogPage() {
  const pageCount = Math.max(1, Math.ceil(activeBarRows.length / tablePageSize));
  dialogPage = Math.min(dialogPage, pageCount - 1);
  const start = dialogPage * tablePageSize;
  $("barDialogRows").innerHTML = activeBarRows.slice(start, start + tablePageSize).map(row =>
    `<tr>${dialogColumns.map(([, key]) => `<td>${escapeHtml(displayField(row, key))}</td>`).join("")}</tr>`
  ).join("");
  $("dialogPageStatus").textContent =
    `${activeBarRows.length ? start + 1 : 0}–${Math.min(start + tablePageSize, activeBarRows.length)} of ${activeBarRows.length.toLocaleString()}`;
  $("dialogPrev").disabled = dialogPage === 0;
  $("dialogNext").disabled = dialogPage >= pageCount - 1;
}

function downloadRows(rows, filename) {
  const csvCell = value => `"${String(value ?? "").replace(/"/g, '""')}"`;
  const content = [
    detailColumns.map(([label]) => csvCell(label)).join(","),
    ...rows.map(row => detailColumns.map(([, key]) => csvCell(displayField(row, key))).join(","))
  ].join("\r\n");
  const blob = new Blob(["\uFEFF", content], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function downloadActiveBar() {
  const safeName = activeBarName.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "bar";
  downloadRows(activeBarRows, `lh-performance-${safeName}.csv`);
  $("downloadStatus").textContent = `${activeBarRows.length.toLocaleString()} workbook rows downloaded.`;
}

function render() {
  const rows = currentRows();
  if (!rows.length) {
    $("filterStatus").textContent = "No shipment records found for the selected date range in the active workbook.";
  } else {
    $("filterStatus").textContent = "";
  }
  drawMetrics(rows);
  drawRegion(rows);
  drawRemarks(rows);
  drawRanked(rows, "lane", "laneChart", "Lane", "lane-layout");
  drawRanked(rows, "vendor", "vendorChart", "Vendor", "airline-layout");
  drawConsistentPerformance();
}

function refreshWorkbookData() {
  const script = document.createElement("script");
  const dataUrl = new URL("workbook-data.js", document.baseURI);
  dataUrl.searchParams.set("refresh", Date.now());
  script.src = dataUrl.href;
  script.onload = () => {
    const version = window.offloadWorkbookUpdatedAt;
    if (!version || version === loadedWorkbookVersion) {
      return;
    }
    if (!Array.isArray(window.offloadWorkbookRecords) || !Array.isArray(window.offloadWorkbookColumns)) {
      console.error("Workbook refresh returned invalid dashboard data.");
      return;
    }
    records = window.offloadWorkbookRecords;
    detailColumns = standardDetailColumns.concat(
      window.offloadWorkbookColumns.map(column => [column.label, column.key])
    );
    loadedWorkbookVersion = version;
    updateLastRefreshStamp();
    populateFilters(true);
    updateDateBounds();
    if (!$("startDate").value || !$("endDate").value || $("startDate").value < minDate || $("endDate").value > maxDate) {
      const fallback = activeRangeMode && activeRangeMode !== "all"
        ? buildRangeValues(activeRangeMode)
        : { start: minDate, end: maxDate };
      $("startDate").value = fallback.start;
      $("endDate").value = fallback.end;
    }
    if (!activeRangeMode || activeRangeMode === "all") {
      $("startDate").value = minDate || "";
      $("endDate").value = maxDate || "";
    }
    if ($("barDialog").open) {
      $("barDialog").close();
    }
    render();
  };
  script.onerror = () => console.error("Could not reload workbook-data.js; the dashboard is still showing its last loaded data.");
  document.head.appendChild(script);
}

function toLocalDateString(date) {
  const offset = date.getTimezoneOffset() * 60000;
  return new Date(date.getTime() - offset).toISOString().slice(0, 10);
}

function buildRangeValues(range) {
  if (range === "all") {
    return { start: minDate, end: maxDate };
  }

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  if (range === "today") {
    // no-op: keep start and end on the same local calendar date
  } else if (range === "yesterday") {
    start.setDate(start.getDate() - 1);
    end.setDate(end.getDate() - 1);
  } else if (range === "7") {
    start.setDate(start.getDate() - 6);
  } else if (range === "30") {
    start.setDate(start.getDate() - 29);
  } else if (range === "month") {
    start.setDate(1);
  }

  return {
    start: toLocalDateString(start),
    end: toLocalDateString(end)
  };
}

function applyRange(range) {
  activeRangeMode = range;
  const values = buildRangeValues(range);
  $("startDate").value = values.start;
  $("endDate").value = values.end;
  document.querySelectorAll(".quick-btn[data-range]").forEach(button =>
    button.classList.toggle("active", button.dataset.range === range));
  render();
}

function syncDateRangeState() {
  const start = $("startDate").value;
  const end = $("endDate").value;
  if (!start || !end) {
    return;
  }

  const matchedRange = ["today", "yesterday", "7", "30", "month", "all"].find(range => {
    const values = buildRangeValues(range);
    return values.start === start && values.end === end;
  });

  activeRangeMode = matchedRange || null;
  document.querySelectorAll(".quick-btn[data-range]").forEach(button =>
    button.classList.toggle("active", button.dataset.range === activeRangeMode));
}

function refreshQuickFilterIfNeeded() {
  if (!activeRangeMode || activeRangeMode === "all" || activeRangeMode === null) {
    return;
  }

  const values = buildRangeValues(activeRangeMode);
  if ($("startDate").value !== values.start || $("endDate").value !== values.end) {
    $("startDate").value = values.start;
    $("endDate").value = values.end;
    render();
  }
}

document.querySelectorAll(".quick-btn[data-range]").forEach(button =>
  button.addEventListener("click", () => applyRange(button.dataset.range)));
["startDate", "endDate", "regionFilter", "vendorFilter", "laneFilter", "originFilter", "flightFilter"].forEach(id => {
  $(id).addEventListener("change", () => {
    document.querySelectorAll(".quick-btn[data-range]").forEach(button => button.classList.remove("active"));
    if ($("startDate").value && $("endDate").value && $("startDate").value > $("endDate").value) {
      $("filterStatus").textContent = "Start date must be on or before end date.";
      return;
    }
    activeRangeMode = null;
    syncDateRangeState();
    render();
  });
});
updateLastRefreshStamp();
$("closeDialog").addEventListener("click", () => $("barDialog").close());
$("downloadBar").addEventListener("click", downloadActiveBar);
$("dialogPrev").addEventListener("click", () => { dialogPage--; renderDialogPage(); });
$("dialogNext").addEventListener("click", () => { dialogPage++; renderDialogPage(); });
resetToAllDates();
window.setInterval(() => {
  refreshQuickFilterIfNeeded();
  const today = toLocalDateString(new Date());
  if (today !== watchlistDate) {
    watchlistDate = today;
    render();
  }
  refreshWorkbookData();
}, 60000);
