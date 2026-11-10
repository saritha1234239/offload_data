interface DashboardSheet {
  name: string;
  values: (string | number | boolean)[][];
}

function normalizeName(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function main(workbook: ExcelScript.Workbook, serializedSheets: string): void {
  const sheets: DashboardSheet[] = JSON.parse(serializedSheets);
  if (!Array.isArray(sheets) || sheets.length !== 4) {
    throw new Error("Expected data for all four dashboard source sheets.");
  }

  const worksheets = workbook.getWorksheets();
  for (const source of sheets) {
    if (!source.name || !Array.isArray(source.values) || source.values.length === 0) {
      throw new Error("A dashboard source sheet is empty or invalid.");
    }
    const target = worksheets.find(
      (candidate) => normalizeName(candidate.getName()) === normalizeName(source.name)
    );
    if (!target) {
      throw new Error(`Target dashboard sheet '${source.name}' was not found.`);
    }

    const usedRange = target.getUsedRange();
    if (usedRange) {
      usedRange.clear(ExcelScript.ClearApplyTo.contents);
    }
    const rowCount = source.values.length;
    const columnCount = Math.max(...source.values.map((row) => row.length));
    if (!columnCount) {
      throw new Error(`Target data for '${source.name}' has no columns.`);
    }
    const normalizedValues = source.values.map((row) =>
      Array.from({ length: columnCount }, (_, column) => row[column] ?? "")
    );
    target.getRangeByIndexes(0, 0, rowCount, columnCount).setValues(normalizedValues);
  }
}
