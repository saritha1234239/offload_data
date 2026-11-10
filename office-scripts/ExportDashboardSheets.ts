function normalizeName(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function main(workbook: ExcelScript.Workbook): string {
  const expectedSheets = [
    "Daily-North-new",
    "Daily-South-new",
    "Daily-West-new",
    "Daily-East-new"
  ];
  const worksheets = workbook.getWorksheets();
  const payload = expectedSheets.map((expectedName) => {
    const sheet = worksheets.find(
      (candidate) => normalizeName(candidate.getName()) === normalizeName(expectedName)
    );
    if (!sheet) {
      throw new Error(`Required dashboard sheet '${expectedName}' was not found.`);
    }

    const usedRange = sheet.getUsedRange(true);
    if (!usedRange) {
      throw new Error(`Dashboard sheet '${sheet.getName()}' is empty.`);
    }
    return {
      name: sheet.getName(),
      values: usedRange.getValues()
    };
  });
  return JSON.stringify(payload);
}
