param(
  [switch]$Once,
  [string]$WorkbookPath = "https://criticalogindia-my.sharepoint.com/personal/saptarshi_maji_criticalog_com/Documents/Desktop/Midmile%20Management/LH%20performance%20tracker.xlsx?web=1",
  [int]$IntervalSeconds = 60
)

$ErrorActionPreference = "Stop"
$dashboardDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataPath = Join-Path $dashboardDirectory "workbook-data.js"
$expectedSheets = @(
  "Daily-North-new",
  "Daily-South-new",
  "Daily-West-new",
  "Daily-East-new"
)
$utf8WithoutBom = New-Object System.Text.UTF8Encoding($false)

function Resolve-WorkbookPath {
  param([string]$SuggestedPath)

  if ($SuggestedPath) {
    if ($SuggestedPath -match '^https?://') {
      return $SuggestedPath
    }
    if (Test-Path -LiteralPath $SuggestedPath) {
      return (Resolve-Path -LiteralPath $SuggestedPath).Path
    }
  }

  $candidatePaths = [System.Collections.Generic.List[string]]::new()

  $shortcutPath = Join-Path $env:USERPROFILE 'OneDrive - CriticaLog India Pvt Ltd\Downloads\LH performance tracker.xlsx.url'
  if (-not (Test-Path -LiteralPath $shortcutPath)) {
    $shortcutPath = Join-Path $env:USERPROFILE 'OneDrive - CriticaLog India Pvt Ltd\LH performance tracker.xlsx.url'
  }
  if (Test-Path -LiteralPath $shortcutPath) {
    try {
      $shortcutText = Get-Content -LiteralPath $shortcutPath -Raw
      $shortcutTarget = [regex]::Match($shortcutText, '(?im)^URL=(.+)$').Groups[1].Value
      if ($shortcutTarget) {
        $decoded = [System.Uri]::UnescapeDataString($shortcutTarget)
        $relativePath = $decoded -replace '^https?://[^/]+/[^/]+/[^/]+/',''
        $relativePath = $relativePath -replace '\?.*$',''
        if ($relativePath -match '^Documents/.*') {
          $normalized = $relativePath -replace '^Documents/',''
          $candidatePaths.Add((Join-Path $env:USERPROFILE "OneDrive - CriticaLog India Pvt Ltd\Documents\$normalized"))
        }
        $candidatePaths.Add((Join-Path $env:USERPROFILE "OneDrive - CriticaLog India Pvt Ltd\$relativePath"))
      }
    } catch {}
  }

  $searchRoots = @(
    $dashboardDirectory,
    (Join-Path $env:USERPROFILE 'Downloads'),
    (Join-Path $env:USERPROFILE 'OneDrive - CriticaLog India Pvt Ltd\Downloads'),
    (Join-Path $env:USERPROFILE 'OneDrive - CriticaLog India Pvt Ltd'),
    (Join-Path $env:USERPROFILE 'OneDrive - CriticaLog India Pvt Ltd\Documents'),
    (Join-Path $env:USERPROFILE 'OneDrive - CriticaLog India Pvt Ltd\Desktop'),
    (Join-Path $env:USERPROFILE 'OneDrive - CriticaLog India Pvt Ltd\Documents\Desktop')
  )

  foreach ($root in $searchRoots) {
    if (-not (Test-Path -LiteralPath $root)) { continue }
    foreach ($file in Get-ChildItem -LiteralPath $root -Filter 'LH performance tracker*.xlsx' -File -ErrorAction SilentlyContinue) {
      $candidatePaths.Add($file.FullName)
    }
    foreach ($file in Get-ChildItem -LiteralPath $root -Filter '*Midmile*' -Directory -ErrorAction SilentlyContinue) {
      foreach ($xlsx in Get-ChildItem -LiteralPath $file.FullName -Filter 'LH performance tracker*.xlsx' -File -ErrorAction SilentlyContinue) {
        $candidatePaths.Add($xlsx.FullName)
      }
    }
  }

  $uniqueCandidates = $candidatePaths | Select-Object -Unique
  $latestCandidate = $null
  foreach ($candidate in $uniqueCandidates) {
    if (-not (Test-Path -LiteralPath $candidate)) { continue }
    if (-not $latestCandidate) {
      $latestCandidate = $candidate
      continue
    }
    if ((Get-Item -LiteralPath $candidate).LastWriteTimeUtc -gt (Get-Item -LiteralPath $latestCandidate).LastWriteTimeUtc) {
      $latestCandidate = $candidate
    }
  }

  if ($latestCandidate) {
    return $latestCandidate
  }

  if ($SuggestedPath) {
    return $SuggestedPath
  }

  return $null
}

function Normalize-SheetName([string]$name) {
  if ($null -eq $name) { return "" }
  return ($name.ToLowerInvariant() -replace '[^a-z0-9]+', '')
}

function Resolve-SheetName($workbook, [string]$expectedName) {
  $normalizedExpected = Normalize-SheetName $expectedName
  foreach ($sheet in $workbook.Worksheets) {
    if ((Normalize-SheetName $sheet.Name) -eq $normalizedExpected) {
      return $sheet.Name
    }
  }
  foreach ($sheet in $workbook.Worksheets) {
    if ((Normalize-SheetName $sheet.Name) -eq (Normalize-SheetName ($expectedName -replace '-', ' '))) {
      return $sheet.Name
    }
  }
  return $null
}

function Get-ExcelColumnLetter([int]$column) {
  $letter = ""
  while ($column -gt 0) {
    $remainder = ($column - 1) % 26
    $letter = [char](65 + $remainder) + $letter
    $column = [math]::Floor(($column - 1) / 26)
  }
  return $letter
}

function Get-CellValue($values, [int]$row, [int]$column, [int]$rowCount, [int]$columnCount) {
  if ($row -lt 1 -or $row -gt $rowCount -or $column -lt 1 -or $column -gt $columnCount) {
    return $null
  }
  if ($rowCount -eq 1 -and $columnCount -eq 1) {
    return $values
  }
  return $values.GetValue($row, $column)
}

function Get-DisplayCellValue($sheet, [int]$row, [int]$column, [int]$rowCount, [int]$columnCount) {
  if ($row -lt 1 -or $row -gt $rowCount -or $column -lt 1 -or $column -gt $columnCount) {
    return $null
  }
  try {
    $cell = $sheet.Cells.Item($row, $column)
    $formatted = $cell.Text
    if ($null -ne $formatted -and [string]$formatted -ne "") {
      return $formatted
    }

    $rawValue = $cell.Value2
    if ($null -ne $rawValue) {
      if ($rawValue -is [ValueType] -and -not ($rawValue -is [bool])) {
        return $rawValue
      }
      $text = [string]$rawValue
      if ($text.Trim()) {
        return $text
      }
    }
  } catch {}
  return Get-CellValue $sheet.UsedRange.Value2 $row $column $rowCount $columnCount
}

function Get-TextValue($value) {
  if ($null -eq $value) {
    return ""
  }
  return ([string]$value).Trim()
}

function Get-NumberValue($value) {
  if ($null -eq $value) {
    return $null
  }
  if ($value -is [string]) {
    $text = $value.Trim()
    if (-not $text -or $text -eq "NA" -or $text -eq "�") {
      return $null
    }
    $parsed = 0.0
    if (-not [double]::TryParse($text.Replace(",", ""), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
      return $null
    }
    $value = $parsed
  }
  if ($value -is [bool] -or -not ($value -is [ValueType])) {
    return $null
  }
  return [double]$value
}

function Get-DateValue($value, [string]$region = "") {
  return (Get-DateVariants $value $region | Select-Object -First 1)
}

function Get-DateVariants($value, [string]$region = "") {
  $variants = [System.Collections.Generic.List[string]]::new()

  if ($null -eq $value -or [string]$value -eq "") {
    return @($variants.ToArray())
  }
  if ($value -is [ValueType] -and -not ($value -is [bool])) {
    try {
      $serial = [double]$value
      if ($serial -ge 1000 -and $serial -le 2958465) {
        $variants.Add([datetime]::FromOADate($serial).ToString("yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture))
        return @($variants.ToArray())
      }
    } catch {}
  }

  $text = ([string]$value).Trim()
  if (-not $text) {
    return @($variants.ToArray())
  }

  $text = $text -replace '^(?:\s|\u200B|\uFEFF)+', ''
  $text = $text -replace '(?:\s|\u200B|\uFEFF)+$', ''

  if ($text -match '^(Date|Departure_Date|ETD|ETA|Org|Dest|Mawb|Vendor_name|Flight_Number|Remarks|NDC_Remark|NDC_Remarks|Column_.*)$' -or $text -match '^[A-Za-z\s_\-/]+$') {
    return @($variants.ToArray())
  }

  $numericSerial = $null
  if ($text -match '^(?:[+-]?)\d+(?:\.\d+)?$') {
    $numericSerial = 0.0
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$numericSerial)) {
      $minOADate = [datetime]::FromOADate(0)
      $maxOADate = [datetime]::FromOADate(2958465)
      if ($numericSerial -ge 1000 -and $numericSerial -le 2958465) {
        try {
          $variants.Add([datetime]::FromOADate($numericSerial).ToString("yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture))
          return @($variants.ToArray())
        } catch {}
      }
    }
  }

  $normalizedText = $text -replace '\s+', ' '
  $candidateOverrides = [System.Collections.Generic.List[string]]::new()
  $ambiguousDateMatch = [regex]::Match($text, '^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$')
  if ($ambiguousDateMatch.Success) {
    $dayOrMonth = $ambiguousDateMatch.Groups[1].Value
    $monthOrDay = $ambiguousDateMatch.Groups[2].Value
    $year = $ambiguousDateMatch.Groups[3].Value
    if ($year.Length -eq 2) {
      $year = if ($year -match '^(\d+)$') { [int]$year } else { $year }
    }
    if ($dayOrMonth -and $monthOrDay -and $year) {
      $candidateOverrides.Add("$dayOrMonth/$monthOrDay/$year")
      $candidateOverrides.Add("$monthOrDay/$dayOrMonth/$year")
    }
  }

  $datePatterns = @(
    "yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd",
    "dd-MM-yyyy", "dd/MM/yyyy", "dd.MM.yyyy",
    "d-M-yyyy", "d/M/yyyy", "d.M.yyyy",
    "dd-M-yyyy", "dd/M/yyyy", "dd.M.yyyy",
    "MM-dd-yyyy", "MM/dd/yyyy", "MM.dd.yyyy",
    "M-d-yyyy", "M/d/yyyy", "M.d.yyyy",
    "MM-dd-yy", "MM/dd/yy", "MM.dd.yy",
    "M-d-yy", "M/d/yy", "M.d.yy",
    "dd-MM-yy", "dd/MM/yy", "dd.MM.yy",
    "d-M-yy", "d/M/yy", "d.M.yy"
  )

  $primaryFormats = @()
  $isNorthRegion = $region -match '(?i)north'
  if ($isNorthRegion) {
    $primaryFormats += @("MM-dd-yyyy", "MM/dd/yyyy", "MM.dd.yyyy", "M-d-yyyy", "M/d/yyyy", "M.d.yyyy")
    $primaryFormats += @("dd-MM-yyyy", "dd/MM/yyyy", "dd.MM.yyyy", "d-M-yyyy", "d/M/yyyy", "d.M.yyyy")
  } else {
    $primaryFormats += @("dd-MM-yyyy", "dd/MM/yyyy", "dd.MM.yyyy", "d-M-yyyy", "d/M/yyyy", "d.M.yyyy")
    $primaryFormats += @("MM-dd-yyyy", "MM/dd/yyyy", "MM.dd.yyyy", "M-d-yyyy", "M/d/yyyy", "M.d.yyyy")
  }

  $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
  $tryFormats = @($primaryFormats + $datePatterns)

  $candidateSet = @($normalizedText, $text) + @($candidateOverrides.ToArray())
  foreach ($candidate in $candidateSet) {
    foreach ($culture in @(
      [Globalization.CultureInfo]::GetCultureInfo("en-GB"),
      [Globalization.CultureInfo]::InvariantCulture,
      [Globalization.CultureInfo]::GetCultureInfo("en-US")
    )) {
      foreach ($format in $tryFormats) {
        $parsedDate = [datetime]::MinValue
        if ([datetime]::TryParseExact($candidate, $format, $culture, [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$parsedDate)) {
          $iso = $parsedDate.ToString("yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
          if (-not $seen.Contains($iso)) {
            $seen.Add($iso) | Out-Null
            $variants.Add($iso)
          }
        }
      }
    }
  }

  foreach ($culture in @(
    [Globalization.CultureInfo]::GetCultureInfo("en-GB"),
    [Globalization.CultureInfo]::InvariantCulture,
    [Globalization.CultureInfo]::GetCultureInfo("en-US")
  )) {
    $parsedAsCurrentCulture = [datetime]::MinValue
    if ([datetime]::TryParse($text, $culture, [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$parsedAsCurrentCulture)) {
      $iso = $parsedAsCurrentCulture.ToString("yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
      if (-not $seen.Contains($iso)) {
        $seen.Add($iso) | Out-Null
        $variants.Add($iso)
      }
    }
  }

  if ($variants.Count -gt 0) {
    return @($variants.ToArray())
  }

  return @()
}

function Get-TimeValue($value) {
  if ($null -eq $value -or [string]$value -eq "") {
    return $null
  }
  if ($value -is [ValueType] -and -not ($value -is [bool])) {
    return [double]$value
  }
  $text = ([string]$value).Trim()
  if ($text -eq "NA" -or $text -eq "�") {
    return $null
  }
  $parsedTime = [datetime]::MinValue
  if ([datetime]::TryParse($text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedTime)) {
    return ($parsedTime.Hour * 3600 + $parsedTime.Minute * 60 + $parsedTime.Second) / 86400
  }
  return $null
}

function Get-HeaderToken([string]$header, [string]$columnLetter) {
  $clean = $header.Trim()
  if (-not $clean) {
    return "Column_$columnLetter"
  }
  $token = [regex]::Replace($clean, "[^\p{L}\p{N}_]+", "_").Trim("_")
  if (-not $token) {
    return "Column_$columnLetter"
  }
  return $token
}

function Get-FirstHeaderColumn($headerColumns, [string[]]$aliases) {
  foreach ($alias in $aliases) {
    if ($headerColumns.ContainsKey($alias)) {
      return [int]$headerColumns[$alias][0]
    }
  }
  return 0
}

function Is-HeaderRow($values, [int]$maxColumns = 20) {
  $tokens = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($token in @("departure_date", "date", "org", "dest", "mawb", "vendor_name", "flight_number", "etd", "eta", "bags", "cd_weight", "total_cons", "offloaded_bags", "impacted_cons", "next_flight", "remarks", "ndc_remark", "ndc_remarks")) {
    [void]$tokens.Add($token)
  }

  $matches = 0
  for ($index = 1; $index -le [Math]::Min($maxColumns, $values.Count); $index++) {
    $cell = $values[$index - 1]
    if ($null -eq $cell) { continue }
    $text = ([string]$cell).Trim()
    if (-not $text) { continue }
    $normalized = $text.ToLowerInvariant() -replace '[^a-z0-9]+', '_'
    if ($tokens.Contains($normalized)) {
      $matches++
      if ($matches -ge 3) {
        return $true
      }
    }
  }
  return $false
}

function Get-WorkbookData {
  $WorkbookPath = Resolve-WorkbookPath -SuggestedPath $WorkbookPath
  if (-not $WorkbookPath) {
    throw "Workbook not found for the dashboard. Check the workbook path or save LH performance tracker.xlsx in the Downloads folder."
  }
  $isRemoteWorkbook = $WorkbookPath -match '^https?://'
  if (-not $isRemoteWorkbook -and -not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found at $WorkbookPath."
  }
  Write-Host "Using workbook: $WorkbookPath"
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $workbook = $null
  try {
    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)

    $resolvedSheetNames = [System.Collections.Generic.List[string]]::new()
    foreach ($expectedSheet in $expectedSheets) {
      $resolved = Resolve-SheetName $workbook $expectedSheet
      if (-not $resolved) {
        throw "Expected source sheet '$expectedSheet' was not found."
      }
      $resolvedSheetNames.Add($resolved)
    }

    $records = [System.Collections.Generic.List[object]]::new()
    $columnsByKey = [ordered]@{}
    $sources = [System.Collections.Generic.List[object]]::new()

    foreach ($sheetName in $resolvedSheetNames) {
    $sheet = $workbook.Worksheets.Item($sheetName)
    $range = $sheet.UsedRange
    $rowCount = [int]$range.Rows.Count
    $columnCount = [int]$range.Columns.Count
    if ($range.Row -ne 1 -or $range.Column -ne 1 -or $rowCount -lt 1) {
      throw "Unexpected used range in '$sheetName'; expected headers in the first row."
    }
    $values = $range.Value2
    $headers = [System.Collections.Generic.List[string]]::new()
    $headerColumns = @{}
    for ($column = 1; $column -le $columnCount; $column++) {
      $header = Get-TextValue (Get-CellValue $values 1 $column $rowCount $columnCount)
      $headers.Add($header)
      $letter = Get-ExcelColumnLetter $column
      $token = Get-HeaderToken $header $letter
      $key = "raw_${column}_${token}"
      $label = if ($header) { "$header ($letter)" } else { "Column $letter ($letter)" }
      if (-not $columnsByKey.Contains($key)) {
        $columnsByKey[$key] = [ordered]@{ key = $key; label = $label }
      }
      $normalizedHeader = $header.ToLowerInvariant() -replace "[^a-z0-9]+", "_"
      if ($normalizedHeader) {
        if (-not $headerColumns.ContainsKey($normalizedHeader)) {
          $headerColumns[$normalizedHeader] = [System.Collections.Generic.List[int]]::new()
        }
        $headerColumns[$normalizedHeader].Add($column)
      }
    }

    $dateColumn = Get-FirstHeaderColumn $headerColumns @("date", "departure_date")
    $originColumn = Get-FirstHeaderColumn $headerColumns @("org", "origin")
    $destinationColumn = Get-FirstHeaderColumn $headerColumns @("dest", "destination")
    $laneColumn = Get-FirstHeaderColumn $headerColumns @("lane")
    $mawbColumn = Get-FirstHeaderColumn $headerColumns @("mawb")
    $vendorColumn = Get-FirstHeaderColumn $headerColumns @("vendor_name", "vendor")
    $flightColumn = Get-FirstHeaderColumn $headerColumns @("flight_number", "flight")
    $bagsColumn = Get-FirstHeaderColumn $headerColumns @("bags")
    $weightColumn = Get-FirstHeaderColumn $headerColumns @("cd_weight")
    $totalConsColumn = Get-FirstHeaderColumn $headerColumns @("total_cons")
    $offloadedColumn = Get-FirstHeaderColumn $headerColumns @("offloaded_bags")
    $impactedColumn = Get-FirstHeaderColumn $headerColumns @("impacted_cons")
    $nextFlightColumn = Get-FirstHeaderColumn $headerColumns @("next_flight")
    $remarksColumn = Get-FirstHeaderColumn $headerColumns @("remarks")
    $ndcRemarksColumn = Get-FirstHeaderColumn $headerColumns @("ndc_remark", "ndc_remarks")
    $etdColumns = if ($headerColumns.ContainsKey("etd")) { $headerColumns["etd"] } else { @() }
    $etaColumns = if ($headerColumns.ContainsKey("eta")) { $headerColumns["eta"] } else { @() }
    $etdColumn = if ($etdColumns.Count -ge 1) { [int]$etdColumns[0] } else { 0 }
    $nextEtdColumn = if ($etdColumns.Count -ge 2) { [int]$etdColumns[1] } else { 0 }
    $etaColumn = if ($etaColumns.Count -ge 1) { [int]$etaColumns[0] } else { 0 }
    $nextEtaColumn = if ($etaColumns.Count -ge 2) { [int]$etaColumns[1] } else { 0 }

    if (-not $dateColumn -or -not $originColumn -or -not $destinationColumn -or -not $bagsColumn -or -not $offloadedColumn) {
      throw "Required shipment columns are missing from '$sheetName'."
    }

    $region = if ($sheetName -like "*North*") { "North" } elseif ($sheetName -like "*South*") { "South" } elseif ($sheetName -like "*West*") { "West" } elseif ($sheetName -like "*East*") { "East" } else { "Unspecified" }
    $sheetRecords = [System.Collections.Generic.List[object]]::new()
    $dates = [System.Collections.Generic.List[string]]::new()
    for ($row = 2; $row -le $rowCount; $row++) {
      $sourceValues = [ordered]@{}
      $hasValues = $false
      for ($column = 1; $column -le $columnCount; $column++) {
        $value = Get-CellValue $values $row $column $rowCount $columnCount
        if ($null -ne $value -and ([string]$value).Trim() -ne "") {
          $hasValues = $true
        }
        $header = $headers[$column - 1]
        $key = "raw_${column}_$(Get-HeaderToken $header (Get-ExcelColumnLetter $column))"
        $sourceValues[$key] = $value
      }
      if (-not $hasValues) {
        continue
      }

      $rowCells = @()
      for ($column = 1; $column -le $columnCount; $column++) {
        $rowCells += Get-CellValue $values $row $column $rowCount $columnCount
      }

      $getCell = {
        param([int]$column)
        if (-not $column) { return $null }
        return Get-CellValue $values $row $column $rowCount $columnCount
      }
      $dateValue = Get-DisplayCellValue $sheet $row $dateColumn $rowCount $columnCount
      $dateVariants = Get-DateVariants $dateValue $region
      $recordDate = if ($dateVariants.Count -gt 0) { $dateVariants[0] } else { "" }

      if (Is-HeaderRow $rowCells 20 -or ($recordDate -eq "" -and $rowCells -match '(?i)^(?:date|origin|dest|mawb|vendor|flight|bags|offloaded|remarks|ndc)')) {
        continue
      }

      $origin = Get-TextValue (& $getCell $originColumn)
      $destination = Get-TextValue (& $getCell $destinationColumn)
      $mawb = Get-TextValue (& $getCell $mawbColumn)
      $vendor = Get-TextValue (& $getCell $vendorColumn)
      $flight = Get-TextValue (& $getCell $flightColumn)
      $bags = Get-NumberValue (& $getCell $bagsColumn)
      $cdWeight = Get-NumberValue (& $getCell $weightColumn)
      $totalCons = Get-NumberValue (& $getCell $totalConsColumn)
      $offloadedBags = Get-NumberValue (& $getCell $offloadedColumn)
      $impactedCons = Get-NumberValue (& $getCell $impactedColumn)
      $nextFlight = Get-TextValue (& $getCell $nextFlightColumn)
      $recordType = if ($mawb -or $vendor -or $flight -or $null -ne $bags -or $null -ne $cdWeight -or $null -ne $totalCons -or $null -ne $offloadedBags) { "Shipment" } else { "Worksheet detail" }
      $lane = if ($laneColumn) { Get-TextValue (& $getCell $laneColumn) } elseif ($origin -and $destination) { "$origin-$destination" } else { "" }
      $record = [pscustomobject][ordered]@{
        sourceSheet = $sheetName
        sourceRow = $row
        recordType = $recordType
        date = $recordDate
        dateVariants = @($dateVariants)
        region = $region
        origin = $origin
        destination = $destination
        lane = $lane
        mawb = $mawb
        vendor = $vendor
        flight = $flight
        etd = Get-TimeValue (& $getCell $etdColumn)
        eta = Get-TimeValue (& $getCell $etaColumn)
        bags = $bags
        cdWeight = $cdWeight
        totalCons = $totalCons
        offloadedBags = $offloadedBags
        impactedCons = $impactedCons
        nextFlight = $nextFlight
        nextEtd = Get-TimeValue (& $getCell $nextEtdColumn)
        nextEta = Get-TimeValue (& $getCell $nextEtaColumn)
        remarks = Get-TextValue (& $getCell $remarksColumn)
        ndcRemarks = Get-TextValue (& $getCell $ndcRemarksColumn)
        sourceValues = $sourceValues
      }
      $records.Add($record)
      $sheetRecords.Add($record)
      if ($recordDate) { $dates.Add($recordDate) }
    }

    $sources.Add([pscustomobject][ordered]@{
      name = $sheetName
      rows = $sheetRecords.Count
      shipments = @($sheetRecords | Where-Object { $_.recordType -eq "Shipment" }).Count
      worksheetDetails = @($sheetRecords | Where-Object { $_.recordType -ne "Shipment" }).Count
      undated = @($sheetRecords | Where-Object { -not $_.date }).Count
      minDate = if ($dates.Count) { ($dates | Sort-Object | Select-Object -First 1) } else { "" }
      maxDate = if ($dates.Count) { ($dates | Sort-Object -Descending | Select-Object -First 1) } else { "" }
    })
  }

    return [pscustomobject]@{
      columns = @($columnsByKey.Values)
      records = @($records.ToArray())
      sources = @($sources.ToArray())
    }
  } finally {
    if ($null -ne $workbook) {
      $workbook.Close($false)
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($null -ne $excel) {
      $excel.Quit()
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
  }
}

function Write-WorkbookData {
  $data = Get-WorkbookData
  $columnsJson = ConvertTo-Json -InputObject @($data.columns) -Depth 100 -Compress
  $recordsJson = ConvertTo-Json -InputObject @($data.records) -Depth 100 -Compress
  $sourcesJson = ConvertTo-Json -InputObject @($data.sources) -Depth 100 -Compress
  $sha = [Security.Cryptography.SHA256]::Create()
  try {
    $hashBytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($recordsJson))
    $version = [BitConverter]::ToString($hashBytes).Replace("-", "").ToLowerInvariant()
  } finally {
    $sha.Dispose()
  }
  $content = @(
    "window.offloadWorkbookUpdatedAt = `"$version`";"
    "window.offloadWorkbookColumns = $columnsJson;"
    "window.offloadWorkbookRecords = $recordsJson;"
    "window.offloadWorkbookSources = $sourcesJson;"
    "window.offloadWorkbookLists = {};"
  ) -join [Environment]::NewLine
  $temporaryPath = "$dataPath.$PID.tmp"
  [System.IO.File]::WriteAllText($temporaryPath, $content, $utf8WithoutBom)
  Move-Item -LiteralPath $temporaryPath -Destination $dataPath -Force
  $shipmentCount = @($data.records | Where-Object { $_.recordType -eq "Shipment" }).Count
  Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Refreshed $($data.records.Count) rows ($shipmentCount shipments); version $($version.Substring(0, 12))."
  return $version
}

if ($Once) {
  Write-WorkbookData | Out-Null
  exit 0
}

Write-Host "Refreshing dashboard from the SharePoint workbook every $IntervalSeconds seconds. Press Ctrl+C to stop."
while ($true) {
  try {
    Write-WorkbookData | Out-Null
  } catch {
    Write-Host "Workbook refresh failed; the last good dashboard data was kept. $($_.Exception.Message)" -ForegroundColor Red
  }
  Start-Sleep -Seconds $IntervalSeconds
}
