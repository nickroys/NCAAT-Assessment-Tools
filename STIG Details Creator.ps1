# =============================================================================
# Split-STIGDetails.ps1
#
# Reads the 'STIG Details' sheet from a Vulnerator-generated Excel file,
# finds all unique file names in column S, and creates a new sheet for each
# one containing the header row + all matching data rows.
# =============================================================================

# --------------------------------------------------------------------------
# Helper: sanitize a string so it's safe to use as an Excel sheet name
# --------------------------------------------------------------------------
function Get-SafeSheetName {
    param([string]$Name)

    # Excel sheet names: max 31 chars, no \ / ? * [ ] :
    $safe = $Name -replace '[\\\/\?\*\[\]\:]', '_'
    if ($safe.Length -gt 31) { $safe = $safe.Substring(0, 31) }
    return $safe.Trim()
}

# --------------------------------------------------------------------------
# 1.  File picker dialog
# --------------------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title  = "Select Vulnerator Excel File"
$dialog.Filter = "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All Files (*.*)|*.*"

if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "No file selected. Exiting." -ForegroundColor Yellow
    exit
}

$filePath = $dialog.FileName
Write-Host "`nOpening: $filePath" -ForegroundColor Cyan

# --------------------------------------------------------------------------
# 2.  Open Excel via COM
# --------------------------------------------------------------------------
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible        = $false   # run silently; change to $true to watch
    $excel.DisplayAlerts  = $false
}
catch {
    Write-Error "Could not launch Excel. Make sure Microsoft Excel is installed."
    exit 1
}

try {
    $workbook = $excel.Workbooks.Open($filePath)
}
catch {
    Write-Error "Could not open the file: $_"
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 1
}

# --------------------------------------------------------------------------
# 3.  Locate the 'STIG Details' sheet
# --------------------------------------------------------------------------
$sourceSheet = $null
foreach ($ws in $workbook.Worksheets) {
    if ($ws.Name -eq 'STIG Details') {
        $sourceSheet = $ws
        break
    }
}

if ($null -eq $sourceSheet) {
    Write-Error "'STIG Details' sheet not found in the workbook."
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 1
}

# --------------------------------------------------------------------------
# 4.  Read all data from 'STIG Details'
# --------------------------------------------------------------------------
$usedRange  = $sourceSheet.UsedRange
$totalRows  = $usedRange.Rows.Count
$totalCols  = $usedRange.Columns.Count

# Column S = index 19 (A=1, B=2 … S=19)
$fileNameColIndex = 19

Write-Host "Found $($totalRows - 1) data rows across $totalCols columns." -ForegroundColor Cyan

# Grab every cell value into a 2-D array for fast access
# ComObject arrays are 1-based
$allValues = $usedRange.Value2

if ($totalRows -lt 2) {
    Write-Host "No data rows found in 'STIG Details'. Nothing to do." -ForegroundColor Yellow
    $workbook.Close($false)
    $excel.Quit()
    exit
}

# --------------------------------------------------------------------------
# 5.  Collect unique file names (column S, rows 2-N)
# --------------------------------------------------------------------------
$uniqueNames = [System.Collections.Generic.SortedSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

for ($r = 2; $r -le $totalRows; $r++) {
    $val = $allValues[$r, $fileNameColIndex]
    if (![string]::IsNullOrWhiteSpace($val)) {
        $uniqueNames.Add($val.ToString().Trim()) | Out-Null
    }
}

Write-Host "Unique file names found: $($uniqueNames.Count)" -ForegroundColor Cyan
$uniqueNames | ForEach-Object { Write-Host "  - $_" }

# --------------------------------------------------------------------------
# 6.  Pre-count rows per file name so per-sheet progress is accurate
# --------------------------------------------------------------------------
$rowCountMap = @{}
foreach ($name in $uniqueNames) { $rowCountMap[$name] = 0 }

for ($r = 2; $r -le $totalRows; $r++) {
    $val = $allValues[$r, $fileNameColIndex]
    if (![string]::IsNullOrWhiteSpace($val)) {
        $key = $val.ToString().Trim()
        if ($rowCountMap.ContainsKey($key)) { $rowCountMap[$key]++ }
    }
}

# --------------------------------------------------------------------------
# 7.  For each unique file name: create a new sheet and copy matching rows
# --------------------------------------------------------------------------
$totalSheets  = $uniqueNames.Count
$sheetCounter = 0

foreach ($fileName in $uniqueNames) {

    $sheetCounter++
    $sheetName  = Get-SafeSheetName -Name $fileName
    $overallPct = [math]::Round(($sheetCounter - 1) / $totalSheets * 100)

    # -- Overall progress bar (ID 1) --
    Write-Progress -Id 1 `
        -Activity "Processing sheets" `
        -Status "Sheet $sheetCounter of $totalSheets : $sheetName" `
        -PercentComplete $overallPct

    # If a sheet with this name already exists, delete it first so we get
    # a clean copy (useful when re-running the script on the same file).
    foreach ($ws in @($workbook.Worksheets)) {
        if ($ws.Name -eq $sheetName -and $ws.Name -ne 'STIG Details') {
            $ws.Delete()
            break
        }
    }

    # Create the new sheet at the end of the workbook
    $newSheet      = $workbook.Worksheets.Add([System.Reflection.Missing]::Value,
                         $workbook.Worksheets[$workbook.Worksheets.Count])
    $newSheet.Name = $sheetName

    # -- Copy the header row (row 1) --
    $srcHeaderRange = $sourceSheet.Range(
        $sourceSheet.Cells(1, 1),
        $sourceSheet.Cells(1, $totalCols)
    )
    $dstHeaderRange = $newSheet.Range(
        $newSheet.Cells(1, 1),
        $newSheet.Cells(1, $totalCols)
    )
    $srcHeaderRange.Copy($dstHeaderRange) | Out-Null

    # -- Copy matching data rows --
    $destRow      = 2
    $expectedRows = $rowCountMap[$fileName]
    $copiedRows   = 0

    for ($r = 2; $r -le $totalRows; $r++) {
        $cellVal = $allValues[$r, $fileNameColIndex]
        if (![string]::IsNullOrWhiteSpace($cellVal) -and
            $cellVal.ToString().Trim() -ieq $fileName) {

            $srcRowRange = $sourceSheet.Range(
                $sourceSheet.Cells($r, 1),
                $sourceSheet.Cells($r, $totalCols)
            )
            $dstRowRange = $newSheet.Range(
                $newSheet.Cells($destRow, 1),
                $newSheet.Cells($destRow, $totalCols)
            )
            $srcRowRange.Copy($dstRowRange) | Out-Null
            $destRow++
            $copiedRows++

            # -- Per-sheet progress bar (ID 2) --
            if ($expectedRows -gt 0) {
                $rowPct = [math]::Round($copiedRows / $expectedRows * 100)
                Write-Progress -Id 2 -ParentId 1 `
                    -Activity "Copying rows for: $sheetName" `
                    -Status "Row $copiedRows of $expectedRows" `
                    -PercentComplete $rowPct
            }
        }
    }

    # Clear the per-sheet bar when this sheet is done
    Write-Progress -Id 2 -ParentId 1 -Activity "Copying rows for: $sheetName" -Completed

    # Auto-fit columns for readability
    $newSheet.UsedRange.EntireColumn.AutoFit() | Out-Null

    Write-Host "  [Sheet $sheetCounter/$totalSheets] Created '$sheetName' -- $copiedRows rows" -ForegroundColor Green
}

# Clear the overall bar
Write-Progress -Id 1 -Activity "Processing sheets" -Completed

# --------------------------------------------------------------------------
# 7.  Save and close
# --------------------------------------------------------------------------
$workbook.Save()
$workbook.Close($false)
$excel.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)    | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "`nDone! Sheets created and file saved." -ForegroundColor Green

# --------------------------------------------------------------------------
# 8.  Optional: rename the file
# --------------------------------------------------------------------------
$rename = Read-Host "`nWould you like to rename the file? (Y/N)"

if ($rename -match '^[Yy]') {
    $fileDir      = [System.IO.Path]::GetDirectoryName($filePath)
    $fileExt      = [System.IO.Path]::GetExtension($filePath)
    $currentName  = [System.IO.Path]::GetFileNameWithoutExtension($filePath)

    Write-Host "Current file name: $currentName" -ForegroundColor Cyan
    $newName = Read-Host "Enter new file name (without extension)"

    if ([string]::IsNullOrWhiteSpace($newName)) {
        Write-Host "No name entered. File was not renamed." -ForegroundColor Yellow
    }
    else {
        # Strip any extension the user may have accidentally typed
        $newName     = [System.IO.Path]::GetFileNameWithoutExtension($newName)
        $newFilePath = [System.IO.Path]::Combine($fileDir, "$newName$fileExt")

        if (Test-Path $newFilePath) {
            Write-Host "A file named '$newName$fileExt' already exists in that folder. File was not renamed." -ForegroundColor Red
        }
        else {
            Rename-Item -Path $filePath -NewName "$newName$fileExt"
            Write-Host "File renamed to: $newName$fileExt" -ForegroundColor Green
        }
    }
}
