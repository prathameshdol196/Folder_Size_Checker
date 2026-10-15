param(
    [string]$RootPath   = "FOLDER YOU WANT TO CHECK",
    [string]$OutputPath = "EXCEL FILE CONTAINING FOLDER SIZES WILL BE GENERATED HERE",
    [string]$SheetName  = "FILE NAME"
)

# 1) Validate root path
if (-not (Test-Path -LiteralPath $RootPath)) {
    Write-Error "Root path not found: $RootPath"
    exit 1
}

Write-Host "Scanning folders under: $RootPath" -ForegroundColor Cyan

# 2) Enumerate immediate subfolders
$folders = Get-ChildItem -LiteralPath $RootPath -Directory -ErrorAction Stop
Write-Host "Found $($folders.Count) immediate subfolder(s)." -ForegroundColor Cyan

if (-not $folders -or $folders.Count -eq 0) {
    Write-Warning "No subfolders found under '$RootPath'. The sheet would be empty."
    exit 0
}

# 3) Compute sizes (recursive), ignoring access errors
$results = foreach ($f in $folders) {
    try {
        $bytes = (Get-ChildItem -LiteralPath $f.FullName -Recurse -File -ErrorAction SilentlyContinue |
                  Measure-Object -Property Length -Sum).Sum
        if (-not $bytes) { $bytes = 0 }

        # Convert to MB with two decimals, keep numeric precision for sorting/Excel
        $sizeMB = [math]::Round(($bytes / 1MB), 2)

        [pscustomobject]@{
            'Folder Name' = $f.Name
            'Full Path'   = $f.FullName
            'Size (MB)'   = [double]$sizeMB
        }
    } catch {
        Write-Warning "Failed to scan: $($f.FullName) — $($_.Exception.Message)"
    }
}

# 4) Sort and basic validation
$sorted = $results | Sort-Object 'Size (MB)' -Descending
if (-not $sorted -or $sorted.Count -eq 0) {
    Write-Warning "No size data collected. Possible causes: access denied, empty folders, or empty directories."
    exit 0
}

# Show top 5 in console for verification
Write-Host "Top 5 folders by size (MB):" -ForegroundColor Cyan
$sorted | Select-Object -First 5 | ForEach-Object {
    Write-Host (" - {0} ({1} MB)" -f $_.'Folder Name', $_.'Size (MB)')
}

# 5) Ensure output directory exists
$outDir = Split-Path -Path $OutputPath -Parent
if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

# 6) Launch Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# 7) Open existing workbook or create new
if (Test-Path -LiteralPath $OutputPath) {
    Write-Host "Opening existing workbook: $OutputPath" -ForegroundColor Yellow
    $wb = $excel.Workbooks.Open($OutputPath)
} else {
    Write-Host "Creating new workbook: $OutputPath" -ForegroundColor Yellow
    $wb = $excel.Workbooks.Add()
}

# 8) Get or create target worksheet
$ws = $null
foreach ($sheet in $wb.Worksheets) {
    if ($sheet.Name -eq $SheetName) {
        $ws = $sheet
        break
    }
}
if (-not $ws) {
    $ws = $wb.Worksheets.Add()
    $ws.Name = $SheetName
}
$ws.Activate()

# 9) Clear existing contents (keep formatting)
$ws.Cells.ClearContents() | Out-Null

# 10) Write headers (only 3 columns now)
$headers = @('Folder Name','Full Path','Size (MB)')
for ($i=0; $i -lt $headers.Count; $i++) {
    $ws.Cells.Item(1, $i+1).Value2 = $headers[$i]
}
$ws.Rows.Item(1).Font.Bold = $true
$ws.Rows.Item(1).Interior.ColorIndex = 15  # light gray header

# 11) Write data
$row = 2
foreach ($item in $sorted) {
    $ws.Cells.Item($row, 1).Value2 = $item.'Folder Name'
    $ws.Cells.Item($row, 2).Value2 = $item.'Full Path'
    $ws.Cells.Item($row, 3).Value2 = [double]$item.'Size (MB)'
    $row++
}

# 12) Format columns
$usedRange = $ws.UsedRange
$usedRange.Columns.AutoFit() | Out-Null
$ws.Columns.Item(3).NumberFormat = "0.00"  # show MB with 2 decimals

# 13) Save workbook
if (Test-Path -LiteralPath $OutputPath) {
    $wb.Save()
} else {
    $wb.SaveAs($OutputPath)
}

# 14) Cleanup COM objects
$wb.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "Done. Wrote $($sorted.Count) row(s) to sheet '$SheetName' in: $OutputPath" -ForegroundColor Green