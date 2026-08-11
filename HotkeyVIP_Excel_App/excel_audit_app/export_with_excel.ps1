param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$PayloadPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
Add-Type -AssemblyName System.Drawing

function Get-ExcelColor([string]$Hex) {
    $clean = $Hex.TrimStart('#')
    $red = [Convert]::ToInt32($clean.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($clean.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($clean.Substring(4, 2), 16)
    return [System.Drawing.ColorTranslator]::ToOle(
        [System.Drawing.Color]::FromArgb($red, $green, $blue)
    )
}

function Get-WorksheetCaseInsensitive($Workbook, [string]$Name) {
    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $candidate = $Workbook.Worksheets.Item($index)
        if ([string]::Equals($candidate.Name, $Name, [StringComparison]::OrdinalIgnoreCase)) {
            return $candidate
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($candidate)
    }
    return $null
}

function Release-ComObject($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

$payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
$excel = $null
$workbook = $null
$tongAll = $null
$keHoach = $null
$activeWindow = $null
$targetRange = $null
$statusRange = $null
$oldCalculation = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.DisplayStatusBar = $false
    $excel.AutomationSecurity = 3

    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $false)
    $workbook.CheckCompatibility = $false
    $oldCalculation = $excel.Calculation
    $excel.Calculation = -4135
    $excel.CalculateBeforeSave = $false

    # Only update the source-status column in KE_HOACH.
    $keConfig = $payload.ke_hoach
    $keHoach = Get-WorksheetCaseInsensitive $workbook ([string]$keConfig.sheet_name)
    if ($null -eq $keHoach) {
        throw "Cannot find KE_HOACH in the working copy."
    }

    $statusColumn = 0
    if ($null -ne $keConfig.status_column) {
        $statusColumn = [int]$keConfig.status_column
    }
    if ($statusColumn -le 0) {
        $statusColumn = [int]$keConfig.max_column + 1
    }
    $headerRow = [int]$keConfig.header_row
    $lastDataRow = [int]$keConfig.last_data_row
    if ($lastDataRow -lt ($headerRow + 1)) {
        $lastDataRow = $headerRow + 1
    }

    $headerCell = $keHoach.Cells.Item($headerRow, $statusColumn)
    if ([string]::IsNullOrWhiteSpace([string]$headerCell.Value2)) {
        $headerCell.Value2 = [string]$keConfig.status_header
    }
    Release-ComObject $headerCell

    $firstStatusRow = $headerRow + 1
    $statusRange = $keHoach.Range(
        $keHoach.Cells.Item($firstStatusRow, $statusColumn),
        $keHoach.Cells.Item($lastDataRow, $statusColumn)
    )
    $statusRowCount = $lastDataRow - $firstStatusRow + 1
    $statusValues = $statusRange.Formula
    if ($statusRowCount -eq 1) {
        $statusArray = New-Object 'object[,]' 1, 1
        $statusArray[0, 0] = $statusValues
        $arrayOffset = 0
    }
    else {
        $statusArray = $statusValues
        $arrayOffset = $statusArray.GetLowerBound(0)
        $columnOffset = $statusArray.GetLowerBound(1)
    }

    for ($offset = 0; $offset -lt $statusRowCount; $offset++) {
        $arrayRow = $offset + $arrayOffset
        if ($statusRowCount -eq 1) {
            $current = [string]$statusArray[0, 0]
        }
        else {
            $current = [string]$statusArray[$arrayRow, $columnOffset]
        }
        $isOwnedStatus = $false
        foreach ($ownedStatus in @($keConfig.owned_statuses)) {
            if ($current -eq [string]$ownedStatus) {
                $isOwnedStatus = $true
                break
            }
        }
        if ($isOwnedStatus) {
            if ($statusRowCount -eq 1) {
                $statusArray[0, 0] = $null
            }
            else {
                $statusArray[$arrayRow, $columnOffset] = $null
            }
        }
    }

    foreach ($update in @($keConfig.updates)) {
        $rowNumber = [int]$update.row
        if ($rowNumber -lt $firstStatusRow -or $rowNumber -gt $lastDataRow) {
            continue
        }
        $offset = $rowNumber - $firstStatusRow
        if ($statusRowCount -eq 1) {
            $statusArray[0, 0] = [string]$update.status
        }
        else {
            $statusArray[$offset + $arrayOffset, $columnOffset] = [string]$update.status
        }
    }
    $statusRange.Formula = $statusArray

    # Create or refresh Tong_all without changing other sheets.
    $tongAll = Get-WorksheetCaseInsensitive $workbook 'Tong_all'
    if ($null -eq $tongAll) {
        $tongAll = $workbook.Worksheets.Add()
        $tongAll.Name = 'Tong_all'
    }
    else {
        try { $tongAll.Cells.UnMerge() } catch {}
        while ($tongAll.ListObjects.Count -gt 0) {
            $tongAll.ListObjects.Item(1).Delete()
        }
        while ($tongAll.Shapes.Count -gt 0) {
            $tongAll.Shapes.Item(1).Delete()
        }
        $tongAll.UsedRange.Clear()
    }

    $report = $payload.report
    $rowCount = [int]$report.row_count
    $columnCount = [int]$report.column_count
    $lastColumn = [string]$report.last_column
    if ([string]::IsNullOrWhiteSpace($lastColumn)) {
        $lastColumn = 'K'
    }
    $matrix = New-Object 'object[,]' $rowCount, $columnCount
    for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
        $sourceRow = @($report.rows[$rowIndex])
        for ($columnIndex = 0; $columnIndex -lt $columnCount; $columnIndex++) {
            if ($columnIndex -lt $sourceRow.Count) {
                $matrix[$rowIndex, $columnIndex] = $sourceRow[$columnIndex]
            }
        }
    }

    $targetRange = $tongAll.Range(
        $tongAll.Cells.Item(1, 1),
        $tongAll.Cells.Item($rowCount, $columnCount)
    )
    $targetRange.Value2 = $matrix
    $targetRange.Font.Name = 'Aptos'
    $targetRange.Font.Size = 10
    $targetRange.VerticalAlignment = -4108
    $targetRange.RowHeight = 20

    foreach ($mergeAddress in @($report.merges)) {
        $tongAll.Range([string]$mergeAddress).Merge()
    }

    $navy = Get-ExcelColor '#17324D'
    $teal = Get-ExcelColor '#0F766E'
    $blue = Get-ExcelColor '#315B7D'
    $lightBlue = Get-ExcelColor '#DCEAF5'
    $lightGray = Get-ExcelColor '#F1F5F9'
    $green = Get-ExcelColor '#DCFCE7'
    $red = Get-ExcelColor '#FEE2E2'
    $white = Get-ExcelColor '#FFFFFF'
    $darkText = Get-ExcelColor '#243047'

    $titleRange = $tongAll.Range("A1:${lastColumn}1")
    $titleRange.Interior.Color = $navy
    $titleRange.Font.Color = $white
    $titleRange.Font.Bold = $true
    $titleRange.Font.Size = 18
    $titleRange.HorizontalAlignment = -4131
    $titleRange.RowHeight = 32
    Release-ComObject $titleRange

    $metaRange = $tongAll.Range("A2:${lastColumn}7")
    $metaRange.Interior.Color = $lightGray
    $metaRange.Font.Color = $darkText
    Release-ComObject $metaRange

    foreach ($row in @($report.section_rows)) {
        $range = $tongAll.Range("A${row}:${lastColumn}${row}")
        $range.Interior.Color = $teal
        $range.Font.Color = $white
        $range.Font.Bold = $true
        $range.RowHeight = 24
        Release-ComObject $range
    }
    foreach ($row in @($report.header_rows)) {
        $range = $tongAll.Range("A${row}:${lastColumn}${row}")
        $range.Interior.Color = $blue
        $range.Font.Color = $white
        $range.Font.Bold = $true
        $range.HorizontalAlignment = -4108
        $range.WrapText = $true
        $range.RowHeight = 34
        Release-ComObject $range
    }
    foreach ($row in @($report.total_rows)) {
        $range = $tongAll.Range("A${row}:${lastColumn}${row}")
        $range.Interior.Color = $lightBlue
        $range.Font.Bold = $true
        Release-ComObject $range
    }
    foreach ($address in @($report.status_cells)) {
        $cell = $tongAll.Range([string]$address)
        if ([string]$cell.Value2 -eq [string]$report.ok_status) {
            $cell.Interior.Color = $green
            $cell.Font.Color = Get-ExcelColor '#166534'
        }
        else {
            $cell.Interior.Color = $red
            $cell.Font.Color = Get-ExcelColor '#B91C1C'
        }
        $cell.Font.Bold = $true
        $cell.HorizontalAlignment = -4108
        Release-ComObject $cell
    }

    $numberRange = $tongAll.Range(
        $tongAll.Cells.Item(1, 2),
        $tongAll.Cells.Item($rowCount, $columnCount)
    )
    $numberRange.NumberFormat = '#,##0'
    Release-ComObject $numberRange

    foreach ($address in @($report.center_ranges)) {
        $centerRange = $tongAll.Range([string]$address)
        $centerRange.HorizontalAlignment = -4108
        Release-ComObject $centerRange
    }

    for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
        $tongAll.Columns.Item($columnIndex).ColumnWidth = [double]$report.column_widths[$columnIndex - 1]
    }

    foreach ($address in @($report.filter_ranges)) {
        $filterRange = $tongAll.Range([string]$address)
        [void]$filterRange.AutoFilter()
        Release-ComObject $filterRange
    }

    $tongAll.Activate()
    $activeWindow = $excel.ActiveWindow
    if ($null -ne $activeWindow) {
        $activeWindow.DisplayGridlines = $false
        $activeWindow.SplitRow = 5
        $activeWindow.FreezePanes = $true
    }
    $tongAll.Range('A1').Select()

    $workbook.Save()
    Write-Output 'OK'
}
finally {
    Release-ComObject $statusRange
    Release-ComObject $targetRange
    Release-ComObject $activeWindow
    Release-ComObject $keHoach
    Release-ComObject $tongAll
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    Release-ComObject $workbook
    if ($null -ne $excel) {
        if ($null -ne $oldCalculation) {
            try { $excel.Calculation = $oldCalculation } catch {}
        }
        try { $excel.Quit() } catch {}
    }
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
