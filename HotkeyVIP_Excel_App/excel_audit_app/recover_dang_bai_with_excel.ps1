param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$PayloadPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

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
$dangBai = $null
$targetRange = $null
$oldCalculation = $null

try {
    $config = $payload.dang_bai
    $rowCount = [int]$config.row_count
    $columnCount = [int]$config.column_count
    if ($rowCount -le 0 -or $columnCount -le 0) {
        throw 'Recovery payload is empty.'
    }
    if (@($config.rows).Count -ne $rowCount) {
        throw 'Recovery row count does not match the payload.'
    }

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

    $dangBai = Get-WorksheetCaseInsensitive $workbook ([string]$config.sheet_name)
    if ($null -eq $dangBai) {
        throw 'Cannot find DANG_BAI in the working copy.'
    }

    $headerRow = [int]$config.header_row
    $lastDataRow = [int]$config.last_data_row
    $startRow = [Math]::Max($headerRow + 1, $lastDataRow + 1)
    $endRow = $startRow + $rowCount - 1
    $targetRange = $dangBai.Range(
        $dangBai.Cells.Item($startRow, 1),
        $dangBai.Cells.Item($endRow, $columnCount)
    )
    if ([double]$excel.WorksheetFunction.CountA($targetRange) -ne 0) {
        throw 'Target rows in DANG_BAI are no longer empty. Analyze the source file again.'
    }

    $matrix = New-Object 'object[,]' $rowCount, $columnCount
    for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
        $sourceRow = @($config.rows[$rowIndex])
        if ($sourceRow.Count -ne $columnCount) {
            throw "Recovery row $($rowIndex + 1) has an invalid column count."
        }
        for ($columnIndex = 0; $columnIndex -lt $columnCount; $columnIndex++) {
            $matrix[$rowIndex, $columnIndex] = $sourceRow[$columnIndex]
        }
    }
    $targetRange.Value2 = $matrix

    $workbook.Save()
    Write-Output "OK|$startRow|$rowCount"
}
finally {
    Release-ComObject $targetRange
    Release-ComObject $dangBai
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
