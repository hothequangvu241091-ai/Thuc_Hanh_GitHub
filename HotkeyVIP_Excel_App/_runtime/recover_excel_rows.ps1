param([string]$TargetFile)

$WorkbookPath = $TargetFile

$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
$sheet = $null
$backup = $null
try {
    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
        throw "Workbook not found: $WorkbookPath"
    }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $dir = [IO.Path]::GetDirectoryName($WorkbookPath)
    $base = [IO.Path]::GetFileNameWithoutExtension($WorkbookPath)
    $ext = [IO.Path]::GetExtension($WorkbookPath)
    $backup = Join-Path $dir ($base + '.before_recovery_' + $stamp + $ext)
    Copy-Item -LiteralPath $WorkbookPath -Destination $backup

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($WorkbookPath, 0, $false)
    if ($book.ReadOnly) { throw 'Workbook opened read-only.' }
    $sheet = $book.Worksheets.Item('VIET_BAI')

    $porterWord = [string]$sheet.Range('G4520').Value2
    $environmentWord = [string]$sheet.Range('G4524').Value2
    $porterStem = [IO.Path]::Combine([IO.Path]::GetDirectoryName($porterWord), [IO.Path]::GetFileNameWithoutExtension($porterWord))
    $environmentStem = [IO.Path]::Combine([IO.Path]::GetDirectoryName($environmentWord), [IO.Path]::GetFileNameWithoutExtension($environmentWord))
    $porter1 = $porterStem + ' 1.png'
    $porter2 = $porterStem + ' 2.png'
    $environment1 = $environmentStem + ' 1.png'
    $environment2 = $environmentStem + ' 2.png'
    foreach ($path in @($porterWord, $environmentWord, $porter1, $porter2, $environment1, $environment2)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing file: $path" }
    }

    $statusImg1 = [string]$sheet.Range('T4511').Value2
    $statusImg2 = [string]$sheet.Range('V4511').Value2
    if (-not $statusImg1 -or -not $statusImg2) { throw 'Reference image statuses are blank.' }

    # Row 4520: Word and images are proven; briefs are still missing, so X stays open.
    $sheet.Range('I4520').Value2 = 'OK'
    $sheet.Range('T4520').Value2 = $statusImg1
    $sheet.Range('U4520').Value2 = $porter1
    $sheet.Range('V4520').Value2 = $statusImg2
    $sheet.Range('W4520').Value2 = $porter2

    # Row 4524: Word, briefs and both image files are proven complete.
    $sheet.Range('T4524').Value2 = $statusImg1
    $sheet.Range('U4524').Value2 = $environment1
    $sheet.Range('V4524').Value2 = $statusImg2
    $sheet.Range('W4524').Value2 = $environment2
    $sheet.Range('X4524').Value2 = 'OK'

    $book.Save()
    $book.Close($false)
    [Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
    $sheet = $null
    [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
    $book = $null
    $excel.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    $excel = $null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($WorkbookPath, 0, $true)
    $sheet = $book.Worksheets.Item('VIET_BAI')
    $checks = [ordered]@{
        I4520 = [string]$sheet.Range('I4520').Value2
        U4520 = [string]$sheet.Range('U4520').Value2
        W4520 = [string]$sheet.Range('W4520').Value2
        X4520 = [string]$sheet.Range('X4520').Value2
        U4524 = [string]$sheet.Range('U4524').Value2
        W4524 = [string]$sheet.Range('W4524').Value2
        X4524 = [string]$sheet.Range('X4524').Value2
    }
    $checks | ConvertTo-Json -Compress
    Write-Output ('BACKUP=' + $backup)
}
finally {
    if ($book) { try { $book.Close($false) } catch {} }
    if ($sheet) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null } catch {} }
    if ($book) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null } catch {} }
    if ($excel) { try { $excel.Quit() } catch {}; try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
