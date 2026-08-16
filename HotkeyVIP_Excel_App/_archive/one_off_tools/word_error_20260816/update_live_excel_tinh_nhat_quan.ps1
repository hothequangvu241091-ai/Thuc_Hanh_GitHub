$ErrorActionPreference = 'Stop'
$targetPath = 'D:\CodexProjects\Hotkeyvip\04_excel\hotkeyvip_test.xlsm'
$targetKeyword = 'tính nhất quán phương pháp luận'
$wordPath = 'D:\CodexProjects\Hotkeyvip\07_ket_qua\bai_viet\bantinkhoahoc.com\tính nhất quán phương pháp luận.docx'

$excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$book = $null
$sheet = $null
try {
    foreach ($candidate in $excel.Workbooks) {
        if ([string]::Equals($candidate.FullName, $targetPath, [StringComparison]::OrdinalIgnoreCase)) {
            $book = $candidate
            break
        }
    }
    if ($null -eq $book) {
        throw "Workbook đang mở không đúng file đích."
    }
    $sheet = $book.Worksheets.Item('VIET_BAI')
    $lastCol = $sheet.UsedRange.Columns.Count
    $headers = @{}
    for ($col = 1; $col -le $lastCol; $col++) {
        $name = [string]$sheet.Cells.Item(1, $col).Value2
        if ($name) { $headers[$name.Trim()] = $col }
    }
    foreach ($required in @('Từ khóa','Đường dẫn Word','Trạng thái viết','Lỗi viết','Số từ Word')) {
        if (-not $headers.ContainsKey($required)) {
            throw "Thiếu cột: $required"
        }
    }
    $lastRow = $sheet.UsedRange.Rows.Count
    $targetRow = $null
    for ($row = 2; $row -le $lastRow; $row++) {
        $keyword = [string]$sheet.Cells.Item($row, $headers['Từ khóa']).Value2
        if ($keyword.Trim() -eq $targetKeyword) {
            $targetRow = $row
            break
        }
    }
    if ($null -eq $targetRow) {
        throw "Không tìm thấy từ khóa đích."
    }
    $sheet.Cells.Item($targetRow, $headers['Đường dẫn Word']).Value2 = $wordPath
    $sheet.Cells.Item($targetRow, $headers['Trạng thái viết']).Value2 = 'OK'
    $sheet.Cells.Item($targetRow, $headers['Lỗi viết']).Value2 = ''
    $sheet.Cells.Item($targetRow, $headers['Số từ Word']).Value2 = 3222
    $book.Save()
    Write-Output "Đã cập nhật trực tiếp workbook đang mở: dòng $targetRow = OK"
    Write-Output ("Kiểm tra tại chỗ: " + [string]$sheet.Cells.Item($targetRow, $headers['Trạng thái viết']).Value2)
}
finally {
    if ($null -ne $sheet) { [Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null }
    if ($null -ne $book) { [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null }
    if ($null -ne $excel) { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
