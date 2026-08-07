# File hệ thống: thiết lập tuần tự các profile Edge chuyên dùng để submit URL.
param(
    [ValidateRange(1, 14)]
    [int]$StartFrom = 1
)

$ErrorActionPreference = "Stop"

$edgeCandidates = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
    "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"
)
$edgePath = $edgeCandidates |
    Where-Object { Test-Path -LiteralPath $_ } |
    Select-Object -First 1

if (-not $edgePath) {
    throw "Không tìm thấy Microsoft Edge."
}

$profilesRoot = "D:\CodexProjects\Hotkeyvip\06_du_lieu_chay\submit_edge_profiles"
New-Item -ItemType Directory -Path $profilesRoot -Force | Out-Null

Add-Type -AssemblyName System.Windows.Forms

for ($number = $StartFrom; $number -le 14; $number++) {
    $profileName = "submit_{0:D2}" -f $number
    $profilePath = Join-Path $profilesRoot $profileName
    New-Item -ItemType Directory -Path $profilePath -Force | Out-Null

    $message = @"
Chuẩn bị đăng nhập Gmail cho $profileName.

1. Đăng nhập đúng một tài khoản Gmail trong cửa sổ Edge sắp mở.
2. Không cần bật đồng bộ Edge.
3. Đăng nhập xong thì đóng toàn bộ cửa sổ Edge của profile này.
4. Profile kế tiếp sẽ tự mở.
"@
    [System.Windows.Forms.MessageBox]::Show(
        $message,
        "Thiết lập $profileName",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    $arguments = @(
        "--user-data-dir=$profilePath",
        "--profile-directory=Default",
        "--no-first-run",
        "--no-default-browser-check",
        "--start-maximized",
        "--new-window",
        "https://accounts.google.com/"
    )
    $edgeProcess = Start-Process `
        -FilePath $edgePath `
        -ArgumentList $arguments `
        -PassThru

    [System.Windows.Forms.MessageBox]::Show(
        "Sau khi đăng nhập Gmail và đóng cửa sổ Edge của $profileName, bấm OK để mở profile kế tiếp.",
        "Hoàn tất $profileName",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
    Start-Sleep -Milliseconds 800
}

[System.Windows.Forms.MessageBox]::Show(
    "Đã hoàn thành vòng đăng nhập 14 profile submit.",
    "Thiết lập hoàn tất",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null
