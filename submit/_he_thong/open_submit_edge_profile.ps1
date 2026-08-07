# File hệ thống: mở đúng profile submit và lưu ánh xạ profile với Gmail.
param(
    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 14)]
    [int]$Number,

    [string]$Url = "https://myaccount.google.com/",

    [string]$InspectionUrl = "",

    [switch]$RequestIndexing
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

$profileName = "submit_{0:D2}" -f $Number
$remoteDebugPort = 9300 + $Number
$profilePath = Join-Path `
    "D:\CodexProjects\Hotkeyvip\06_du_lieu_chay\submit_edge_profiles" `
    $profileName
New-Item -ItemType Directory -Path $profilePath -Force | Out-Null

Start-Process `
    -FilePath $edgePath `
    -ArgumentList @(
        "--user-data-dir=$profilePath",
        "--profile-directory=Default",
        "--remote-debugging-address=127.0.0.1",
        "--remote-debugging-port=$remoteDebugPort",
        "--remote-allow-origins=*",
        "--no-first-run",
        "--no-default-browser-check",
        "--start-maximized",
        "--new-window",
        $Url
    )

if ($InspectionUrl) {
    $domHelper = Join-Path $PSScriptRoot "cdp_fill_search_console_url.py"
    $domArguments = @(
        "--port", $remoteDebugPort,
        "--text", $InspectionUrl,
        "--timeout", 35
    )
    if ($RequestIndexing) {
        $domArguments += @(
            "--click-request-indexing",
            "--request-timeout", 45
        )
    }
    & python $domHelper @domArguments
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
