# Quản lý Gmail, tên miền của Gmail và profile được gán.
param([switch]$ValidateOnly)

$ErrorActionPreference = "Stop"
$configPath = Join-Path $PSScriptRoot "submit_edge_profiles.json"
$urlStorePath = Join-Path $PSScriptRoot "submit_url_history.json"
$openProfileScript = Join-Path $PSScriptRoot "open_submit_edge_profile.ps1"
$restartLauncher = Join-Path $PSScriptRoot "restart_submit_manager.vbs"
$autoSubmitScript = Join-Path $PSScriptRoot "auto_submit_queue.py"
$autoProgressPath = Join-Path $PSScriptRoot "auto_submit_progress.json"
$autoStopFlagPath = Join-Path $PSScriptRoot "stop_auto_submit.flag"

function Read-Config {
    $data = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 |
        ConvertFrom-Json
    if (@($data.profiles).Count -ne 14) {
        throw "Cấu hình phải có đúng 14 profile."
    }
    return $data
}

function Initialize-Accounts {
    param($Config)

    if ($Config.PSObject.Properties.Name -contains "accounts") {
        return
    }

    $accounts = @()
    foreach ($group in @($Config.pendingAccountGroups)) {
        foreach ($email in @($group.emails)) {
            $accounts += [PSCustomObject]@{
                email = ([string]$email).Trim()
                domains = @($group.domains)
            }
        }
    }

    $Config | Add-Member -NotePropertyName "accounts" `
        -NotePropertyValue @($accounts | Sort-Object email)
}

function Get-Account {
    param($Config, [string]$Email)

    return $Config.accounts |
        Where-Object {
            [string]::Equals(
                ([string]$_.email).Trim(),
                $Email.Trim(),
                [System.StringComparison]::OrdinalIgnoreCase
            )
        } |
        Select-Object -First 1
}

function Get-DomainFromInput {
    param([string]$Value)

    $text = $Value.Trim().ToLowerInvariant()
    if (-not $text) {
        return ""
    }

    if ($text -match "^[a-z][a-z0-9+.-]*://") {
        try {
            $text = ([Uri]$text).Host.Trim(".").ToLowerInvariant()
        } catch {
            return ""
        }
    } else {
        $text = (($text -split "[/?#]", 2)[0]).Trim(".")
    }

    if ($text.StartsWith("www.")) {
        $text = $text.Substring(4)
    }
    return $text
}

function Save-Config {
    param($Config)

    Copy-Item -LiteralPath $configPath -Destination "$configPath.bak" -Force
    $json = $Config | ConvertTo-Json -Depth 8
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($configPath, $json, $utf8)
}

function Read-UrlEntries {
    if (-not (Test-Path -LiteralPath $urlStorePath)) {
        return @()
    }

    $data = Get-Content -LiteralPath $urlStorePath -Raw -Encoding UTF8 |
        ConvertFrom-Json
    if (-not ($data.PSObject.Properties.Name -contains "urls")) {
        return @()
    }
    return @($data.urls)
}

function Save-UrlEntries {
    param([array]$Entries)

    if (Test-Path -LiteralPath $urlStorePath) {
        Copy-Item -LiteralPath $urlStorePath -Destination "$urlStorePath.bak" -Force
    }

    $data = [ordered]@{
        version = 1
        urls = @($Entries)
    }
    $json = $data | ConvertTo-Json -Depth 6
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($urlStorePath, $json, $utf8)
}

function Get-CleanUrl {
    param([string]$Value)

    $text = $Value.Trim()
    if (-not $text) {
        return ""
    }

    try {
        $uri = [Uri]$text
        if (($uri.Scheme -ne "http") -and ($uri.Scheme -ne "https")) {
            return ""
        }
        if (-not $uri.Host) {
            return ""
        }
        return $text
    } catch {
        return ""
    }
}

function Get-UrlStatusLabel {
    param([string]$Code)

    switch ($Code.Trim().ToUpperInvariant()) {
        "PENDING" { return "CHƯA SUBMIT" }
        "PRIORITY" { return "ƯU TIÊN CHẠY TRƯỚC" }
        "RUNNING" { return "ĐANG SUBMIT" }
        "SUBMITTED" { return "HOÀN THÀNH" }
        "ERROR" { return "LỖI" }
        "QUOTA" { return "VƯỢT HẠN NGẠCH" }
        "SKIPPED" { return "BỎ QUA" }
        "UNMAPPED" { return "KHÔNG TÌM THẤY PROFILE" }
        default { return "CHƯA SUBMIT" }
    }
}

function Get-UrlStatusCode {
    param([string]$Label)

    switch ($Label.Trim().ToUpperInvariant()) {
        "CHƯA SUBMIT" { return "PENDING" }
        "ƯU TIÊN CHẠY TRƯỚC" { return "PRIORITY" }
        "ĐANG SUBMIT" { return "RUNNING" }
        "HOÀN THÀNH" { return "SUBMITTED" }
        "LỖI" { return "ERROR" }
        "VƯỢT HẠN NGẠCH" { return "QUOTA" }
        "BỎ QUA" { return "SKIPPED" }
        "KHÔNG TÌM THẤY PROFILE" { return "UNMAPPED" }
        default { return "PENDING" }
    }
}

$script:config = Read-Config
Initialize-Accounts -Config $script:config
$script:urlEntries = @(Read-UrlEntries)

if ($ValidateOnly) {
    Write-Output ("OK: {0} Gmail, {1} profiles." -f
        @($script:config.accounts).Count,
        @($script:config.profiles).Count)
    exit 0
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:loading = $false
$script:originalEmail = ""
$script:gmailChoiceMap = @()
$script:searchChoiceMap = @()
$script:pendingSearchAction = "submit"
$script:searchOnly = $false
$script:urlFilter = "TẤT CẢ"
$script:autoProcess = $null
$script:lastAutoProgress = ""
$script:lastLogPath = ""
$script:autoStartedAt = $null
$script:autoPulseFrame = 0
$script:mainUrlChoiceMap = @()
$script:selectedMainUrl = ""
$script:mainListOrderMode = "activity"
$navy = [System.Drawing.Color]::FromArgb(28, 48, 78)
$blue = [System.Drawing.Color]::FromArgb(33, 105, 218)
$red = [System.Drawing.Color]::FromArgb(190, 55, 55)
$green = [System.Drawing.Color]::FromArgb(22, 135, 78)
$amber = [System.Drawing.Color]::FromArgb(184, 112, 0)
$light = [System.Drawing.Color]::FromArgb(245, 247, 250)
$muted = [System.Drawing.Color]::FromArgb(83, 96, 112)

$form = New-Object System.Windows.Forms.Form
$form.Text = "Gmail - Tên miền - Edge Profile"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(900, 860)
$form.MinimumSize = New-Object System.Drawing.Size(820, 760)
$form.BackColor = $light
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$header = New-Object System.Windows.Forms.Panel
$header.Dock = "Top"
$header.Height = 76
$header.BackColor = $navy
$form.Controls.Add($header)

$title = New-Object System.Windows.Forms.Label
$title.Text = "HỆ THỐNG SUBMIT URL"
$title.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 18)
$title.ForeColor = [System.Drawing.Color]::White
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(22, 10)
$header.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "Submit URL và thiết lập Gmail/Profile được tách thành hai khu vực riêng."
$subtitle.ForeColor = [System.Drawing.Color]::FromArgb(205, 218, 238)
$subtitle.AutoSize = $true
$subtitle.Location = New-Object System.Drawing.Point(25, 45)
$header.Controls.Add($subtitle)

$openProjectFolderButton = New-Object System.Windows.Forms.Button
$openProjectFolderButton.Text = "MỞ THƯ MỤC DỰ ÁN"
$openProjectFolderButton.Size = New-Object System.Drawing.Size(170, 32)
$openProjectFolderButton.Anchor = "Top,Right"
$openProjectFolderButton.Location = New-Object System.Drawing.Point(704, 22)
$openProjectFolderButton.BackColor = [System.Drawing.Color]::White
$openProjectFolderButton.ForeColor = $navy
$openProjectFolderButton.FlatStyle = "Flat"
$openProjectFolderButton.FlatAppearance.BorderSize = 0
$openProjectFolderButton.Cursor = "Hand"
$header.Controls.Add($openProjectFolderButton)

$content = New-Object System.Windows.Forms.Panel
$content.Dock = "Fill"
$content.Padding = New-Object System.Windows.Forms.Padding(24)
$content.BackColor = $light
$form.Controls.Add($content)
$content.BringToFront()

$urlListForm = New-Object System.Windows.Forms.Form
$urlListForm.Text = "Danh sách URL submit"
$urlListForm.StartPosition = "CenterScreen"
$urlListForm.Size = New-Object System.Drawing.Size(980, 680)
$urlListForm.MinimumSize = New-Object System.Drawing.Size(760, 520)
$urlListForm.BackColor = $light
$urlListForm.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$urlImportForm = New-Object System.Windows.Forms.Form
$urlImportForm.Text = "Nạp thêm URL"
$urlImportForm.StartPosition = "CenterParent"
$urlImportForm.Size = New-Object System.Drawing.Size(760, 520)
$urlImportForm.MinimumSize = New-Object System.Drawing.Size(620, 420)
$urlImportForm.BackColor = $light
$urlImportForm.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$urlImportForm.Padding = New-Object System.Windows.Forms.Padding(18)

$urlImportLabel = New-Object System.Windows.Forms.Label
$urlImportLabel.Text = "DÁN URL MỚI — mỗi dòng một URL"
$urlImportLabel.Dock = "Top"
$urlImportLabel.Height = 34
$urlImportLabel.ForeColor = $navy
$urlImportLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$urlImportForm.Controls.Add($urlImportLabel)

$importUrlsButton = New-Object System.Windows.Forms.Button
$importUrlsButton.Text = "NẠP VÀO DANH SÁCH"
$importUrlsButton.Dock = "Bottom"
$importUrlsButton.Height = 42
$importUrlsButton.BackColor = $blue
$importUrlsButton.ForeColor = [System.Drawing.Color]::White
$importUrlsButton.FlatStyle = "Flat"
$importUrlsButton.FlatAppearance.BorderSize = 0
$importUrlsButton.Cursor = "Hand"
$urlImportForm.Controls.Add($importUrlsButton)

$urlImportBox = New-Object System.Windows.Forms.TextBox
$urlImportBox.Multiline = $true
$urlImportBox.AcceptsReturn = $true
$urlImportBox.WordWrap = $false
$urlImportBox.ScrollBars = "Both"
$urlImportBox.Dock = "Fill"
$urlImportBox.Font = New-Object System.Drawing.Font("Segoe UI", 10.5)
$urlImportForm.Controls.Add($urlImportBox)
$urlImportBox.BringToFront()

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "TÌM TÊN MIỀN HOẶC URL"
$searchLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$searchLabel.ForeColor = $navy
$searchLabel.AutoSize = $true
$searchLabel.Location = New-Object System.Drawing.Point(24, 18)
$content.Controls.Add($searchLabel)

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(24, 44)
$searchBox.Size = New-Object System.Drawing.Size(520, 31)
$searchBox.Anchor = "Top,Left,Right"
$content.Controls.Add($searchBox)

$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "SUBMIT"
$searchButton.Size = New-Object System.Drawing.Size(158, 32)
$searchButton.Anchor = "Top,Right"
$searchButton.Location = New-Object System.Drawing.Point(554, 42)
$searchButton.BackColor = $blue
$searchButton.ForeColor = [System.Drawing.Color]::White
$searchButton.FlatStyle = "Flat"
$searchButton.FlatAppearance.BorderSize = 0
$searchButton.Cursor = "Hand"
$content.Controls.Add($searchButton)

$searchResult = New-Object System.Windows.Forms.TextBox
$searchResult.Text = "Nhập tên miền hoặc toàn bộ URL để tìm Gmail và profile."
$searchResult.ForeColor = $muted
$searchResult.ReadOnly = $true
$searchResult.Multiline = $true
$searchResult.ScrollBars = "Vertical"
$searchResult.BackColor = [System.Drawing.Color]::White
$searchResult.Location = New-Object System.Drawing.Point(24, 80)
$searchResult.Size = New-Object System.Drawing.Size(688, 80)
$searchResult.Anchor = "Top,Left,Right"
$searchResult.Visible = $false
$content.Controls.Add($searchResult)

$searchChoiceCombo = New-Object System.Windows.Forms.ComboBox
$searchChoiceCombo.DropDownStyle = "DropDownList"
$searchChoiceCombo.Location = New-Object System.Drawing.Point(24, 168)
$searchChoiceCombo.Size = New-Object System.Drawing.Size(520, 31)
$searchChoiceCombo.Anchor = "Top,Left,Right"
$content.Controls.Add($searchChoiceCombo)

$openGscButton = New-Object System.Windows.Forms.Button
$openGscButton.Text = "MỞ GSC"
$openGscButton.Size = New-Object System.Drawing.Size(158, 32)
$openGscButton.Anchor = "Top,Right"
$openGscButton.Location = New-Object System.Drawing.Point(554, 166)
$openGscButton.BackColor = [System.Drawing.Color]::White
$openGscButton.ForeColor = $blue
$openGscButton.FlatStyle = "Flat"
$openGscButton.FlatAppearance.BorderColor = $blue
$openGscButton.Cursor = "Hand"
$content.Controls.Add($openGscButton)

$gmailLabel = New-Object System.Windows.Forms.Label
$gmailLabel.Text = "GMAIL — các Gmail CHƯA GÁN được xếp trên cùng"
$gmailLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$gmailLabel.ForeColor = $navy
$gmailLabel.AutoSize = $true
$gmailLabel.Location = New-Object System.Drawing.Point(24, 212)
$content.Controls.Add($gmailLabel)

$gmailCombo = New-Object System.Windows.Forms.ComboBox
$gmailCombo.DropDownStyle = "DropDown"
$gmailCombo.AutoCompleteMode = "SuggestAppend"
$gmailCombo.AutoCompleteSource = "ListItems"
$gmailCombo.Location = New-Object System.Drawing.Point(24, 238)
$gmailCombo.Size = New-Object System.Drawing.Size(360, 31)
$gmailCombo.Anchor = "Top,Left,Right"
$content.Controls.Add($gmailCombo)

$gmailOkButton = New-Object System.Windows.Forms.Button
$gmailOkButton.Text = "OK"
$gmailOkButton.Size = New-Object System.Drawing.Size(60, 34)
$gmailOkButton.Anchor = "Top,Right"
$gmailOkButton.Location = New-Object System.Drawing.Point(394, 236)
$gmailOkButton.BackColor = $blue
$gmailOkButton.ForeColor = [System.Drawing.Color]::White
$gmailOkButton.FlatStyle = "Flat"
$gmailOkButton.FlatAppearance.BorderSize = 0
$gmailOkButton.Cursor = "Hand"
$content.Controls.Add($gmailOkButton)

$newButton = New-Object System.Windows.Forms.Button
$newButton.Text = "+ THÊM MỚI"
$newButton.Size = New-Object System.Drawing.Size(120, 34)
$newButton.Anchor = "Top,Right"
$newButton.Location = New-Object System.Drawing.Point(462, 236)
$newButton.FlatStyle = "Flat"
$newButton.BackColor = [System.Drawing.Color]::White
$newButton.ForeColor = $blue
$newButton.Cursor = "Hand"
$content.Controls.Add($newButton)

$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "XÓA GMAIL"
$deleteButton.Size = New-Object System.Drawing.Size(120, 34)
$deleteButton.Anchor = "Top,Right"
$deleteButton.Location = New-Object System.Drawing.Point(592, 236)
$deleteButton.FlatStyle = "Flat"
$deleteButton.BackColor = [System.Drawing.Color]::White
$deleteButton.ForeColor = $red
$deleteButton.FlatAppearance.BorderColor = $red
$deleteButton.Cursor = "Hand"
$content.Controls.Add($deleteButton)

$domainLookupLabel = New-Object System.Windows.Forms.Label
$domainLookupLabel.Text = "TRA CỨU TÊN MIỀN — kiểm tra đang nằm trong Gmail nào"
$domainLookupLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$domainLookupLabel.ForeColor = $navy
$domainLookupLabel.AutoSize = $true
$content.Controls.Add($domainLookupLabel)

$domainLookupBox = New-Object System.Windows.Forms.TextBox
$content.Controls.Add($domainLookupBox)

$domainLookupButton = New-Object System.Windows.Forms.Button
$domainLookupButton.Text = "KIỂM TRA"
$domainLookupButton.BackColor = [System.Drawing.Color]::White
$domainLookupButton.ForeColor = $blue
$domainLookupButton.FlatStyle = "Flat"
$domainLookupButton.FlatAppearance.BorderColor = $blue
$domainLookupButton.Cursor = "Hand"
$content.Controls.Add($domainLookupButton)

$domainLookupResult = New-Object System.Windows.Forms.TextBox
$domainLookupResult.ReadOnly = $true
$domainLookupResult.Multiline = $true
$domainLookupResult.ScrollBars = "Vertical"
$domainLookupResult.BackColor = [System.Drawing.Color]::White
$domainLookupResult.ForeColor = $muted
$content.Controls.Add($domainLookupResult)

$domainLabel = New-Object System.Windows.Forms.Label
$domainLabel.Text = "TÊN MIỀN CỦA GMAIL — mỗi dòng một tên miền"
$domainLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$domainLabel.ForeColor = $navy
$domainLabel.AutoSize = $true
$domainLabel.Location = New-Object System.Drawing.Point(24, 294)
$content.Controls.Add($domainLabel)

$domainBox = New-Object System.Windows.Forms.TextBox
$domainBox.Multiline = $true
$domainBox.AcceptsReturn = $true
$domainBox.ScrollBars = "Vertical"
$domainBox.Location = New-Object System.Drawing.Point(24, 322)
$domainBox.Size = New-Object System.Drawing.Size(330, 250)
$domainBox.Anchor = "Top,Bottom,Left"
$domainBox.Font = New-Object System.Drawing.Font("Segoe UI", 10.5)
$content.Controls.Add($domainBox)

$allUrlsLabel = New-Object System.Windows.Forms.Label
$allUrlsLabel.Text = "DANH SÁCH URL ĐÃ LƯU"
$allUrlsLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$allUrlsLabel.ForeColor = $navy
$allUrlsLabel.AutoSize = $true
$allUrlsLabel.Location = New-Object System.Drawing.Point(376, 294)
$content.Controls.Add($allUrlsLabel)

$resetOuterOrderButton = New-Object System.Windows.Forms.Button
$resetOuterOrderButton.Text = "XẾP THEO STT GỐC"
$resetOuterOrderButton.Size = New-Object System.Drawing.Size(160, 28)
$resetOuterOrderButton.BackColor = [System.Drawing.Color]::White
$resetOuterOrderButton.ForeColor = $navy
$resetOuterOrderButton.FlatStyle = "Flat"
$resetOuterOrderButton.FlatAppearance.BorderColor = $navy
$resetOuterOrderButton.Cursor = "Hand"
$content.Controls.Add($resetOuterOrderButton)

$allUrlsBox = New-Object System.Windows.Forms.TextBox
$allUrlsBox.Multiline = $true
$allUrlsBox.ReadOnly = $true
$allUrlsBox.WordWrap = $false
$allUrlsBox.ScrollBars = "Both"
$allUrlsBox.Location = New-Object System.Drawing.Point(376, 322)
$allUrlsBox.Size = New-Object System.Drawing.Size(336, 166)
$allUrlsBox.Anchor = "Top,Bottom,Left,Right"
$allUrlsBox.Font = New-Object System.Drawing.Font("Segoe UI", 10.5)
$allUrlsBox.Visible = $false
$content.Controls.Add($allUrlsBox)

$allUrlsList = New-Object System.Windows.Forms.ListBox
$allUrlsList.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$allUrlsList.HorizontalScrollbar = $true
$allUrlsList.IntegralHeight = $false
$allUrlsList.BackColor = [System.Drawing.Color]::White
$allUrlsList.DrawMode = "OwnerDrawFixed"
$allUrlsList.ItemHeight = 22
$content.Controls.Add($allUrlsList)

$allUrlsList.add_DrawItem({
    param($sender, $eventArgs)
    if ($eventArgs.Index -lt 0 -or $eventArgs.Index -ge $sender.Items.Count) {
        return
    }

    $eventArgs.DrawBackground()
    $entry = if ($eventArgs.Index -lt $script:mainUrlChoiceMap.Count) {
        $script:mainUrlChoiceMap[$eventArgs.Index]
    } else {
        $null
    }
    $status = if ($entry) { ([string]$entry.status).ToUpperInvariant() } else { "" }
    $color = switch ($status) {
        "PRIORITY" { $amber; break }
        "RUNNING" { $blue; break }
        "SUBMITTED" { $green; break }
        "ERROR" { $red; break }
        "QUOTA" { [System.Drawing.Color]::FromArgb(180, 75, 0); break }
        "SKIPPED" { $muted; break }
        "UNMAPPED" { [System.Drawing.Color]::FromArgb(160, 40, 40); break }
        default { $navy; break }
    }
    if (($eventArgs.State -band [System.Windows.Forms.DrawItemState]::Selected) -ne 0) {
        $color = [System.Drawing.Color]::White
    }
    [System.Windows.Forms.TextRenderer]::DrawText(
        $eventArgs.Graphics,
        [string]$sender.Items[$eventArgs.Index],
        $sender.Font,
        $eventArgs.Bounds,
        $color,
        [System.Windows.Forms.TextFormatFlags]::Left -bor
            [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
            [System.Windows.Forms.TextFormatFlags]::NoPrefix
    )
    $eventArgs.DrawFocusRectangle()
})

$quickSubmitPanel = New-Object System.Windows.Forms.Panel
$quickSubmitPanel.BackColor = [System.Drawing.Color]::White
$quickSubmitPanel.BorderStyle = "FixedSingle"
$content.Controls.Add($quickSubmitPanel)

$quickSubmitTitle = New-Object System.Windows.Forms.Label
$quickSubmitTitle.Text = "URL ĐÃ CHỌN"
$quickSubmitTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$quickSubmitTitle.ForeColor = $navy
$quickSubmitTitle.Location = New-Object System.Drawing.Point(14, 12)
$quickSubmitTitle.AutoSize = $true
$quickSubmitPanel.Controls.Add($quickSubmitTitle)

$quickSubmitUrl = New-Object System.Windows.Forms.Label
$quickSubmitUrl.Text = "Bấm một URL trong danh sách để chọn."
$quickSubmitUrl.ForeColor = $muted
$quickSubmitUrl.Location = New-Object System.Drawing.Point(14, 40)
$quickSubmitUrl.Size = New-Object System.Drawing.Size(210, 70)
$quickSubmitUrl.AutoEllipsis = $true
$quickSubmitPanel.Controls.Add($quickSubmitUrl)

$browserModeLabel = New-Object System.Windows.Forms.Label
$browserModeLabel.Text = "CHẾ ĐỘ CHẠY"
$browserModeLabel.ForeColor = $navy
$browserModeLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$browserModeLabel.Location = New-Object System.Drawing.Point(14, 116)
$browserModeLabel.AutoSize = $true
$quickSubmitPanel.Controls.Add($browserModeLabel)

$browserModeCombo = New-Object System.Windows.Forms.ComboBox
$browserModeCombo.DropDownStyle = "DropDownList"
$browserModeCombo.Items.Add("HIỆN TRÌNH DUYỆT") | Out-Null
$browserModeCombo.Items.Add("CHẠY ẨN (KHÔNG MỞ EDGE)") | Out-Null
$browserModeCombo.SelectedIndex = 0
$browserModeCombo.Location = New-Object System.Drawing.Point(14, 136)
$browserModeCombo.Size = New-Object System.Drawing.Size(210, 28)
$quickSubmitPanel.Controls.Add($browserModeCombo)

$submitSelectedButton = New-Object System.Windows.Forms.Button
$submitSelectedButton.Text = "SUBMIT URL NÀY"
$submitSelectedButton.Size = New-Object System.Drawing.Size(210, 36)
$submitSelectedButton.Location = New-Object System.Drawing.Point(14, 174)
$submitSelectedButton.BackColor = $blue
$submitSelectedButton.ForeColor = [System.Drawing.Color]::White
$submitSelectedButton.FlatStyle = "Flat"
$submitSelectedButton.FlatAppearance.BorderSize = 0
$submitSelectedButton.Enabled = $false
$quickSubmitPanel.Controls.Add($submitSelectedButton)

$priorityButton = New-Object System.Windows.Forms.Button
$priorityButton.Text = "ƯU TIÊN CHẠY TRƯỚC"
$priorityButton.Size = New-Object System.Drawing.Size(210, 36)
$priorityButton.Location = New-Object System.Drawing.Point(14, 218)
$priorityButton.BackColor = [System.Drawing.Color]::White
$priorityButton.ForeColor = $blue
$priorityButton.FlatStyle = "Flat"
$priorityButton.FlatAppearance.BorderColor = $blue
$priorityButton.Enabled = $false
$quickSubmitPanel.Controls.Add($priorityButton)

$copySelectedUrlButton = New-Object System.Windows.Forms.Button
$copySelectedUrlButton.Text = "COPY URL"
$copySelectedUrlButton.Size = New-Object System.Drawing.Size(210, 30)
$copySelectedUrlButton.Location = New-Object System.Drawing.Point(14, 262)
$copySelectedUrlButton.BackColor = [System.Drawing.Color]::White
$copySelectedUrlButton.ForeColor = $navy
$copySelectedUrlButton.FlatStyle = "Flat"
$copySelectedUrlButton.FlatAppearance.BorderColor = $navy
$copySelectedUrlButton.Enabled = $false
$quickSubmitPanel.Controls.Add($copySelectedUrlButton)

$urlListContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$copyUrlMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("COPY URL")
$urlListContextMenu.Items.Add($copyUrlMenuItem) | Out-Null
$allUrlsList.ContextMenuStrip = $urlListContextMenu

$autoProgressPanel = New-Object System.Windows.Forms.Panel
$autoProgressPanel.BackColor = [System.Drawing.Color]::White
$autoProgressPanel.BorderStyle = "FixedSingle"
$autoProgressPanel.Location = New-Object System.Drawing.Point(24, 474)
$autoProgressPanel.Size = New-Object System.Drawing.Size(812, 108)
$autoProgressPanel.Anchor = "Bottom,Left,Right"
$autoProgressPanel.Visible = $false
$content.Controls.Add($autoProgressPanel)

$autoStateLabel = New-Object System.Windows.Forms.Label
$autoStateLabel.Text = "● CHƯA CHẠY"
$autoStateLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10.5)
$autoStateLabel.ForeColor = $muted
$autoStateLabel.Location = New-Object System.Drawing.Point(14, 8)
$autoStateLabel.Size = New-Object System.Drawing.Size(300, 24)
$autoProgressPanel.Controls.Add($autoStateLabel)

$autoCountLabel = New-Object System.Windows.Forms.Label
$autoCountLabel.Text = "0 / 0 URL"
$autoCountLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$autoCountLabel.ForeColor = $navy
$autoCountLabel.Location = New-Object System.Drawing.Point(520, 8)
$autoCountLabel.Size = New-Object System.Drawing.Size(274, 24)
$autoCountLabel.Anchor = "Top,Right"
$autoCountLabel.TextAlign = "MiddleRight"
$autoProgressPanel.Controls.Add($autoCountLabel)

$autoProgressBar = New-Object System.Windows.Forms.ProgressBar
$autoProgressBar.Location = New-Object System.Drawing.Point(16, 35)
$autoProgressBar.Size = New-Object System.Drawing.Size(778, 15)
$autoProgressBar.Anchor = "Top,Left,Right"
$autoProgressBar.Style = "Continuous"
$autoProgressBar.Minimum = 0
$autoProgressBar.Maximum = 100
$autoProgressBar.Value = 0
$autoProgressPanel.Controls.Add($autoProgressBar)

$autoStepLabel = New-Object System.Windows.Forms.Label
$autoStepLabel.Text = "Bấm SUBMIT TỰ ĐỘNG để bắt đầu."
$autoStepLabel.ForeColor = $muted
$autoStepLabel.Location = New-Object System.Drawing.Point(14, 55)
$autoStepLabel.Size = New-Object System.Drawing.Size(780, 22)
$autoStepLabel.Anchor = "Top,Left,Right"
$autoStepLabel.AutoEllipsis = $true
$autoProgressPanel.Controls.Add($autoStepLabel)

$autoUrlLabel = New-Object System.Windows.Forms.Label
$autoUrlLabel.Text = "URL hiện tại: —"
$autoUrlLabel.ForeColor = $navy
$autoUrlLabel.Location = New-Object System.Drawing.Point(14, 79)
$autoUrlLabel.Size = New-Object System.Drawing.Size(780, 22)
$autoUrlLabel.Anchor = "Top,Left,Right"
$autoUrlLabel.AutoEllipsis = $true
$autoProgressPanel.Controls.Add($autoUrlLabel)

$autoSubmitButton = New-Object System.Windows.Forms.Button
$autoSubmitButton.Text = "SUBMIT TỰ ĐỘNG"
$autoSubmitButton.Size = New-Object System.Drawing.Size(160, 34)
$autoSubmitButton.Anchor = "Bottom,Left"
$autoSubmitButton.Location = New-Object System.Drawing.Point(376, 496)
$autoSubmitButton.BackColor = $blue
$autoSubmitButton.ForeColor = [System.Drawing.Color]::White
$autoSubmitButton.FlatStyle = "Flat"
$autoSubmitButton.FlatAppearance.BorderSize = 0
$autoSubmitButton.Cursor = "Hand"
$content.Controls.Add($autoSubmitButton)

$stopAutoButton = New-Object System.Windows.Forms.Button
$stopAutoButton.Text = "DỪNG AN TOÀN"
$stopAutoButton.Size = New-Object System.Drawing.Size(168, 34)
$stopAutoButton.Anchor = "Bottom,Right"
$stopAutoButton.Location = New-Object System.Drawing.Point(544, 496)
$stopAutoButton.BackColor = [System.Drawing.Color]::White
$stopAutoButton.ForeColor = $red
$stopAutoButton.FlatStyle = "Flat"
$stopAutoButton.FlatAppearance.BorderColor = $red
$stopAutoButton.Cursor = "Hand"
$stopAutoButton.Enabled = $false
$content.Controls.Add($stopAutoButton)

$stopNowButton = New-Object System.Windows.Forms.Button
$stopNowButton.Text = "DỪNG NGAY"
$stopNowButton.Size = New-Object System.Drawing.Size(118, 34)
$stopNowButton.Anchor = "Bottom,Left"
$stopNowButton.Location = New-Object System.Drawing.Point(372, 496)
$stopNowButton.BackColor = $red
$stopNowButton.ForeColor = [System.Drawing.Color]::White
$stopNowButton.FlatStyle = "Flat"
$stopNowButton.FlatAppearance.BorderSize = 0
$stopNowButton.Cursor = "Hand"
$stopNowButton.Enabled = $false
$content.Controls.Add($stopNowButton)

$addUrlsButton = New-Object System.Windows.Forms.Button
$addUrlsButton.Text = "NẠP THÊM URL"
$addUrlsButton.Size = New-Object System.Drawing.Size(160, 34)
$addUrlsButton.Anchor = "Bottom,Left"
$addUrlsButton.Location = New-Object System.Drawing.Point(376, 538)
$addUrlsButton.BackColor = $blue
$addUrlsButton.ForeColor = [System.Drawing.Color]::White
$addUrlsButton.FlatStyle = "Flat"
$addUrlsButton.FlatAppearance.BorderSize = 0
$addUrlsButton.Cursor = "Hand"
$content.Controls.Add($addUrlsButton)

$openUrlListButton = New-Object System.Windows.Forms.Button
$openUrlListButton.Text = "XEM / SỬA"
$openUrlListButton.Size = New-Object System.Drawing.Size(168, 34)
$openUrlListButton.Anchor = "Bottom,Right"
$openUrlListButton.Location = New-Object System.Drawing.Point(544, 538)
$openUrlListButton.BackColor = [System.Drawing.Color]::White
$openUrlListButton.ForeColor = $blue
$openUrlListButton.FlatStyle = "Flat"
$openUrlListButton.FlatAppearance.BorderColor = $blue
$openUrlListButton.Cursor = "Hand"
$content.Controls.Add($openUrlListButton)

$profileLabel = New-Object System.Windows.Forms.Label
$profileLabel.Text = "GÁN GMAIL NÀY VÀO PROFILE"
$profileLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$profileLabel.ForeColor = $navy
$profileLabel.AutoSize = $true
$profileLabel.Anchor = "Bottom,Left"
$profileLabel.Location = New-Object System.Drawing.Point(24, 593)
$content.Controls.Add($profileLabel)

$profileCombo = New-Object System.Windows.Forms.ComboBox
$profileCombo.DropDownStyle = "DropDownList"
$profileCombo.Anchor = "Bottom,Left"
$profileCombo.Location = New-Object System.Drawing.Point(24, 620)
$profileCombo.Size = New-Object System.Drawing.Size(250, 31)
$profileCombo.Items.Add("— CHƯA GÁN PROFILE —") | Out-Null
foreach ($profile in @($script:config.profiles | Sort-Object id)) {
    $profileCombo.Items.Add([string]$profile.id) | Out-Null
}
$profileCombo.SelectedIndex = 0
$content.Controls.Add($profileCombo)

$viewProfileLabel = New-Object System.Windows.Forms.Label
$viewProfileLabel.Text = "XEM NGƯỢC THEO PROFILE"
$viewProfileLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$viewProfileLabel.ForeColor = $navy
$viewProfileLabel.AutoSize = $true
$viewProfileLabel.Anchor = "Bottom,Left"
$viewProfileLabel.Location = New-Object System.Drawing.Point(300, 593)
$content.Controls.Add($viewProfileLabel)

$viewProfileCombo = New-Object System.Windows.Forms.ComboBox
$viewProfileCombo.DropDownStyle = "DropDownList"
$viewProfileCombo.Anchor = "Bottom,Left"
$viewProfileCombo.Location = New-Object System.Drawing.Point(300, 620)
$viewProfileCombo.Size = New-Object System.Drawing.Size(250, 31)
$viewProfileCombo.Items.Add("— CHỌN PROFILE ĐỂ XEM —") | Out-Null
foreach ($profile in @($script:config.profiles | Sort-Object id)) {
    $viewProfileCombo.Items.Add([string]$profile.id) | Out-Null
}
$viewProfileCombo.SelectedIndex = 0
$content.Controls.Add($viewProfileCombo)

$openProfileButton = New-Object System.Windows.Forms.Button
$openProfileButton.Text = "MỞ PROFILE NÀY"
$openProfileButton.Size = New-Object System.Drawing.Size(142, 34)
$openProfileButton.Anchor = "Bottom,Right"
$openProfileButton.Location = New-Object System.Drawing.Point(570, 578)
$openProfileButton.BackColor = [System.Drawing.Color]::White
$openProfileButton.ForeColor = $blue
$openProfileButton.FlatStyle = "Flat"
$openProfileButton.FlatAppearance.BorderColor = $blue
$openProfileButton.Cursor = "Hand"
$content.Controls.Add($openProfileButton)

$quickAssignButton = New-Object System.Windows.Forms.Button
$quickAssignButton.Text = "GÁN NHANH"
$quickAssignButton.Size = New-Object System.Drawing.Size(142, 34)
$quickAssignButton.Anchor = "Bottom,Right"
$quickAssignButton.Location = New-Object System.Drawing.Point(570, 618)
$quickAssignButton.BackColor = $blue
$quickAssignButton.ForeColor = [System.Drawing.Color]::White
$quickAssignButton.FlatStyle = "Flat"
$quickAssignButton.FlatAppearance.BorderSize = 0
$quickAssignButton.Cursor = "Hand"
$content.Controls.Add($quickAssignButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Chọn Gmail để xem tên miền và profile."
$statusLabel.ForeColor = $muted
$statusLabel.AutoSize = $true
$statusLabel.Anchor = "Bottom,Left"
$statusLabel.Location = New-Object System.Drawing.Point(24, 670)
$content.Controls.Add($statusLabel)

$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "LƯU THAY ĐỔI"
$saveButton.Size = New-Object System.Drawing.Size(170, 42)
$saveButton.Anchor = "Bottom,Right"
$saveButton.Location = New-Object System.Drawing.Point(542, 660)
$saveButton.BackColor = $blue
$saveButton.ForeColor = [System.Drawing.Color]::White
$saveButton.FlatStyle = "Flat"
$saveButton.FlatAppearance.BorderSize = 0
$saveButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$saveButton.Cursor = "Hand"
$content.Controls.Add($saveButton)

$resetButton = New-Object System.Windows.Forms.Button
$resetButton.Text = "KHỞI ĐỘNG LẠI APP"
$resetButton.Size = New-Object System.Drawing.Size(150, 42)
$resetButton.Anchor = "Bottom,Right"
$resetButton.Location = New-Object System.Drawing.Point(382, 660)
$resetButton.BackColor = [System.Drawing.Color]::White
$resetButton.ForeColor = $navy
$resetButton.FlatStyle = "Flat"
$resetButton.Cursor = "Hand"
$content.Controls.Add($resetButton)

# Tách rõ hai khu vực: vận hành submit và thiết lập Gmail/Profile.
$content.Padding = New-Object System.Windows.Forms.Padding(0)

$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = "Fill"
$mainLayout.ColumnCount = 1
$mainLayout.RowCount = 3
$mainLayout.ColumnStyles.Add(
    (New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    ))
)
$mainLayout.RowStyles.Add(
    (New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        48
    ))
)
$mainLayout.RowStyles.Add(
    (New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    ))
)
$mainLayout.RowStyles.Add(
    (New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        42
    ))
)
$content.Controls.Add($mainLayout)

$modeBar = New-Object System.Windows.Forms.FlowLayoutPanel
$modeBar.Dock = "Fill"
$modeBar.FlowDirection = "LeftToRight"
$modeBar.WrapContents = $false
$modeBar.Padding = New-Object System.Windows.Forms.Padding(18, 6, 0, 4)
$modeBar.BackColor = [System.Drawing.Color]::White
$mainLayout.Controls.Add($modeBar, 0, 0)

$submitModeButton = New-Object System.Windows.Forms.Button
$submitModeButton.Text = "SUBMIT URL"
$submitModeButton.Size = New-Object System.Drawing.Size(180, 34)
$submitModeButton.BackColor = $blue
$submitModeButton.ForeColor = [System.Drawing.Color]::White
$submitModeButton.FlatStyle = "Flat"
$submitModeButton.FlatAppearance.BorderSize = 0
$submitModeButton.Cursor = "Hand"
$modeBar.Controls.Add($submitModeButton)

$setupModeButton = New-Object System.Windows.Forms.Button
$setupModeButton.Text = "THIẾT LẬP GMAIL & PROFILE"
$setupModeButton.Size = New-Object System.Drawing.Size(260, 34)
$setupModeButton.BackColor = [System.Drawing.Color]::White
$setupModeButton.ForeColor = $navy
$setupModeButton.FlatStyle = "Flat"
$setupModeButton.FlatAppearance.BorderColor = $navy
$setupModeButton.Cursor = "Hand"
$modeBar.Controls.Add($setupModeButton)

$modeWorkspace = New-Object System.Windows.Forms.Panel
$modeWorkspace.Dock = "Fill"
$modeWorkspace.BackColor = $light
$modeWorkspace.Size = New-Object System.Drawing.Size(884, 650)
$mainLayout.Controls.Add($modeWorkspace, 0, 1)

$submitPanel = New-Object System.Windows.Forms.Panel
$submitPanel.Size = New-Object System.Drawing.Size(884, 650)
$submitPanel.BackColor = $light
$modeWorkspace.Controls.Add($submitPanel)

$setupPanel = New-Object System.Windows.Forms.Panel
$setupPanel.Size = New-Object System.Drawing.Size(884, 650)
$setupPanel.BackColor = $light
$modeWorkspace.Controls.Add($setupPanel)

$submitControls = @(
    $searchLabel,
    $searchBox,
    $searchButton,
    $searchChoiceCombo,
    $openGscButton,
    $allUrlsLabel,
    $resetOuterOrderButton,
    $allUrlsList,
    $quickSubmitPanel,
    $autoProgressPanel,
    $autoSubmitButton,
    $stopAutoButton,
    $stopNowButton,
    $addUrlsButton,
    $openUrlListButton
)
foreach ($control in $submitControls) {
    $control.Parent = $submitPanel
}

$setupControls = @(
    $gmailLabel,
    $gmailCombo,
    $gmailOkButton,
    $newButton,
    $deleteButton,
    $domainLookupLabel,
    $domainLookupBox,
    $domainLookupButton,
    $domainLookupResult,
    $domainLabel,
    $domainBox,
    $profileLabel,
    $profileCombo,
    $viewProfileLabel,
    $viewProfileCombo,
    $openProfileButton,
    $quickAssignButton,
    $saveButton,
    $resetButton
)
foreach ($control in $setupControls) {
    $control.Parent = $setupPanel
}

# Reset là thao tác chung cho cả hai tab, đặt trên thanh chuyển chế độ.
$resetButton.Parent = $modeBar
$resetButton.Visible = $true
$resetButton.Size = New-Object System.Drawing.Size(155, 34)
$resetButton.Anchor = "None"
$resetButton.Margin = New-Object System.Windows.Forms.Padding(26, 0, 0, 0)

$statusLabel.Parent = $mainLayout
$statusLabel.Dock = "Fill"
$statusLabel.AutoSize = $false
$statusLabel.Anchor = "None"
$statusLabel.Padding = New-Object System.Windows.Forms.Padding(22, 0, 12, 0)
$statusLabel.TextAlign = "MiddleLeft"
$statusLabel.BackColor = [System.Drawing.Color]::White
$mainLayout.SetCellPosition(
    $statusLabel,
    (New-Object System.Windows.Forms.TableLayoutPanelCellPosition(0, 2))
)

# Bố cục khu vực SUBMIT URL.
$searchLabel.Location = New-Object System.Drawing.Point(24, 18)
$searchBox.Location = New-Object System.Drawing.Point(24, 44)
$searchBox.Size = New-Object System.Drawing.Size(620, 31)
$searchButton.Location = New-Object System.Drawing.Point(668, 42)
$searchButton.Size = New-Object System.Drawing.Size(168, 32)

$searchChoiceCombo.Location = New-Object System.Drawing.Point(24, 82)
$searchChoiceCombo.Size = New-Object System.Drawing.Size(620, 31)
$openGscButton.Location = New-Object System.Drawing.Point(668, 80)
$openGscButton.Size = New-Object System.Drawing.Size(168, 32)

$allUrlsLabel.Location = New-Object System.Drawing.Point(24, 134)
$allUrlsLabel.Text = "DANH SÁCH URL ĐÃ LƯU — trạng thái hiện tại"
$allUrlsList.Location = New-Object System.Drawing.Point(24, 162)
$allUrlsList.Size = New-Object System.Drawing.Size(548, 300)
$allUrlsList.Anchor = "Top,Bottom,Left,Right"
$quickSubmitPanel.Location = New-Object System.Drawing.Point(590, 162)
$quickSubmitPanel.Size = New-Object System.Drawing.Size(246, 300)
$quickSubmitPanel.Anchor = "Top,Bottom,Right"

$autoProgressPanel.Location = New-Object System.Drawing.Point(24, 474)
$autoProgressPanel.Size = New-Object System.Drawing.Size(812, 108)

$autoSubmitButton.Location = New-Object System.Drawing.Point(24, 474)
$autoSubmitButton.Anchor = "Bottom,Left"
$stopAutoButton.Location = New-Object System.Drawing.Point(194, 474)
$stopAutoButton.Anchor = "Bottom,Left"
$stopNowButton.Location = New-Object System.Drawing.Point(372, 474)
$stopNowButton.Anchor = "Bottom,Left"
$addUrlsButton.Location = New-Object System.Drawing.Point(500, 474)
$addUrlsButton.Anchor = "Bottom,Right"
$openUrlListButton.Location = New-Object System.Drawing.Point(668, 474)
$openUrlListButton.Anchor = "Bottom,Right"

# Bố cục khu vực THIẾT LẬP GMAIL & PROFILE: một Gmail là một cụm.
$gmailLabel.Location = New-Object System.Drawing.Point(24, 18)
$gmailLabel.Text = "CHỌN GMAIL ĐỂ CHỈNH — tên miền và profile sẽ tự hiện bên dưới"
$gmailCombo.Location = New-Object System.Drawing.Point(24, 44)
$gmailCombo.Size = New-Object System.Drawing.Size(620, 31)
$gmailOkButton.Visible = $false
$newButton.Text = "+ THÊM GMAIL"
$newButton.Location = New-Object System.Drawing.Point(654, 42)
$deleteButton.Location = New-Object System.Drawing.Point(786, 42)

$domainLookupLabel.Location = New-Object System.Drawing.Point(24, 84)
$domainLookupBox.Location = New-Object System.Drawing.Point(24, 106)
$domainLookupBox.Size = New-Object System.Drawing.Size(500, 28)
$domainLookupBox.Anchor = "Top,Left,Right"
$domainLookupButton.Location = New-Object System.Drawing.Point(534, 104)
$domainLookupButton.Size = New-Object System.Drawing.Size(120, 32)
$domainLookupButton.Anchor = "Top,Right"
$domainLookupResult.Location = New-Object System.Drawing.Point(24, 144)
$domainLookupResult.Size = New-Object System.Drawing.Size(812, 96)
$domainLookupResult.Anchor = "Top,Left,Right"

$domainLabel.Location = New-Object System.Drawing.Point(24, 252)
$domainLabel.Text = "TÊN MIỀN ĐI KÈM GMAIL NÀY — mỗi dòng một tên miền"
$domainBox.Location = New-Object System.Drawing.Point(24, 280)
$domainBox.Size = New-Object System.Drawing.Size(812, 208)
$domainBox.Anchor = "Top,Bottom,Left,Right"

$profileLabel.Location = New-Object System.Drawing.Point(24, 518)
$profileLabel.Text = "PROFILE GẮN VỚI GMAIL NÀY"
$profileCombo.Location = New-Object System.Drawing.Point(24, 545)
$profileCombo.Size = New-Object System.Drawing.Size(400, 31)
$openProfileButton.Location = New-Object System.Drawing.Point(438, 543)
$openProfileButton.Size = New-Object System.Drawing.Size(160, 34)
$openProfileButton.Anchor = "Bottom,Left"
$viewProfileLabel.Visible = $false
$viewProfileCombo.Visible = $false
$quickAssignButton.Visible = $false

$saveButton.Location = New-Object System.Drawing.Point(610, 543)
$saveButton.Size = New-Object System.Drawing.Size(200, 40)
$saveButton.Anchor = "Bottom,Left"

$submitPanel.Dock = "Fill"
$setupPanel.Dock = "Fill"
$submitPanel.Visible = $true
$setupPanel.Visible = $false
$submitPanel.BringToFront()

$submitModeButton.add_Click({
    $submitPanel.Visible = $true
    $setupPanel.Visible = $false
    $submitPanel.BringToFront()
    $submitModeButton.BackColor = $blue
    $submitModeButton.ForeColor = [System.Drawing.Color]::White
    $setupModeButton.BackColor = [System.Drawing.Color]::White
    $setupModeButton.ForeColor = $navy
})

$setupModeButton.add_Click({
    $submitPanel.Visible = $false
    $setupPanel.Visible = $true
    $setupPanel.BringToFront()
    $setupModeButton.BackColor = $blue
    $setupModeButton.ForeColor = [System.Drawing.Color]::White
    $submitModeButton.BackColor = [System.Drawing.Color]::White
    $submitModeButton.ForeColor = $navy
})

$openProjectFolderButton.add_Click({
    try {
        Start-Process -FilePath "explorer.exe" -ArgumentList @(
            "`"$((Split-Path -Parent $PSScriptRoot))`""
        )
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Không thể mở thư mục dự án: $($_.Exception.Message)",
            "Lỗi",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

$urlListForm.Padding = New-Object System.Windows.Forms.Padding(16)

$urlLayout = New-Object System.Windows.Forms.TableLayoutPanel
$urlLayout.Dock = "Fill"
$urlLayout.ColumnCount = 1
$urlLayout.RowCount = 3
$urlLayout.ColumnStyles.Add(
    (New-Object System.Windows.Forms.ColumnStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    ))
)
$urlLayout.RowStyles.Add(
    (New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        48
    ))
)
$urlLayout.RowStyles.Add(
    (New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Percent,
        100
    ))
)
$urlLayout.RowStyles.Add(
    (New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Absolute,
        58
    ))
)
$urlListForm.Controls.Add($urlLayout)

$urlQueueTopBar = New-Object System.Windows.Forms.FlowLayoutPanel
$urlQueueTopBar.Dock = "Fill"
$urlQueueTopBar.FlowDirection = "LeftToRight"
$urlQueueTopBar.WrapContents = $false
$urlLayout.Controls.Add($urlQueueTopBar, 0, 0)

$urlQueueSummary = New-Object System.Windows.Forms.Label
$urlQueueSummary.Text = "Chưa có URL trong file."
$urlQueueSummary.ForeColor = $navy
$urlQueueSummary.Size = New-Object System.Drawing.Size(470, 42)
$urlQueueSummary.TextAlign = "MiddleLeft"
$urlQueueSummary.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$urlQueueTopBar.Controls.Add($urlQueueSummary)

$urlFilterLabel = New-Object System.Windows.Forms.Label
$urlFilterLabel.Text = "HIỂN THỊ:"
$urlFilterLabel.Size = New-Object System.Drawing.Size(78, 42)
$urlFilterLabel.TextAlign = "MiddleRight"
$urlFilterLabel.ForeColor = $navy
$urlQueueTopBar.Controls.Add($urlFilterLabel)

$urlFilterCombo = New-Object System.Windows.Forms.ComboBox
$urlFilterCombo.DropDownStyle = "DropDownList"
$urlFilterCombo.Size = New-Object System.Drawing.Size(220, 32)
$urlFilterCombo.Margin = New-Object System.Windows.Forms.Padding(3, 7, 3, 3)
foreach ($status in @(
    "TẤT CẢ",
    "ƯU TIÊN CHẠY TRƯỚC",
    "CHƯA SUBMIT",
    "ĐANG SUBMIT",
    "HOÀN THÀNH",
    "LỖI",
    "VƯỢT HẠN NGẠCH",
    "BỎ QUA",
    "KHÔNG TÌM THẤY PROFILE"
)) {
    $urlFilterCombo.Items.Add($status) | Out-Null
}
$urlFilterCombo.SelectedIndex = 0
$urlQueueTopBar.Controls.Add($urlFilterCombo)

$urlQueueGrid = New-Object System.Windows.Forms.DataGridView
$urlQueueGrid.Dock = "Fill"
$urlQueueGrid.BackgroundColor = [System.Drawing.Color]::White
$urlQueueGrid.BorderStyle = "FixedSingle"
$urlQueueGrid.AllowUserToAddRows = $false
$urlQueueGrid.AllowUserToDeleteRows = $false
$urlQueueGrid.AllowUserToOrderColumns = $false
$urlQueueGrid.RowHeadersVisible = $false
$urlQueueGrid.SelectionMode = "FullRowSelect"
$urlQueueGrid.MultiSelect = $true
$urlQueueGrid.AutoSizeRowsMode = "AllCells"
$urlQueueGrid.DefaultCellStyle.WrapMode = "False"
$urlQueueGrid.ColumnHeadersHeightSizeMode = "AutoSize"

$indexColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$indexColumn.Name = "Index"
$indexColumn.HeaderText = "STT"
$indexColumn.Width = 48
$indexColumn.ReadOnly = $true
$indexColumn.SortMode = "NotSortable"
$indexColumn.DefaultCellStyle.Alignment = "MiddleCenter"
$urlQueueGrid.Columns.Add($indexColumn) | Out-Null

$statusColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$statusColumn.Name = "Status"
$statusColumn.HeaderText = "TRẠNG THÁI"
$statusColumn.Width = 195
$statusColumn.SortMode = "Automatic"
foreach ($status in @(
    "ƯU TIÊN CHẠY TRƯỚC",
    "CHƯA SUBMIT",
    "ĐANG SUBMIT",
    "HOÀN THÀNH",
    "LỖI",
    "VƯỢT HẠN NGẠCH",
    "BỎ QUA",
    "KHÔNG TÌM THẤY PROFILE"
)) {
    $statusColumn.Items.Add($status) | Out-Null
}
$urlQueueGrid.Columns.Add($statusColumn) | Out-Null

$urlColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$urlColumn.Name = "Url"
$urlColumn.HeaderText = "URL"
$urlColumn.AutoSizeMode = "Fill"
$urlColumn.MinimumWidth = 280
$urlColumn.SortMode = "Automatic"
$urlQueueGrid.Columns.Add($urlColumn) | Out-Null

$messageColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$messageColumn.Name = "Message"
$messageColumn.HeaderText = "GHI CHÚ / LỖI"
$messageColumn.Width = 190
$urlQueueGrid.Columns.Add($messageColumn) | Out-Null

$updatedColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$updatedColumn.Name = "UpdatedAt"
$updatedColumn.HeaderText = "CẬP NHẬT"
$updatedColumn.Width = 145
$updatedColumn.ReadOnly = $true
$urlQueueGrid.Columns.Add($updatedColumn) | Out-Null

$urlQueueGrid.add_Sorted({
    for ($index = 0; $index -lt $urlQueueGrid.Rows.Count; $index++) {
        $urlQueueGrid.Rows[$index].Cells["Index"].Value = $index + 1
    }
})

$urlLayout.Controls.Add($urlQueueGrid, 0, 1)

$urlButtonBar = New-Object System.Windows.Forms.FlowLayoutPanel
$urlButtonBar.Dock = "Fill"
$urlButtonBar.FlowDirection = "LeftToRight"
$urlButtonBar.WrapContents = $false
$urlButtonBar.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)
$urlLayout.Controls.Add($urlButtonBar, 0, 2)

$saveUrlListButton = New-Object System.Windows.Forms.Button
$saveUrlListButton.Text = "LƯU DANH SÁCH"
$saveUrlListButton.Size = New-Object System.Drawing.Size(145, 36)
$saveUrlListButton.BackColor = $blue
$saveUrlListButton.ForeColor = [System.Drawing.Color]::White
$saveUrlListButton.FlatStyle = "Flat"
$saveUrlListButton.FlatAppearance.BorderSize = 0
$saveUrlListButton.Cursor = "Hand"
$urlButtonBar.Controls.Add($saveUrlListButton)

$bulkStatusCombo = New-Object System.Windows.Forms.ComboBox
$bulkStatusCombo.DropDownStyle = "DropDownList"
$bulkStatusCombo.Size = New-Object System.Drawing.Size(165, 36)
foreach ($status in @(
    "ƯU TIÊN CHẠY TRƯỚC",
    "CHƯA SUBMIT",
    "ĐANG SUBMIT",
    "HOÀN THÀNH",
    "LỖI",
    "VƯỢT HẠN NGẠCH",
    "BỎ QUA",
    "KHÔNG TÌM THẤY PROFILE"
)) {
    $bulkStatusCombo.Items.Add($status) | Out-Null
}
$bulkStatusCombo.SelectedIndex = 0
$urlButtonBar.Controls.Add($bulkStatusCombo)

$applyStatusButton = New-Object System.Windows.Forms.Button
$applyStatusButton.Text = "ÁP DỤNG TRẠNG THÁI"
$applyStatusButton.Size = New-Object System.Drawing.Size(180, 36)
$applyStatusButton.BackColor = [System.Drawing.Color]::White
$applyStatusButton.ForeColor = $blue
$applyStatusButton.FlatStyle = "Flat"
$applyStatusButton.FlatAppearance.BorderColor = $blue
$applyStatusButton.Cursor = "Hand"
$urlButtonBar.Controls.Add($applyStatusButton)

$deleteUrlButton = New-Object System.Windows.Forms.Button
$deleteUrlButton.Text = "XÓA DÒNG ĐÃ CHỌN"
$deleteUrlButton.Size = New-Object System.Drawing.Size(175, 36)
$deleteUrlButton.BackColor = [System.Drawing.Color]::White
$deleteUrlButton.ForeColor = $red
$deleteUrlButton.FlatStyle = "Flat"
$deleteUrlButton.FlatAppearance.BorderColor = $red
$deleteUrlButton.Cursor = "Hand"
$urlButtonBar.Controls.Add($deleteUrlButton)

$deleteExternalUrlButton = New-Object System.Windows.Forms.Button
$deleteExternalUrlButton.Text = "XÓA NHANH TOÀN BỘ URL"
$deleteExternalUrlButton.Size = New-Object System.Drawing.Size(190, 36)
$deleteExternalUrlButton.BackColor = [System.Drawing.Color]::White
$deleteExternalUrlButton.ForeColor = $red
$deleteExternalUrlButton.FlatStyle = "Flat"
$deleteExternalUrlButton.FlatAppearance.BorderColor = $red
$deleteExternalUrlButton.Cursor = "Hand"
$deleteExternalUrlButton.Parent = $submitPanel

$viewLogButton = New-Object System.Windows.Forms.Button
$viewLogButton.Text = "XEM LOG"
$viewLogButton.Size = New-Object System.Drawing.Size(110, 36)
$viewLogButton.BackColor = [System.Drawing.Color]::White
$viewLogButton.ForeColor = $navy
$viewLogButton.FlatStyle = "Flat"
$viewLogButton.Cursor = "Hand"
$urlButtonBar.Controls.Add($viewLogButton)

$autoProgressTimer = New-Object System.Windows.Forms.Timer
$autoProgressTimer.Interval = 1000

function Refresh-GmailList {
    param([string]$SelectEmail = "")

    $script:loading = $true
    $gmailCombo.Items.Clear()
    $script:gmailChoiceMap = @()

    $choices = foreach ($account in @($script:config.accounts)) {
        $email = ([string]$account.email).Trim()
        $assigned = $script:config.profiles |
            Where-Object {
                [string]::Equals(
                    ([string]$_.email).Trim(),
                    $email,
                    [System.StringComparison]::OrdinalIgnoreCase
                )
            } |
            Select-Object -First 1
        [PSCustomObject]@{
            Email = $email
            Profile = if ($assigned) { [string]$assigned.id } else { "" }
            Unassigned = if ($assigned) { 1 } else { 0 }
        }
    }

    $choices = @($choices | Sort-Object Unassigned, Profile, Email)
    $selectedIndex = -1
    for ($index = 0; $index -lt $choices.Count; $index++) {
        $choice = $choices[$index]
        $label = if ($choice.Profile) {
            "{0}  —  {1}" -f $choice.Email, $choice.Profile
        } else {
            "{0}  —  CHƯA GÁN" -f $choice.Email
        }
        $gmailCombo.Items.Add($label) | Out-Null
        $script:gmailChoiceMap += [string]$choice.Email
        if (
            $SelectEmail -and
            [string]::Equals(
                [string]$choice.Email,
                $SelectEmail,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            $selectedIndex = $index
        }
    }
    if ($selectedIndex -ge 0) {
        $gmailCombo.SelectedIndex = $selectedIndex
    } elseif ($SelectEmail) {
        $gmailCombo.Text = $SelectEmail
    }
    $script:loading = $false
}

function Get-GmailComboEmail {
    if (
        $gmailCombo.SelectedIndex -ge 0 -and
        $gmailCombo.SelectedIndex -lt $script:gmailChoiceMap.Count
    ) {
        return [string]$script:gmailChoiceMap[$gmailCombo.SelectedIndex]
    }

    $text = $gmailCombo.Text.Trim()
    $separatorIndex = $text.IndexOf(
        "  —  ",
        [System.StringComparison]::Ordinal
    )
    if ($separatorIndex -gt 0) {
        return $text.Substring(0, $separatorIndex).Trim()
    }
    return $text
}

function Select-GmailInCombo {
    param([string]$Email)

    $previousLoading = $script:loading
    $script:loading = $true
    try {
        $gmailCombo.SelectedIndex = -1
        for (
            $index = 0;
            $index -lt $script:gmailChoiceMap.Count;
            $index++
        ) {
            if ([string]::Equals(
                [string]$script:gmailChoiceMap[$index],
                $Email,
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
                $gmailCombo.SelectedIndex = $index
                return
            }
        }
        $gmailCombo.Text = $Email
    } finally {
        $script:loading = $previousLoading
    }
}

function Load-Gmail {
    param([string]$Email)

    $account = Get-Account -Config $script:config -Email $Email
    if (-not $account) {
        return
    }

    $script:originalEmail = [string]$account.email
    Select-GmailInCombo -Email ([string]$account.email)
    $domainBox.Text = @($account.domains) -join "`r`n"

    $assigned = $script:config.profiles |
        Where-Object {
            [string]::Equals(
                ([string]$_.email).Trim(),
                $script:originalEmail,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        } |
        Select-Object -First 1

    if ($assigned) {
        $profileCombo.SelectedItem = [string]$assigned.id
        $statusLabel.Text = "$($account.email) đang được gán vào $($assigned.id)."
    } else {
        $profileCombo.SelectedIndex = 0
        $statusLabel.Text = "$($account.email) hiện chưa gán vào profile nào."
    }
}

function Show-DomainAssignments {
    $domain = Get-DomainFromInput -Value $domainLookupBox.Text
    if (-not $domain) {
        $domainLookupResult.ForeColor = $red
        $domainLookupResult.Text = "Nhập tên miền hợp lệ để kiểm tra."
        return
    }

    # Khi người dùng dán cả URL, trả ô nhập về tên miền chuẩn đang được tra cứu.
    $domainLookupBox.Text = $domain

    $matches = @()
    foreach ($account in @($script:config.accounts)) {
        $hasDomain = @($account.domains | Where-Object {
            (Get-DomainFromInput -Value ([string]$_)) -eq $domain
        }).Count -gt 0
        if (-not $hasDomain) {
            continue
        }
        $email = ([string]$account.email).Trim()
        $profile = $script:config.profiles | Where-Object {
            [string]::Equals(
                ([string]$_.email).Trim(),
                $email,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        } | Select-Object -First 1
        $profileText = if ($profile) {
            [string]$profile.id
        } else {
            "CHƯA GÁN PROFILE"
        }
        $matches += "$email — $profileText"
    }

    if ($matches.Count -eq 0) {
        $domainLookupResult.ForeColor = $red
        $domainLookupResult.Text = "$domain — chưa nằm trong Gmail nào."
        return
    }

    $domainLookupResult.ForeColor = $green
    $domainLookupResult.Text = (
        "$domain — có trong $($matches.Count) Gmail:`r`n" +
        ($matches -join "`r`n")
    )
}

function Refresh-UrlGrid {
    $previousSelectedUrl = $script:selectedMainUrl
    $urlQueueGrid.Rows.Clear()
    $allUrlsList.Items.Clear()
    $script:mainUrlChoiceMap = @()
    $urlQueueSummary.ForeColor = $navy

    foreach ($entry in @($script:urlEntries)) {
        $status = ([string]$entry.status).Trim()
        if (-not $status) {
            $status = "PENDING"
        }
        $statusLabelText = Get-UrlStatusLabel -Code $status
        if (
            ($script:urlFilter -ne "TẤT CẢ") -and
            ($statusLabelText -ne $script:urlFilter)
        ) {
            continue
        }

        $rowIndex = $urlQueueGrid.Rows.Add(
            $urlQueueGrid.Rows.Count + 1,
            $statusLabelText,
            [string]$entry.url,
            [string]$entry.message,
            [string]$entry.updatedAt
        )
        $urlQueueGrid.Rows[$rowIndex].Tag = [string]$entry.url
    }

    $displayItems = @()
    for ($queueIndex = 0; $queueIndex -lt $script:urlEntries.Count; $queueIndex++) {
        $entry = $script:urlEntries[$queueIndex]
        $statusCode = ([string]$entry.status).Trim().ToUpperInvariant()
        if (-not $statusCode) { $statusCode = "PENDING" }
        $displayItems += [PSCustomObject]@{
            Entry = $entry
            QueueNumber = $queueIndex + 1
            StatusCode = $statusCode
            UpdatedAt = [string]$entry.updatedAt
        }
    }

    if ($script:mainListOrderMode -eq "activity") {
        $runningItems = @($displayItems | Where-Object {
            $_.StatusCode -eq "RUNNING"
        } | Sort-Object UpdatedAt -Descending)
        $recentItems = @($displayItems | Where-Object {
            $_.StatusCode -in @(
                "SUBMITTED", "ERROR", "QUOTA", "SKIPPED", "UNMAPPED"
            )
        } | Sort-Object UpdatedAt -Descending)
        $priorityItems = @($displayItems | Where-Object {
            $_.StatusCode -eq "PRIORITY"
        } | Sort-Object QueueNumber)
        $waitingItems = @($displayItems | Where-Object {
            $_.StatusCode -notin @(
                "RUNNING", "SUBMITTED", "ERROR", "QUOTA", "SKIPPED",
                "UNMAPPED", "PRIORITY"
            )
        } | Sort-Object QueueNumber)
        $displayItems = @(
            $runningItems + $recentItems + $priorityItems + $waitingItems
        )
    }

    foreach ($displayItem in @($displayItems)) {
        $entry = $displayItem.Entry
        $statusText = Get-UrlStatusLabel -Code $displayItem.StatusCode
        $queueNumber = $displayItem.QueueNumber
        $allUrlsList.Items.Add((
            "[{0}] [{1}] {2}" -f
            $queueNumber, $statusText, ([string]$entry.url)
        )) |
            Out-Null
        $script:mainUrlChoiceMap += $entry
    }
    $selectedIndex = -1
    for ($index = 0; $index -lt $script:mainUrlChoiceMap.Count; $index++) {
        if ([string]$script:mainUrlChoiceMap[$index].url -eq $previousSelectedUrl) {
            $selectedIndex = $index
            break
        }
    }
    if ($selectedIndex -ge 0) {
        $allUrlsList.SelectedIndex = $selectedIndex
    } elseif ($allUrlsList.Items.Count -gt 0) {
        $allUrlsList.SelectedIndex = 0
    }

    $pendingCount = @(
        $script:urlEntries |
            Where-Object {
                ([string]$_.status) -in @("PENDING", "PRIORITY")
            }
    ).Count
    $submittedCount = @(
        $script:urlEntries |
            Where-Object { ([string]$_.status) -eq "SUBMITTED" }
    ).Count
    $errorCount = @(
        $script:urlEntries |
            Where-Object { ([string]$_.status) -eq "ERROR" }
    ).Count
    $quotaCount = @(
        $script:urlEntries |
            Where-Object { ([string]$_.status) -eq "QUOTA" }
    ).Count

    $urlQueueSummary.Text = (
        "Tổng: {0}   |   Đang hiện: {1}   |   Chờ: {2}   |   Hoàn thành: {3}   |   Lỗi: {4}   |   Hạn ngạch: {5} (tự thử lại ngày mai)" -f
        @($script:urlEntries).Count,
        $urlQueueGrid.Rows.Count,
        $pendingCount,
        $submittedCount,
        $errorCount,
        $quotaCount
    )
}

function Highlight-MainUrl {
    param([string]$Url)

    if (-not $Url) {
        return
    }
    for ($index = 0; $index -lt $script:mainUrlChoiceMap.Count; $index++) {
        if ([string]$script:mainUrlChoiceMap[$index].url -eq $Url) {
            $script:selectedMainUrl = $Url
            $allUrlsList.SelectedIndex = $index
            $allUrlsList.TopIndex = [math]::Max(0, $index - 4)
            return
        }
    }
}

function Save-UrlGrid {
    $urlQueueGrid.EndEdit()
    $updates = @{}
    $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    foreach ($row in @($urlQueueGrid.Rows)) {
        if ($row.IsNewRow) {
            continue
        }

        $url = Get-CleanUrl -Value ([string]$row.Cells["Url"].Value)
        if (-not $url) {
            throw "Có URL không hợp lệ trong danh sách đang hiển thị."
        }

        $status = Get-UrlStatusCode -Label (
            [string]$row.Cells["Status"].Value
        )

        $originalUrl = ([string]$row.Tag).Trim()
        if (-not $originalUrl) {
            $originalUrl = $url
        }
        $updates[$originalUrl.ToLowerInvariant()] = [PSCustomObject]@{
            url = $url
            status = $status
            message = ([string]$row.Cells["Message"].Value).Trim()
            updatedAt = $now
        }
    }

    $entries = @()
    foreach ($entry in @($script:urlEntries)) {
        $originalKey = ([string]$entry.url).Trim().ToLowerInvariant()
        if ($updates.ContainsKey($originalKey)) {
            $entries += $updates[$originalKey]
            $updates.Remove($originalKey) | Out-Null
        } else {
            $entries += $entry
        }
    }
    foreach ($entry in @($updates.Values)) {
        $entries += $entry
    }

    $seen = @{}
    foreach ($entry in $entries) {
        $key = ([string]$entry.url).Trim().ToLowerInvariant()
        if ($seen.ContainsKey($key)) {
            throw "Danh sách có URL bị trùng: $($entry.url)"
        }
        $seen[$key] = $true
    }

    $script:urlEntries = @($entries)
    Save-UrlEntries -Entries $script:urlEntries
    Refresh-UrlGrid
}

function Update-SubmitLayout {
    [int]$panelWidth = $submitPanel.ClientSize.Width
    [int]$panelHeight = $submitPanel.ClientSize.Height
    if ($panelWidth -lt 400 -or $panelHeight -lt 300) {
        return
    }

    [int]$buttonY = $panelHeight - 46
    [int]$listTop = 162
    [int]$listBottom = $buttonY - 12

    if ($autoProgressPanel.Visible) {
        [int]$progressY = $buttonY - 120
        [int]$progressWidth = $panelWidth - 48
        if ($progressWidth -lt 300) { $progressWidth = 300 }
        $autoProgressPanel.Location = [System.Drawing.Point]::new(24, $progressY)
        $autoProgressPanel.Size = [System.Drawing.Size]::new($progressWidth, 108)
        $listBottom = $progressY - 12
    }

    [int]$listWidth = $panelWidth - 336
    if ($listWidth -lt 300) { $listWidth = 300 }
    [int]$listHeight = $listBottom - $listTop
    if ($listHeight -lt 180) { $listHeight = 180 }
    $allUrlsList.Location = [System.Drawing.Point]::new(24, $listTop)
    $allUrlsList.Size = [System.Drawing.Size]::new($listWidth, $listHeight)
    $allUrlsList.Anchor = "Top,Left,Right"

    [int]$orderButtonX = $listWidth - 136
    if ($orderButtonX -lt 330) { $orderButtonX = 330 }
    $resetOuterOrderButton.Location = [System.Drawing.Point]::new(
        $orderButtonX,
        130
    )

    [int]$quickPanelX = $listWidth + 42
    $quickSubmitPanel.Location = [System.Drawing.Point]::new($quickPanelX, $listTop)
    $quickSubmitPanel.Size = [System.Drawing.Size]::new(246, $listHeight)
    $quickSubmitPanel.Anchor = "Top,Right"

    $autoSubmitButton.Location = [System.Drawing.Point]::new(24, $buttonY)
    $stopAutoButton.Location = [System.Drawing.Point]::new(194, $buttonY)
    $stopNowButton.Location = [System.Drawing.Point]::new(372, $buttonY)
    $addUrlsButton.Location = [System.Drawing.Point]::new(500, $buttonY)
    $openUrlListButton.Location = [System.Drawing.Point]::new(668, $buttonY)
    $deleteExternalUrlButton.Location = [System.Drawing.Point]::new(668, $buttonY - 40)
    $deleteExternalUrlButton.Anchor = "Bottom,Right"
}

function Set-AutoProgressPanelVisible {
    param([bool]$Visible)

    $autoProgressPanel.Visible = $Visible
    Update-SubmitLayout
}

$submitPanel.add_Resize({ Update-SubmitLayout })

function Get-BrowserRunMode {
    if ($browserModeCombo.SelectedIndex -eq 1) {
        return "hidden"
    }
    return "visible"
}

function Set-AutoUiState {
    param([bool]$Running)

    $autoSubmitButton.Enabled = -not $Running
    $submitSelectedButton.Enabled = (-not $Running) -and
        ($allUrlsList.SelectedIndex -ge 0)
    $priorityButton.Enabled = (-not $Running) -and
        ($allUrlsList.SelectedIndex -ge 0)
    $browserModeCombo.Enabled = -not $Running
    $stopAutoButton.Enabled = $Running
    $stopNowButton.Enabled = $Running
    $addUrlsButton.Enabled = -not $Running
    $saveUrlListButton.Enabled = -not $Running
    $applyStatusButton.Enabled = -not $Running
    $deleteUrlButton.Enabled = -not $Running
    $deleteExternalUrlButton.Enabled = -not $Running
    $resetButton.Enabled = -not $Running
    $urlQueueGrid.ReadOnly = $Running

    if ($Running) {
        $script:mainListOrderMode = "activity"
        Refresh-UrlGrid
        Set-AutoProgressPanelVisible -Visible $true
        $autoProgressPanel.BackColor = [System.Drawing.Color]::FromArgb(
            239,
            250,
            244
        )
        $autoStateLabel.Text = "● ĐANG CHẠY"
        $autoStateLabel.ForeColor = $green
        $autoStepLabel.Text = "Đang khởi động tiến trình submit..."
        $autoUrlLabel.Text = "URL hiện tại: đang lấy từ hàng đợi..."
    } else {
        Set-AutoProgressPanelVisible -Visible $false
    }
}

function Update-AutoProgressUi {
    param($Progress)

    if (-not $Progress) {
        return
    }

    $state = ([string]$Progress.state).ToUpperInvariant()
    $position = 0
    $total = 0
    $processed = 0
    $pending = 0
    [void][int]::TryParse([string]$Progress.position, [ref]$position)
    [void][int]::TryParse([string]$Progress.total, [ref]$total)
    if ($Progress.summary) {
        [void][int]::TryParse(
            [string]$Progress.summary.processed,
            [ref]$processed
        )
        [void][int]::TryParse(
            [string]$Progress.summary.pending,
            [ref]$pending
        )
        if ($total -lt 1) {
            [void][int]::TryParse(
                [string]$Progress.summary.selected,
                [ref]$total
            )
        }
    }

    if (-not $script:autoStartedAt -and [string]$Progress.startedAt) {
        $parsedStart = [datetime]::MinValue
        if (
            [datetime]::TryParseExact(
                [string]$Progress.startedAt,
                "yyyy-MM-dd HH:mm:ss",
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None,
                [ref]$parsedStart
            )
        ) {
            $script:autoStartedAt = $parsedStart
        }
    }

    $elapsed = [TimeSpan]::Zero
    if ($script:autoStartedAt) {
        $elapsed = (Get-Date) - $script:autoStartedAt
    } elseif ([string]$Progress.elapsedSeconds) {
        $elapsed = [TimeSpan]::FromSeconds(
            [double]$Progress.elapsedSeconds
        )
    }
    $elapsedText = "{0:00}:{1:00}:{2:00}" -f
        [math]::Floor($elapsed.TotalHours),
        $elapsed.Minutes,
        $elapsed.Seconds

    if ($total -gt 0) {
        $autoCountLabel.Text = (
            "{0} / {1} URL   •   Đã chạy {2}" -f
            $position,
            $total,
            $elapsedText
        )
        $percent = [math]::Floor(($processed * 100.0) / $total)
        $autoProgressBar.Value = [math]::Max(
            0,
            [math]::Min(100, $percent)
        )
    } else {
        $autoCountLabel.Text = "Đang chuẩn bị   •   $elapsedText"
        $autoProgressBar.Value = 0
    }

    $autoStepLabel.Text = [string]$Progress.message
    if ([string]$Progress.currentUrl) {
        $autoUrlLabel.Text = "URL hiện tại: $([string]$Progress.currentUrl)"
    } else {
        $autoUrlLabel.Text = "URL hiện tại: —"
    }

    switch ($state) {
        "RUNNING" {
            $script:autoPulseFrame = ($script:autoPulseFrame + 1) % 4
            $pulse = "." * ($script:autoPulseFrame + 1)
            $autoStateLabel.Text = "● ĐANG CHẠY$pulse"
            $autoStateLabel.ForeColor = $green
            $autoProgressPanel.BackColor = [System.Drawing.Color]::FromArgb(
                239,
                250,
                244
            )
        }
        "FINISHED" {
            $autoStateLabel.Text = "✓ ĐÃ HOÀN TẤT"
            $autoStateLabel.ForeColor = $green
            $autoProgressPanel.BackColor = [System.Drawing.Color]::FromArgb(
                239,
                250,
                244
            )
            $autoProgressBar.Value = 100
        }
        "STOPPED" {
            $autoStateLabel.Text = "■ KHÔNG CÓ TIẾN TRÌNH ĐANG CHẠY"
            $autoStateLabel.ForeColor = $red
            $autoProgressPanel.BackColor = [System.Drawing.Color]::FromArgb(
                255,
                244,
                244
            )
            if ($total -gt 0) {
                $autoStepLabel.Text = ((
                    "Phiên trước đã dừng: đã xử lý {0}/{1} URL. " +
                    "Bấm SUBMIT TỰ ĐỘNG để chạy tiếp."
                ) -f
                    $processed,
                    $total
                )
                $autoUrlLabel.Text = (
                    "Còn {0} URL CHƯA SUBMIT trong hàng đợi." -f $pending
                )
            } else {
                $autoStepLabel.Text = (
                    "Không có tiến trình đang chạy. " +
                    "Bấm SUBMIT TỰ ĐỘNG để bắt đầu."
                )
                $autoUrlLabel.Text = "URL hiện tại: —"
            }
        }
    }
}

function Start-SubmitChoice {
    param($Choice)

    if (-not $Choice -or -not $Choice.Profile) {
        $statusLabel.Text = "Gmail này chưa được gán vào profile."
        $statusLabel.ForeColor = $red
        return
    }

    if ($Choice.Profile -notmatch "^submit_(\d{2})$") {
        return
    }
    $profileNumber = [int]$Matches[1]

    $property = "https://$($Choice.Domain)/"
    $resourceId = [Uri]::EscapeDataString($property)
    $gscUrl = "https://search.google.com/u/0/search-console?resource_id=$resourceId&hl=vi"

    Start-Process `
        -FilePath "powershell.exe" `
        -WindowStyle Hidden `
        -ArgumentList @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", "`"$openProfileScript`"",
            "-Number", $profileNumber,
            "-Url", "`"$gscUrl`"",
            "-InspectionUrl", "`"$($Choice.InspectionUrl)`"",
            "-RequestIndexing"
        )
    $statusLabel.Text = "Đang mở $($Choice.Profile), kiểm tra URL và yêu cầu lập chỉ mục..."
    $statusLabel.ForeColor = $blue
}

function Start-OpenGscChoice {
    param($Choice)

    if (-not $Choice -or -not $Choice.Profile) {
        $statusLabel.Text = "Gmail này chưa được gán vào profile."
        $statusLabel.ForeColor = $red
        return
    }

    if ($Choice.Profile -notmatch "^submit_(\d{2})$") {
        return
    }
    $profileNumber = [int]$Matches[1]

    $property = "https://$($Choice.Domain)/"
    $resourceId = [Uri]::EscapeDataString($property)
    $gscUrl = "https://search.google.com/u/0/search-console?resource_id=$resourceId&hl=vi"

    Start-Process `
        -FilePath "powershell.exe" `
        -WindowStyle Hidden `
        -ArgumentList @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", "`"$openProfileScript`"",
            "-Number", $profileNumber,
            "-Url", "`"$gscUrl`""
        )
    $statusLabel.Text = "Đang mở GSC bằng $($Choice.Profile)..."
    $statusLabel.ForeColor = $blue
}

$searchButton.add_Click({
    $script:pendingSearchAction = if ($script:searchOnly) {
        "open"
    } else {
        "submit"
    }
    $searchChoiceCombo.Items.Clear()
    $script:searchChoiceMap = @()

    $domain = Get-DomainFromInput -Value $searchBox.Text
    if (-not $domain) {
        $statusLabel.Text = "Không đọc được tên miền từ nội dung đã nhập."
        $statusLabel.ForeColor = $red
        return
    }
    $inspectionUrl = $searchBox.Text.Trim()
    if ($inspectionUrl -notmatch "^[a-z][a-z0-9+.-]*://") {
        $inspectionUrl = "https://$inspectionUrl"
    }

    $matches = @()
    foreach ($account in @($script:config.accounts)) {
        foreach ($savedDomain in @($account.domains)) {
            $candidate = Get-DomainFromInput -Value ([string]$savedDomain)
            if (
                $domain -eq $candidate -or
                $domain.EndsWith(".$candidate")
            ) {
                $matches += $account
                break
            }
        }
    }
    $matches = @($matches | Sort-Object email -Unique)

    if ($matches.Count -eq 0) {
        $statusLabel.Text = "Tên miền $domain không tìm thấy Gmail/profile."
        $statusLabel.ForeColor = $red
        return
    }

    foreach ($account in $matches) {
        $profile = $script:config.profiles |
            Where-Object { $_.email -eq $account.email } |
            Select-Object -First 1
        $profileText = if ($profile) {
            [string]$profile.id
        } else {
            "chưa gán profile"
        }
        $searchChoiceCombo.Items.Add(
            "$($account.email)  |  $profileText"
        ) | Out-Null
        $script:searchChoiceMap += [PSCustomObject]@{
            Domain = $domain
            Email = [string]$account.email
            Profile = $(if ($profile) { [string]$profile.id } else { "" })
            InspectionUrl = $inspectionUrl
        }
    }

    $statusLabel.Text = "Đã tìm thấy $($matches.Count) Gmail/profile cho $domain."
    $statusLabel.ForeColor = $blue
    $validIndexes = @()
    for ($index = 0; $index -lt $script:searchChoiceMap.Count; $index++) {
        if ($script:searchChoiceMap[$index].Profile) {
            $validIndexes += $index
        }
    }

    if ($matches.Count -eq 1) {
        Load-Gmail -Email ([string]$matches[0].email)
        $profile = $script:config.profiles |
            Where-Object { $_.email -eq $matches[0].email } |
            Select-Object -First 1
        if ($profile) {
            $viewProfileCombo.SelectedItem = [string]$profile.id
        } else {
            $viewProfileCombo.SelectedIndex = 0
        }
    }

    if ($validIndexes.Count -eq 1) {
        $searchChoiceCombo.SelectedIndex = $validIndexes[0]
        if ($script:pendingSearchAction -eq "open") {
            Start-OpenGscChoice `
                -Choice $script:searchChoiceMap[$validIndexes[0]]
        } else {
            Start-SubmitChoice `
                -Choice $script:searchChoiceMap[$validIndexes[0]]
        }
    } elseif ($validIndexes.Count -gt 1) {
        $searchChoiceCombo.SelectedIndex = -1
        $searchChoiceCombo.DroppedDown = $true
        $statusLabel.Text = "Tên miền có nhiều profile. Hãy chọn profile muốn submit."
        $statusLabel.ForeColor = $blue
    } else {
        $searchChoiceCombo.SelectedIndex = 0
        $statusLabel.Text = "Các Gmail tìm thấy chưa được gán profile."
        $statusLabel.ForeColor = $red
    }
})

$openGscButton.add_Click({
    $script:pendingSearchAction = "open"

    if ($searchChoiceCombo.SelectedIndex -ge 0) {
        $choice = $script:searchChoiceMap[$searchChoiceCombo.SelectedIndex]
        Start-OpenGscChoice -Choice $choice
        return
    }

    if ($script:searchChoiceMap.Count -gt 1) {
        $searchChoiceCombo.DroppedDown = $true
        $statusLabel.Text = "Hãy chọn profile muốn mở GSC."
        return
    }

    $script:searchOnly = $true
    try {
        $searchButton.PerformClick()
    } finally {
        $script:searchOnly = $false
    }
})

$searchChoiceCombo.add_SelectionChangeCommitted({
    if (
        $script:searchChoiceMap.Count -gt 1 -and
        $searchChoiceCombo.SelectedIndex -ge 0
    ) {
        $choice = $script:searchChoiceMap[$searchChoiceCombo.SelectedIndex]
        if ($script:pendingSearchAction -eq "open") {
            Start-OpenGscChoice -Choice $choice
        } else {
            Start-SubmitChoice -Choice $choice
        }
    }
})

$searchBox.add_KeyDown({
    param($sender, $eventArgs)
    if ($eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $searchButton.PerformClick()
        $eventArgs.SuppressKeyPress = $true
    }
})

$domainLookupButton.add_Click({ Show-DomainAssignments })
$domainLookupBox.add_KeyDown({
    param($sender, $eventArgs)
    if ($eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Show-DomainAssignments
        $eventArgs.SuppressKeyPress = $true
    }
})

$gmailCombo.add_SelectedIndexChanged({
    if (-not $script:loading) {
        Load-Gmail -Email (Get-GmailComboEmail)
    }
})

$gmailCombo.add_Leave({
    $account = Get-Account `
        -Config $script:config `
        -Email (Get-GmailComboEmail)
    if ($account -and $account.email -ne $script:originalEmail) {
        Load-Gmail -Email ([string]$account.email)
    }
})

$gmailOkButton.add_Click({
    $selectedEmail = Get-GmailComboEmail
    $account = Get-Account -Config $script:config -Email $selectedEmail
    if ($account) {
        Load-Gmail -Email ([string]$account.email)
    } else {
        $statusLabel.Text = "Không tìm thấy Gmail: $selectedEmail"
        $statusLabel.ForeColor = $red
    }
})

$viewProfileCombo.add_SelectionChangeCommitted({
    if ($viewProfileCombo.SelectedIndex -le 0) {
        return
    }

    $profileId = [string]$viewProfileCombo.SelectedItem
    $profile = $script:config.profiles |
        Where-Object { $_.id -eq $profileId } |
        Select-Object -First 1
    $email = ([string]$profile.email).Trim()

    if ($email) {
        Load-Gmail -Email $email
        $statusLabel.Text = "$profileId đang được gán với $email."
    } else {
        $script:originalEmail = ""
        $gmailCombo.SelectedIndex = -1
        $gmailCombo.Text = ""
        $domainBox.Clear()
        $profileCombo.SelectedItem = $profileId
        $statusLabel.Text = "$profileId hiện chưa được gán Gmail."
    }
})

$newButton.add_Click({
    $script:originalEmail = ""
    $gmailCombo.SelectedIndex = -1
    $gmailCombo.Text = ""
    $domainBox.Clear()
    $profileCombo.SelectedIndex = 0
    $statusLabel.Text = "Nhập Gmail mới, nhập tên miền rồi bấm Lưu thay đổi."
    $gmailCombo.Focus()
})

$saveButton.add_Click({
    try {
        $email = (Get-GmailComboEmail).Trim()
        if ($email -notmatch "^[^@\s]+@[^@\s]+\.[^@\s]+$") {
            throw "Gmail chưa đúng định dạng."
        }

        $domains = @(
            $domainBox.Text -split "[,;`r`n]+" |
            ForEach-Object { $_.Trim().ToLowerInvariant() } |
            Where-Object { $_ } |
            Sort-Object -Unique
        )

        $account = if ($script:originalEmail) {
            Get-Account -Config $script:config -Email $script:originalEmail
        } else {
            Get-Account -Config $script:config -Email $email
        }

        $emailOwner = Get-Account -Config $script:config -Email $email
        if ($emailOwner -and $account -and $emailOwner -ne $account) {
            throw "Gmail này đã tồn tại."
        }

        if ($account) {
            $account.email = $email
            $account.domains = $domains
        } else {
            $account = [PSCustomObject]@{
                email = $email
                domains = $domains
            }
            $script:config.accounts = @($script:config.accounts) + $account
        }

        foreach ($profile in @($script:config.profiles)) {
            if (
                [string]::Equals(
                    ([string]$profile.email).Trim(),
                    $script:originalEmail,
                    [System.StringComparison]::OrdinalIgnoreCase
                ) -or
                [string]::Equals(
                    ([string]$profile.email).Trim(),
                    $email,
                    [System.StringComparison]::OrdinalIgnoreCase
                )
            ) {
                $profile.email = ""
            }
        }

        if ($profileCombo.SelectedIndex -gt 0) {
            $targetId = [string]$profileCombo.SelectedItem
            $target = $script:config.profiles |
                Where-Object { $_.id -eq $targetId } |
                Select-Object -First 1
            $target.email = $email
        }

        Save-Config -Config $script:config
        $script:originalEmail = $email
        Refresh-GmailList -SelectEmail $email
        Load-Gmail -Email $email
        $statusLabel.Text += "  Đã lưu."
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Không thể lưu:`r`n$($_.Exception.Message)",
            "Lỗi",
            "OK",
            "Error"
        ) | Out-Null
    }
})

$deleteButton.add_Click({
    if (-not $script:originalEmail) {
        return
    }

    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Xóa Gmail $script:originalEmail và bỏ gán khỏi profile?",
        "Xác nhận xóa",
        "YesNo",
        "Warning"
    )
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    $emailToDelete = $script:originalEmail
    $script:config.accounts = @($script:config.accounts |
        Where-Object { $_.email -ne $emailToDelete })
    foreach ($profile in @($script:config.profiles)) {
        if ($profile.email -eq $emailToDelete) {
            $profile.email = ""
        }
    }
    Save-Config -Config $script:config
    Refresh-GmailList
    $newButton.PerformClick()
    $statusLabel.Text = "Đã xóa $emailToDelete."
})

$openProfileButton.add_Click({
    if ($profileCombo.SelectedIndex -le 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Gmail này chưa được gán vào profile nào.",
            "Chưa có profile",
            "OK",
            "Information"
        ) | Out-Null
        return
    }

    $profileId = [string]$profileCombo.SelectedItem
    if ($profileId -notmatch "^submit_(\d{2})$") {
        return
    }

    Start-Process `
        -FilePath "powershell.exe" `
        -WindowStyle Hidden `
        -ArgumentList @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", "`"$openProfileScript`"",
            "-Number", [int]$Matches[1]
        )
    $statusLabel.Text = "Đang mở $profileId..."
})

$quickAssignButton.add_Click({
    try {
        if ($viewProfileCombo.SelectedIndex -le 0) {
            throw "Hãy chọn profile ở ô Xem ngược theo profile."
        }

        $email = (Get-GmailComboEmail).Trim()
        $account = Get-Account -Config $script:config -Email $email
        if (-not $account) {
            throw "Hãy chọn một Gmail đã có trong danh sách."
        }

        $targetId = [string]$viewProfileCombo.SelectedItem
        foreach ($profile in @($script:config.profiles)) {
            if ([string]::Equals(
                ([string]$profile.email).Trim(),
                $email,
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
                $profile.email = ""
            }
        }

        $target = $script:config.profiles |
            Where-Object { $_.id -eq $targetId } |
            Select-Object -First 1
        $target.email = $email

        Save-Config -Config $script:config
        Refresh-GmailList -SelectEmail $email
        Load-Gmail -Email $email
        $viewProfileCombo.SelectedItem = $targetId
        $statusLabel.Text = "Đã gán nhanh $email vào $targetId."
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Chưa thể gán nhanh",
            "OK",
            "Information"
        ) | Out-Null
    }
})

$addUrlsButton.add_Click({
    $urlImportLabel.Text = "DÁN URL MỚI — mỗi dòng một URL"
    $urlImportLabel.ForeColor = $navy
    if (-not $urlImportForm.Visible) {
        $urlImportForm.Show($form)
    }
    $urlImportForm.Activate()
    $urlImportBox.Focus()
})

$importUrlsButton.add_Click({
    try {
        $lines = @(
            $urlImportBox.Text -split "\r?\n" |
                ForEach-Object { $_.Trim() } |
                Where-Object { $_ }
        )

        if ($lines.Count -eq 0) {
            $urlImportLabel.Text = "Chưa có URL nào để nạp."
            $urlImportLabel.ForeColor = $red
            return
        }

        $seen = @{}
        foreach ($entry in @($script:urlEntries)) {
            $url = ([string]$entry.url).Trim()
            if ($url) {
                $seen[$url.ToLowerInvariant()] = $true
            }
        }

        $added = 0
        $duplicate = 0
        $invalid = 0
        $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        foreach ($line in $lines) {
            $url = Get-CleanUrl -Value $line
            if (-not $url) {
                $invalid++
                continue
            }

            $key = $url.ToLowerInvariant()
            if ($seen.ContainsKey($key)) {
                $duplicate++
                continue
            }

            $script:urlEntries += [PSCustomObject]@{
                url = $url
                status = "PENDING"
                message = ""
                updatedAt = $now
            }
            $seen[$key] = $true
            $added++
        }

        Save-UrlEntries -Entries $script:urlEntries
        Refresh-UrlGrid
        if ($added -gt 0) {
            $urlImportBox.Clear()
        }
        $statusLabel.ForeColor = $muted
        $statusLabel.Text = (
            "Đã nạp {0} URL mới; bỏ qua {1} URL trùng; {2} dòng không hợp lệ." -f
            $added,
            $duplicate,
            $invalid
        )
        $urlImportForm.Hide()
    } catch {
        $urlImportLabel.Text = "Không thể nạp: $($_.Exception.Message)"
        $urlImportLabel.ForeColor = $red
    }
})

$openUrlListButton.add_Click({
    Refresh-UrlGrid
    if (-not $urlListForm.Visible) {
        $urlListForm.Show($form)
    }
    $urlListForm.Activate()
})

$allUrlsList.add_SelectedIndexChanged({
    $index = $allUrlsList.SelectedIndex
    if ($index -lt 0 -or $index -ge $script:mainUrlChoiceMap.Count) {
        $script:selectedMainUrl = ""
        $quickSubmitUrl.Text = "Bấm một URL trong danh sách để chọn."
        $submitSelectedButton.Enabled = $false
        $priorityButton.Enabled = $false
        $copySelectedUrlButton.Enabled = $false
        return
    }

    $entry = $script:mainUrlChoiceMap[$index]
    $script:selectedMainUrl = [string]$entry.url
    $quickSubmitUrl.Text = (
        "[{0}]`r`n{1}" -f
        (Get-UrlStatusLabel -Code ([string]$entry.status)),
        [string]$entry.url
    )
    $isRunning = $script:autoProcess -and -not $script:autoProcess.HasExited
    $submitSelectedButton.Enabled = -not $isRunning
    $priorityButton.Enabled = -not $isRunning
    $copySelectedUrlButton.Enabled = $true
})

$resetOuterOrderButton.add_Click({
    $script:mainListOrderMode = "queue"
    Refresh-UrlGrid
    $statusLabel.ForeColor = $blue
    $statusLabel.Text = "Danh sách ngoài đang xếp lại theo STT gốc."
})

$copyCurrentSelectedUrl = {
    if (-not $script:selectedMainUrl) {
        return
    }
    [System.Windows.Forms.Clipboard]::SetText($script:selectedMainUrl)
    $statusLabel.ForeColor = $green
    $statusLabel.Text = "Đã copy URL. Dán vào GSC bằng Ctrl+V để kiểm tra."
}

$copySelectedUrlButton.add_Click($copyCurrentSelectedUrl)
$copyUrlMenuItem.add_Click($copyCurrentSelectedUrl)

$allUrlsList.add_MouseDown({
    param($sender, $eventArgs)
    $index = $sender.IndexFromPoint($eventArgs.Location)
    if ($index -ge 0) {
        $sender.SelectedIndex = $index
    }
})

$allUrlsList.add_KeyDown({
    param($sender, $eventArgs)
    if ($eventArgs.Control -and $eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::C) {
        & $copyCurrentSelectedUrl
        $eventArgs.SuppressKeyPress = $true
    }
})

$priorityButton.add_Click({
    if (-not $script:selectedMainUrl) {
        return
    }
    $entry = @($script:urlEntries | Where-Object {
        [string]$_.url -eq $script:selectedMainUrl
    }) | Select-Object -First 1
    if (-not $entry) {
        return
    }
    $entry.status = "PRIORITY"
    $entry.message = "Ưu tiên chạy trước trong phiên tự động tiếp theo."
    $entry.updatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Save-UrlEntries -Entries $script:urlEntries
    Refresh-UrlGrid
    $statusLabel.ForeColor = $blue
    $statusLabel.Text = "Đã ưu tiên URL đã chọn; chạy tự động sẽ lấy URL này trước."
})

$submitSelectedButton.add_Click({
    try {
        if (-not $script:selectedMainUrl) {
            throw "Hãy chọn một URL trong danh sách."
        }
        if ($script:autoProcess -and -not $script:autoProcess.HasExited) {
            throw "Đang có tiến trình submit khác chạy."
        }
        if (Test-Path -LiteralPath $autoStopFlagPath) {
            Remove-Item -LiteralPath $autoStopFlagPath -Force
        }
        if (Test-Path -LiteralPath $autoProgressPath) {
            Remove-Item -LiteralPath $autoProgressPath -Force
        }

        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = "python"
        $browserMode = Get-BrowserRunMode
        $processInfo.Arguments = (
            '"{0}" --limit 1 --url "{1}" --browser-mode {2}' -f
            $autoSubmitScript,
            $script:selectedMainUrl,
            $browserMode
        )
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true
        $script:autoProcess = [System.Diagnostics.Process]::Start($processInfo)
        $script:lastAutoProgress = ""
        $script:autoStartedAt = Get-Date
        $script:autoPulseFrame = 0
        Set-AutoUiState -Running $true
        $autoCountLabel.Text = "1 URL đã chọn   •   Đã chạy 00:00:00"
        $autoProgressBar.Value = 0
        $statusLabel.ForeColor = $blue
        $statusLabel.Text = "Đang submit riêng URL đã chọn..."
        $autoProgressTimer.Start()
    } catch {
        $statusLabel.ForeColor = $red
        $statusLabel.Text = "Không thể submit URL đã chọn: $($_.Exception.Message)"
    }
})

$autoSubmitButton.add_Click({
    try {
        Save-UrlGrid
        $pendingCount = @(
            $script:urlEntries |
                Where-Object {
                    ([string]$_.status) -in @("PENDING", "PRIORITY")
                }
        ).Count
        if ($pendingCount -eq 0) {
            $statusLabel.ForeColor = $red
            $statusLabel.Text = "Không có URL chờ submit."
            return
        }

        # Một lần bấm chạy toàn bộ hàng đợi còn chờ; URL ưu tiên được lấy trước.
        $limit = $pendingCount

        if (Test-Path -LiteralPath $autoStopFlagPath) {
            Remove-Item -LiteralPath $autoStopFlagPath -Force
        }
        if (Test-Path -LiteralPath $autoProgressPath) {
            Remove-Item -LiteralPath $autoProgressPath -Force
        }

        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = "python"
        $browserMode = Get-BrowserRunMode
        $processInfo.Arguments = (
            '"{0}" --limit {1} --browser-mode {2}' -f
            $autoSubmitScript, $limit, $browserMode
        )
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true
        $script:autoProcess = [System.Diagnostics.Process]::Start($processInfo)
        $script:lastAutoProgress = ""
        $script:autoStartedAt = Get-Date
        $script:autoPulseFrame = 0
        Set-AutoUiState -Running $true
        $autoCountLabel.Text = "0 / $limit URL   •   Đã chạy 00:00:00"
        $autoProgressBar.Value = 0
        $statusLabel.ForeColor = $muted
        $modeText = if ($browserMode -eq "hidden") {
            "chạy ẩn"
        } else {
            "hiện trình duyệt"
        }
        $statusLabel.Text = "Đang khởi động SUBMIT TỰ ĐỘNG ($modeText) cho $limit URL..."
        $autoProgressTimer.Start()
    } catch {
        Set-AutoUiState -Running $false
        $statusLabel.ForeColor = $red
        $statusLabel.Text = "Không thể chạy tự động: $($_.Exception.Message)"
    }
})

$stopAutoButton.add_Click({
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText(
            $autoStopFlagPath,
            (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
            $utf8
        )
        $stopAutoButton.Enabled = $false
        $autoStateLabel.Text = "● ĐANG DỪNG..."
        $autoStateLabel.ForeColor = $amber
        $autoProgressPanel.BackColor = [System.Drawing.Color]::FromArgb(
            255,
            249,
            232
        )
        $autoStepLabel.Text = (
            "Đã nhận lệnh dừng; sẽ hoàn tất URL hiện tại rồi dừng."
        )
        $statusLabel.ForeColor = $red
        $statusLabel.Text = "Đã yêu cầu dừng; hệ thống sẽ dừng sau URL hiện tại."
    } catch {
        $statusLabel.Text = "Không thể gửi yêu cầu dừng: $($_.Exception.Message)"
    }
})

$stopNowButton.add_Click({
    $answer = [System.Windows.Forms.MessageBox]::Show(
        (
            "Dừng ngay sẽ ngắt tiến trình giữa bước đang chạy.`r`n`r`n" +
            "URL đang xử lý sẽ chuyển sang LỖI vì chưa thể xác nhận " +
            "Google đã nhận yêu cầu hay chưa.`r`n`r`nDừng ngay?"
        ),
        "XÁC NHẬN DỪNG NGAY",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText(
            $autoStopFlagPath,
            (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
            $utf8
        )

        if (
            $script:autoProcess -and
            -not $script:autoProcess.HasExited
        ) {
            try {
                $script:autoProcess.Kill()
            } catch {
                if (-not $script:autoProcess.HasExited) {
                    throw
                }
            }
            [void]$script:autoProcess.WaitForExit(3000)
        }
        if (
            $script:autoProcess -and
            -not $script:autoProcess.HasExited
        ) {
            throw "Tiến trình Python chưa chịu dừng."
        }

        $progress = $null
        if (Test-Path -LiteralPath $autoProgressPath) {
            try {
                $progress = Get-Content `
                    -LiteralPath $autoProgressPath `
                    -Raw `
                    -Encoding UTF8 |
                    ConvertFrom-Json
            } catch {
                $progress = $null
            }
        }

        $script:urlEntries = @(Read-UrlEntries)
        $interruptedCount = 0
        $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        foreach ($entry in @($script:urlEntries)) {
            if (([string]$entry.status).ToUpperInvariant() -eq "RUNNING") {
                $entry.status = "ERROR"
                $entry.message = (
                    "Đã DỪNG NGAY khi đang xử lý; kết quả Google chưa " +
                    "được xác nhận. Hãy kiểm tra thủ công trước khi chạy lại."
                )
                $entry.updatedAt = $now
                $interruptedCount++
            }
        }
        Save-UrlEntries -Entries $script:urlEntries

        $runId = ""
        $logPath = $script:lastLogPath
        if ($progress) {
            $runId = [string]$progress.runId
            if ([string]$progress.logPath) {
                $logPath = [string]$progress.logPath
            }
        }
        $script:lastLogPath = $logPath

        if ($logPath -and (Test-Path -LiteralPath $logPath)) {
            $event = [ordered]@{
                timestamp = $now
                runId = $runId
                step = "RUN_FORCE_STOPPED"
                reason = "Người dùng bấm DỪNG NGAY."
                interruptedUrls = $interruptedCount
            }
            $writer = [System.IO.StreamWriter]::new(
                $logPath,
                $true,
                $utf8
            )
            try {
                $writer.WriteLine(
                    ($event | ConvertTo-Json -Compress)
                )
            } finally {
                $writer.Dispose()
            }
        }

        $statuses = @(
            $script:urlEntries |
                ForEach-Object { ([string]$_.status).ToUpperInvariant() }
        )
        $stoppedProgress = [ordered]@{
            runId = $runId
            state = "STOPPED"
            message = "Đã DỪNG NGAY tiến trình submit."
            currentUrl = ""
            updatedAt = $now
            logPath = $logPath
            summary = [ordered]@{
                completed = @($statuses | Where-Object { $_ -eq "SUBMITTED" }).Count
                errors = @($statuses | Where-Object { $_ -eq "ERROR" }).Count
                skipped = @($statuses | Where-Object { $_ -eq "SKIPPED" }).Count
                quota = @($statuses | Where-Object { $_ -eq "QUOTA" }).Count
                unmapped = @($statuses | Where-Object { $_ -eq "UNMAPPED" }).Count
                pending = @($statuses | Where-Object { $_ -eq "PENDING" }).Count
            }
        }
        [System.IO.File]::WriteAllText(
            $autoProgressPath,
            ($stoppedProgress | ConvertTo-Json -Depth 6),
            $utf8
        )

        $autoProgressTimer.Stop()
        $script:autoProcess = $null
        Set-AutoUiState -Running $false
        Update-AutoProgressUi -Progress ([PSCustomObject]$stoppedProgress)
        Refresh-UrlGrid
        $statusLabel.ForeColor = $red
        $statusLabel.Text = (
            "Đã DỪNG NGAY. Có {0} URL đang chạy được chuyển sang LỖI; " +
            "hãy kiểm tra trước khi chạy lại." -f $interruptedCount
        )
    } catch {
        $statusLabel.ForeColor = $red
        $statusLabel.Text = "Không thể dừng ngay: $($_.Exception.Message)"
    }
})

$autoProgressTimer.add_Tick({
    try {
        if (Test-Path -LiteralPath $autoProgressPath) {
            $progress = Get-Content `
                -LiteralPath $autoProgressPath `
                -Raw `
                -Encoding UTF8 |
                ConvertFrom-Json
            $progressKey = "{0}|{1}|{2}" -f
                [string]$progress.updatedAt,
                [string]$progress.state,
                [string]$progress.message
            Update-AutoProgressUi -Progress $progress
            if ($progressKey -ne $script:lastAutoProgress) {
                $script:lastAutoProgress = $progressKey
                $script:lastLogPath = [string]$progress.logPath
                $statusLabel.ForeColor = $muted
                $statusLabel.Text = [string]$progress.message

                $script:urlEntries = @(Read-UrlEntries)
                Refresh-UrlGrid
                Highlight-MainUrl -Url ([string]$progress.currentUrl)
            }

            if (
                ([string]$progress.state -eq "FINISHED") -or
                ([string]$progress.state -eq "STOPPED")
            ) {
                $autoProgressTimer.Stop()
                Set-AutoUiState -Running $false
                Update-AutoProgressUi -Progress $progress
                $script:urlEntries = @(Read-UrlEntries)
                Refresh-UrlGrid
                Highlight-MainUrl -Url ([string]$progress.currentUrl)

                $summary = $progress.summary
                $statusLabel.Text = (
                    "{0} Hoàn thành: {1}; Lỗi: {2}; Bỏ qua: {3}; Log: {4}" -f
                    [string]$progress.message,
                    [int]$summary.completed,
                    [int]$summary.errors,
                    [int]$summary.skipped,
                    [string]$progress.logPath
                )
            }
        }

        if (
            $script:autoProcess -and
            $script:autoProcess.HasExited -and
            $autoProgressTimer.Enabled
        ) {
            $autoProgressTimer.Stop()
            Set-AutoUiState -Running $false
            $script:urlEntries = @(Read-UrlEntries)
            Refresh-UrlGrid
            if (-not (Test-Path -LiteralPath $autoProgressPath)) {
                $statusLabel.ForeColor = $red
                $statusLabel.Text = "Tiến trình tự động đã dừng nhưng không tạo báo cáo."
                $autoStateLabel.Text = "! DỪNG BẤT THƯỜNG"
                $autoStateLabel.ForeColor = $red
                $autoStepLabel.Text = $statusLabel.Text
                $autoProgressPanel.BackColor = [System.Drawing.Color]::FromArgb(
                    255,
                    244,
                    244
                )
            } else {
                $lastState = ""
                try {
                    $lastState = [string](
                        Get-Content `
                            -LiteralPath $autoProgressPath `
                            -Raw `
                            -Encoding UTF8 |
                            ConvertFrom-Json
                    ).state
                } catch {
                }
                if ($lastState -eq "RUNNING") {
                    $statusLabel.ForeColor = $red
                    $statusLabel.Text = (
                        "Tiến trình đã thoát bất thường trước khi hoàn tất."
                    )
                    $autoStateLabel.Text = "! DỪNG BẤT THƯỜNG"
                    $autoStateLabel.ForeColor = $red
                    $autoStepLabel.Text = $statusLabel.Text
                    $autoProgressPanel.BackColor = (
                        [System.Drawing.Color]::FromArgb(255, 244, 244)
                    )
                }
            }
        }
    } catch {
        $statusLabel.ForeColor = $red
        $statusLabel.Text = "Không đọc được tiến trình tự động: $($_.Exception.Message)"
    }
})

$viewLogButton.add_Click({
    if ($script:lastLogPath -and (Test-Path -LiteralPath $script:lastLogPath)) {
        Start-Process -FilePath "notepad.exe" -ArgumentList @(
            "`"$($script:lastLogPath)`""
        )
        return
    }

    $logsPath = Join-Path $PSScriptRoot "logs"
    if (Test-Path -LiteralPath $logsPath) {
        Start-Process -FilePath "explorer.exe" -ArgumentList @(
            "`"$logsPath`""
        )
    } else {
        $urlQueueSummary.Text = "Chưa có file log nào."
    }
})

$urlFilterCombo.add_SelectedIndexChanged({
    if ($urlFilterCombo.SelectedItem) {
        $script:urlFilter = [string]$urlFilterCombo.SelectedItem
        Refresh-UrlGrid
    }
})

$saveUrlListButton.add_Click({
    try {
        Save-UrlGrid
        $urlQueueSummary.Text += "   —   Đã lưu."
    } catch {
        $urlQueueSummary.Text = "Không thể lưu: $($_.Exception.Message)"
        $urlQueueSummary.ForeColor = $red
    }
})

$applyStatusButton.add_Click({
    try {
        if ($urlQueueGrid.SelectedRows.Count -eq 0) {
            $urlQueueSummary.Text = "Hãy chọn ít nhất một dòng."
            return
        }

        $selectedStatus = [string]$bulkStatusCombo.SelectedItem
        foreach ($row in @($urlQueueGrid.SelectedRows)) {
            $row.Cells["Status"].Value = $selectedStatus
            $row.Cells["UpdatedAt"].Value = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        Save-UrlGrid
        $urlQueueSummary.Text += "   —   Đã đổi trạng thái các dòng được chọn."
    } catch {
        $urlQueueSummary.Text = "Không thể đổi trạng thái: $($_.Exception.Message)"
        $urlQueueSummary.ForeColor = $red
    }
})

$deleteUrlButton.add_Click({
    try {
        $deleteKeys = @{}
        foreach ($row in @($urlQueueGrid.SelectedRows)) {
            $originalUrl = ([string]$row.Tag).Trim()
            if (-not $originalUrl) {
                $originalUrl = ([string]$row.Cells["Url"].Value).Trim()
            }
            if ($originalUrl) {
                $deleteKeys[$originalUrl.ToLowerInvariant()] = $true
            }
        }

        $script:urlEntries = @(
            $script:urlEntries |
                Where-Object {
                    -not $deleteKeys.ContainsKey(
                        ([string]$_.url).Trim().ToLowerInvariant()
                    )
                }
        )
        Save-UrlEntries -Entries $script:urlEntries
        Refresh-UrlGrid
    } catch {
        $urlQueueSummary.Text = "Không thể xóa: $($_.Exception.Message)"
        $urlQueueSummary.ForeColor = $red
    }
})

$deleteExternalUrlButton.add_Click({
    try {
        $currentCount = @($script:urlEntries).Count
        if ($currentCount -eq 0) {
            $urlQueueSummary.Text = "Danh sách URL đang trống."
            return
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Xóa toàn bộ $currentCount URL đang có trong danh sách?`n`nThao tác sẽ lưu bản sao .bak trước khi ghi.",
            "Xác nhận xóa nhanh toàn bộ URL",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $script:urlEntries = @()
        Save-UrlEntries -Entries $script:urlEntries
        Refresh-UrlGrid
        $urlQueueSummary.Text = "Đã xóa toàn bộ $currentCount URL."
        return

        # Giữ phần xử lý cũ phía dưới để không làm thay đổi cấu trúc event.
        $knownDomains = @(
            foreach ($account in @($script:config.accounts)) {
                foreach ($domain in @($account.domains)) {
                    $cleanDomain = Get-DomainFromInput ([string]$domain)
                    if ($cleanDomain) { $cleanDomain }
                }
            }
        ) | Sort-Object -Unique

        $externalEntries = @(
            $script:urlEntries | Where-Object {
                $url = Get-CleanUrl ([string]$_.url)
                if (-not $url) { return $true }
                try {
                    $hostName = ([Uri]$url).Host.Trim(".").ToLowerInvariant()
                    $isKnown = $false
                    foreach ($domain in $knownDomains) {
                        if ($hostName -eq $domain -or $hostName.EndsWith(".$domain")) {
                            $isKnown = $true
                            break
                        }
                    }
                    return -not $isKnown
                } catch {
                    return $true
                }
            }
        )

        if ($externalEntries.Count -eq 0) {
            $urlQueueSummary.Text = "Không có URL ngoài domain để xóa."
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Tìm thấy $($externalEntries.Count) URL ngoài các domain đã khai báo. Xóa nhanh các URL này?`n`nThao tác sẽ lưu bản sao .bak trước khi ghi.",
            "Xác nhận xóa URL ngoài domain",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $externalKeys = @{}
        foreach ($entry in $externalEntries) {
            $externalKeys[([string]$entry.url).Trim().ToLowerInvariant()] = $true
        }
        $script:urlEntries = @(
            $script:urlEntries | Where-Object {
                -not $externalKeys.ContainsKey(([string]$_.url).Trim().ToLowerInvariant())
            }
        )
        Save-UrlEntries -Entries $script:urlEntries
        Refresh-UrlGrid
        $urlQueueSummary.Text = "Đã xóa $($externalEntries.Count) URL ngoài domain."
    } catch {
        $urlQueueSummary.Text = "Không thể xóa URL ngoài domain: $($_.Exception.Message)"
        $urlQueueSummary.ForeColor = $red
    }
})

$urlQueueGrid.add_DataError({
    param($sender, $eventArgs)
    $eventArgs.ThrowException = $false
})

$urlListForm.add_FormClosing({
    param($sender, $eventArgs)
    $eventArgs.Cancel = $true
    $urlListForm.Hide()
})

$urlImportForm.add_FormClosing({
    param($sender, $eventArgs)
    $eventArgs.Cancel = $true
    $urlImportForm.Hide()
})

$resetButton.add_Click({
    if ($script:autoProcess -and -not $script:autoProcess.HasExited) {
        $statusLabel.ForeColor = $red
        $statusLabel.Text = "Không thể khởi động lại khi submit tự động đang chạy."
        return
    }
    try {
        if (-not (Test-Path -LiteralPath $restartLauncher)) {
            throw "Không tìm thấy file khởi động lại app."
        }
        Start-Process -FilePath "wscript.exe" -ArgumentList @(
            "`"$restartLauncher`"",
            "`"$PSCommandPath`""
        ) -WindowStyle Hidden
        $form.Close()
    } catch {
        $statusLabel.ForeColor = $red
        $statusLabel.Text = "Không thể khởi động lại app: $($_.Exception.Message)"
    }
})

$form.add_FormClosing({
    if (
        $script:autoProcess -and
        -not $script:autoProcess.HasExited
    ) {
        try {
            $utf8 = New-Object System.Text.UTF8Encoding($false)
            [System.IO.File]::WriteAllText(
                $autoStopFlagPath,
                (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
                $utf8
            )
        } catch {
        }
    }
})

$form.add_Shown({
    Update-SubmitLayout
    Refresh-GmailList
    Refresh-UrlGrid
    if ($gmailCombo.Items.Count -gt 0) {
        $gmailCombo.SelectedIndex = 0
        Load-Gmail -Email (Get-GmailComboEmail)
    }

    if (Test-Path -LiteralPath $autoProgressPath) {
        try {
            $savedProgress = Get-Content `
                -LiteralPath $autoProgressPath `
                -Raw `
                -Encoding UTF8 |
                ConvertFrom-Json
            $savedState = ([string]$savedProgress.state).ToUpperInvariant()
            Update-AutoProgressUi -Progress $savedProgress
            $script:lastLogPath = [string]$savedProgress.logPath

            if ($savedState -eq "RUNNING") {
                $savedPid = 0
                [void][int]::TryParse(
                    [string]$savedProgress.processId,
                    [ref]$savedPid
                )
                $runningProcess = $null
                if ($savedPid -gt 0) {
                    try {
                        $runningProcess = (
                            [System.Diagnostics.Process]::GetProcessById(
                                $savedPid
                            )
                        )
                    } catch {
                        $runningProcess = $null
                    }
                }
                if (-not $runningProcess) {
                    try {
                        $matchingPython = Get-CimInstance `
                            -ClassName Win32_Process `
                            -Filter "Name = 'python.exe'" |
                            Where-Object {
                                ([string]$_.CommandLine) -like
                                    "*auto_submit_queue.py*"
                            } |
                            Select-Object -First 1
                        if ($matchingPython) {
                            $runningProcess = (
                                [System.Diagnostics.Process]::GetProcessById(
                                    [int]$matchingPython.ProcessId
                                )
                            )
                        }
                    } catch {
                        $runningProcess = $null
                    }
                }

                if ($runningProcess) {
                    $script:autoProcess = $runningProcess
                    Set-AutoUiState -Running $true
                    Update-AutoProgressUi -Progress $savedProgress
                    $statusLabel.Text = (
                        "Đã nối lại màn hình theo dõi tiến trình đang chạy."
                    )
                    $autoProgressTimer.Start()
                } else {
                    $autoStateLabel.Text = "! KHÔNG CÒN CHẠY"
                    $autoStateLabel.ForeColor = $red
                    $autoStepLabel.Text = (
                        "Báo cáo cuối vẫn là ĐANG CHẠY nhưng không còn " +
                        "tìm thấy tiến trình submit."
                    )
                    $autoProgressPanel.BackColor = (
                        [System.Drawing.Color]::FromArgb(255, 244, 244)
                    )
                }
            }
        } catch {
            $statusLabel.Text = (
                "Không đọc được tiến trình gần nhất: " +
                $_.Exception.Message
            )
        }
    }
})

[void]$form.ShowDialog()
