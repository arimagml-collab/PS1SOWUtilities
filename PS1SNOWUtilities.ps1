#requires -Version 5.1
<#
PS1SNOWUtilities.ps1
GUI tool for exporting ServiceNow table to CSV.
Stores settings in settings.json next to this script.
Secrets (password/apiKey) are stored encrypted via DPAPI (ConvertFrom-SecureString).

License: MIT License
Copyright (c) ixam.net (https://www.ixam.net)
Disclaimer: This software is an independent utility and is not affiliated with,
endorsed by, or guaranteed by ServiceNow.

Recommended shortcut target:
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "C:\path\PS1SNOWUtilities.ps1"
#>

try {
  Set-StrictMode -Version Latest
  $ErrorActionPreference = "Stop"

  # ----------------------------
  # Ensure STA (WinForms stability)
  # ----------------------------
  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  $apt = [System.Threading.Thread]::CurrentThread.ApartmentState
  if ($apt -ne [System.Threading.ApartmentState]::STA) {
    $ps = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
    $args = @(
      "-NoProfile",
      "-ExecutionPolicy", "Bypass",
      "-STA",
      "-WindowStyle", "Hidden",
      "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path)
    )
    Start-Process -FilePath $ps -ArgumentList $args | Out-Null
    return
  }

  Add-Type -AssemblyName System.Drawing | Out-Null
  [System.Windows.Forms.Application]::EnableVisualStyles()

  # ----------------------------
  # Paths / Settings
  # ----------------------------
  $ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
  $SettingsPath = Join-Path $ScriptDir "settings.json"
  $DefaultExportDir = Join-Path $ScriptDir "ExportedFiles"
  $DefaultAttachmentDir = Join-Path $ScriptDir "DownloadedAttachments"
  $DefaultLogDir = Join-Path (Get-Location).Path "Logs"

  # ----------------------------
  # Core modules
  # ----------------------------
  Import-Module (Join-Path $ScriptDir "modules/Core/Settings.psm1") -Force -Prefix Core
  Import-Module (Join-Path $ScriptDir "modules/Core/ServiceNowClient.psm1") -Force -Prefix Core
  Import-Module (Join-Path $ScriptDir "modules/Core/Logging.psm1") -Force -Prefix Core
  Import-Module (Join-Path $ScriptDir "modules/Core/I18n.psm1") -Force -Prefix Core

  function Import-OptionalFeatureModule {
    param(
      [Parameter(Mandatory=$true)][string]$ModulePath,
      [Parameter(Mandatory=$true)][string]$FeatureKey,
      [Parameter(Mandatory=$true)][string[]]$RequiredFunctions
    )

    if (-not (Test-Path $ModulePath)) {
      return [pscustomobject]@{ Enabled = $false; Message = ("Feature module not found: {0}" -f $ModulePath) }
    }

    Import-Module $ModulePath -Force

    foreach ($funcName in $RequiredFunctions) {
      if (-not (Get-Command -Name $funcName -CommandType Function -ErrorAction SilentlyContinue)) {
        return [pscustomobject]@{ Enabled = $false; Message = ("Feature module missing required function ({0}): {1}" -f $FeatureKey, $funcName) }
      }
    }

    return [pscustomobject]@{ Enabled = $true; Message = $null }
  }

  $featureModuleStatus = @{
    Export = Import-OptionalFeatureModule -ModulePath (Join-Path $ScriptDir "modules/Features/ExportFeature.psm1") -FeatureKey "Export" -RequiredFunctions @("Validate-ExportInput", "Invoke-ExportUseCase")
    ViewEditor = Import-OptionalFeatureModule -ModulePath (Join-Path $ScriptDir "modules/Features/ViewEditorFeature.psm1") -FeatureKey "ViewEditor" -RequiredFunctions @("Validate-ViewInput", "Invoke-CreateViewUseCase")
    Delete = Import-OptionalFeatureModule -ModulePath (Join-Path $ScriptDir "modules/Features/TruncateFeature.psm1") -FeatureKey "Delete" -RequiredFunctions @("Validate-TruncateInput", "Invoke-TruncateUseCase")
    AttachmentHarvester = Import-OptionalFeatureModule -ModulePath (Join-Path $ScriptDir "modules/Features/AttachmentHarvesterFeature.psm1") -FeatureKey "AttachmentHarvester" -RequiredFunctions @("Validate-AttachmentHarvesterInput", "Invoke-AttachmentHarvesterUseCase")
  }

  $script:IsExportFeatureEnabled = [bool]$featureModuleStatus.Export.Enabled
  $script:IsViewEditorFeatureEnabled = [bool]$featureModuleStatus.ViewEditor.Enabled
  $script:IsDeleteFeatureEnabled = [bool]$featureModuleStatus.Delete.Enabled
  $script:IsAttachmentHarvesterFeatureEnabled = [bool]$featureModuleStatus.AttachmentHarvester.Enabled
  $script:DisabledFeatureMessages = @($featureModuleStatus.Values | Where-Object { -not $_.Enabled -and -not [string]::IsNullOrWhiteSpace([string]$_.Message) } | ForEach-Object { [string]$_.Message })

  # ----------------------------
  # i18n
  # ----------------------------
  $LocalesDir = Join-Path $ScriptDir "locales"
  $script:I18N = Load-CoreI18nResources -LocalesDirectory $LocalesDir -DefaultLanguage "ja"
  $script:MissingI18nKeys = @{}

  function T([string]$key) {
    $lang = "ja"
    if ($script:Settings -and $script:Settings.uiLanguage) { $lang = [string]$script:Settings.uiLanguage }

    $text = Resolve-CoreI18nText -I18nResources $script:I18N -Language $lang -Key $key -DefaultLanguage "ja"
    $exists = Test-CoreI18nKeyExists -I18nResources $script:I18N -Language $lang -Key $key -DefaultLanguage "ja"
    if (-not $exists) {
      $missingToken = "{0}:{1}" -f $lang, $key
      if (-not $script:MissingI18nKeys.ContainsKey($missingToken)) {
        $script:MissingI18nKeys[$missingToken] = $true
        Add-Log ("Missing i18n key: {0} ({1})" -f $key, $lang)
      }
    }
    return $text
  }

  # ----------------------------
  # Settings service wrappers
  # ----------------------------
  function Protect-Secret([string]$plain) {
    return (Protect-CoreSecret -Plain $plain)
  }

  function Unprotect-Secret([string]$enc) {
    return (Unprotect-CoreSecret -Encrypted $enc)
  }

  function New-DefaultSettings {
    return (New-CoreDefaultSettings)
  }

  function Load-Settings {
    return (Load-CoreSettings -SettingsPath $SettingsPath)
  }

  function Save-Settings {
    Save-CoreSettings -Settings $script:Settings -SettingsPath $SettingsPath
  }

  function Initialize-SettingsDebounceTimer {
    if ($script:SettingsSaveTimer) { return }
    $script:SettingsSaveTimer = New-Object System.Windows.Forms.Timer
    $script:SettingsSaveTimer.Interval = 500
    $script:SettingsSaveTimer.add_Tick({
      $script:SettingsSaveTimer.Stop()
      Save-Settings
    })
  }

  function Request-SaveSettings([switch]$Immediate) {
    if ($Immediate) {
      if ($script:SettingsSaveTimer) { $script:SettingsSaveTimer.Stop() }
      Save-Settings
      return
    }
    Initialize-SettingsDebounceTimer
    $script:SettingsSaveTimer.Stop()
    $script:SettingsSaveTimer.Start()
  }

  $script:Settings = Load-Settings
  $script:ColumnCache = @{}
  $script:SettingsSaveTimer = $null

  # ----------------------------
  # ServiceNow REST helper wrappers
  # ----------------------------
  function UrlEncode([string]$s) {
    return (CoreUrlEncode -Value $s)
  }

  function Get-BaseUrl {
    return (Get-CoreBaseUrl -Settings $script:Settings)
  }

  function New-SnowHeaders {
    return (New-CoreSnowHeaders -Settings $script:Settings -UnprotectSecret ${function:Unprotect-Secret} -GetText ${function:T})
  }

  function Sync-AuthTypeFromSelection {
    if ($rbUserPass -and $rbUserPass.Checked) {
      $script:Settings.authType = "userpass"
    } elseif ($rbApiKey -and $rbApiKey.Checked) {
      $script:Settings.authType = "apikey"
    }
  }

  function Sync-AuthTypeFromSelection {
    if ($rbUserPass -and $rbUserPass.Checked) {
      $script:Settings.authType = "userpass"
    } elseif ($rbApiKey -and $rbApiKey.Checked) {
      $script:Settings.authType = "apikey"
    }
  }

  function Invoke-SnowRequest {
    param(
      [Parameter(Mandatory=$true)][ValidateSet('Get','Post','Patch','Delete')][string]$Method,
      [Parameter(Mandatory=$true)][string]$Path,
      [AllowNull()]$Body,
      [int]$TimeoutSec = 120
    )

    Sync-AuthTypeFromSelection

    $params = @{
      Method = $Method
      Path = $Path
      Settings = $script:Settings
      UnprotectSecret = ${function:Unprotect-Secret}
      GetText = ${function:T}
      TimeoutSec = $TimeoutSec
    }
    if ($PSBoundParameters.ContainsKey('Body')) { $params.Body = $Body }

    return Invoke-CoreSnowRequest @params
  }

  function Invoke-SnowGet([string]$pathAndQuery) {
    return Invoke-SnowRequest -Method Get -Path $pathAndQuery
  }

  function Invoke-SnowPost([string]$path, [hashtable]$body) {
    return Invoke-SnowRequest -Method Post -Path $path -Body $body
  }

  function Invoke-SnowPatch([string]$path, [hashtable]$body) {
    return Invoke-SnowRequest -Method Patch -Path $path -Body $body
  }

  function Invoke-SnowDelete([string]$path) {
    return Invoke-SnowRequest -Method Delete -Path $path
  }

  function Invoke-SnowDownloadAttachmentBytes([string]$attachmentSysId) {
    Sync-AuthTypeFromSelection

    $base = Get-BaseUrl
    if ([string]::IsNullOrWhiteSpace($base)) { throw (T "WarnInstance") }
    $uri = "{0}/api/now/attachment/{1}/file" -f $base, $attachmentSysId

    if ((Resolve-CoreSnowAuthType -AuthType $script:Settings.authType) -eq "userpass") {
      $user = ([string]$script:Settings.userId).Trim()
      $pass = Unprotect-Secret ([string]$script:Settings.passwordEnc)
      if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) { throw (T "WarnAuth") }

      $raw = "{0}:{1}" -f $user, $pass
      $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
      $headers = @{ Authorization = ("Basic {0}" -f [System.Convert]::ToBase64String($bytes)) }
      $response = Invoke-WebRequest -Uri $uri -Method Get -Headers $headers -UseBasicParsing
    } else {
      $key = Unprotect-Secret ([string]$script:Settings.apiKeyEnc)
      if ([string]::IsNullOrWhiteSpace($key)) { throw (T "WarnAuth") }

      $headers = @{ "x-sn-apikey" = $key }
      $response = Invoke-WebRequest -Uri $uri -Method Get -Headers $headers -UseBasicParsing
    }

    $memory = New-Object System.IO.MemoryStream
    try {
      $response.RawContentStream.CopyTo($memory)
      return $memory.ToArray()
    } finally {
      $memory.Dispose()
    }
  }

  function Invoke-SnowBatchDelete([string]$table, [string[]]$sysIds) {
    return (Invoke-CoreSnowBatchDelete -Table $table -SysIds $sysIds -InvokePost ${function:Invoke-SnowPost})
  }

  function New-VerificationCode([int]$length = 4) {
    if ($length -lt 1) { $length = 4 }
    $chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    $sb = New-Object System.Text.StringBuilder
    $rng = [System.Random]::new()
    for ($i = 0; $i -lt $length; $i++) {
      $n = $rng.Next(0, $chars.Length)
      [void]$sb.Append($chars[$n])
    }
    return $sb.ToString()
  }


  # ----------------------------
  # UI helpers
  # ----------------------------
  function Add-Log([string]$msg) {
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$ts] $msg"

    try {
      $logDir = Ensure-LogDir ([string]$script:Settings.logOutputDirectory)
      $script:Settings.logOutputDirectory = $logDir
      $logPath = Join-Path $logDir ("log_ps1snowutil_{0}.log" -f (Get-Date).ToString('yyyyMMdd'))
      [System.IO.File]::AppendAllText($logPath, $line + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
      # ignore file log failure
    }

    if (-not $script:txtLog) { return }
    Write-CoreUiLog -LogTextBox $script:txtLog -Message $msg
  }

  function Scroll-LogsToBottom {
    if (-not $script:txtLog) { return }
    $script:txtLog.SelectionStart = $script:txtLog.TextLength
    $script:txtLog.ScrollToCaret()
  }

  function Add-AttachmentLog([string]$msg) {
    Add-Log $msg
  }

  function Invoke-Async([string]$name, [scriptblock]$work, [scriptblock]$onCompleted, $state = $null) {
    Add-Log ("Running task: {0}" -f $name)
    try {
      $result = & $work $state
      & $onCompleted $result
    } catch {
      $errorMessage = if ($_ -is [System.Management.Automation.ErrorRecord]) {
        $_.Exception.Message
      } elseif ($_.PSObject.Properties.Name -contains "Message") {
        [string]$_.Message
      } else {
        [string]$_
      }
      Add-Log ("{0}: {1}" -f (T "Failed"), $errorMessage)
    }
  }

  function Ensure-ExportDir([string]$dir) {
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = $DefaultExportDir }
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    return $dir
  }

  function Ensure-LogDir([string]$dir) {
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = $DefaultLogDir }
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    return $dir
  }



  function New-UiCardPanel {
    param([int]$PaddingAll = 16)
    $card = New-Object System.Windows.Forms.Panel
    $card.Padding = New-Object System.Windows.Forms.Padding($PaddingAll)
    $card.Margin = New-Object System.Windows.Forms.Padding(8)
    $card.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    return $card
  }

  function New-UiSectionTitle([string]$text) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.AutoSize = $true
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lbl.Margin = New-Object System.Windows.Forms.Padding(0,0,0,8)
    return $lbl
  }

  function Set-Theme([string]$theme) {
    $resolved = ([string]$theme).Trim().ToLowerInvariant()
    if (@('light','dark') -notcontains $resolved) { $resolved = 'dark' }
    $script:ThemeName = $resolved
    if ($resolved -eq 'light') {
      $script:ThemePalette = @{
        Back = [System.Drawing.Color]::FromArgb(245,247,250)
        Surface = [System.Drawing.Color]::White
        Text = [System.Drawing.Color]::FromArgb(28,28,28)
        Muted = [System.Drawing.Color]::FromArgb(92,92,92)
        Accent = [System.Drawing.Color]::FromArgb(0,120,212)
        AccentText = [System.Drawing.Color]::White
        Danger = [System.Drawing.Color]::FromArgb(196,43,28)
        DangerText = [System.Drawing.Color]::White
        Border = [System.Drawing.Color]::FromArgb(220,223,230)
      }
    } else {
      $script:ThemePalette = @{
        Back = [System.Drawing.Color]::FromArgb(24,28,36)
        Surface = [System.Drawing.Color]::FromArgb(36,41,52)
        Text = [System.Drawing.Color]::FromArgb(240,242,246)
        Muted = [System.Drawing.Color]::FromArgb(172,178,190)
        Accent = [System.Drawing.Color]::FromArgb(74,158,255)
        AccentText = [System.Drawing.Color]::White
        Danger = [System.Drawing.Color]::FromArgb(224,86,73)
        DangerText = [System.Drawing.Color]::White
        Border = [System.Drawing.Color]::FromArgb(62,68,84)
      }
    }
  }

  function Apply-ThemeRecursive([System.Windows.Forms.Control]$control) {
    if ($null -eq $control) { return }
    $palette = $script:ThemePalette
    $control.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    if ($control -is [System.Windows.Forms.Form]) {
      $control.BackColor = $palette.Back
      $control.ForeColor = $palette.Text
    } elseif ($control -is [System.Windows.Forms.TabControl]) {
      $control.BackColor = $palette.Back
      $control.ForeColor = $palette.Text
    } elseif ($control -is [System.Windows.Forms.TabPage]) {
      $control.BackColor = $palette.Back
      $control.ForeColor = $palette.Text
    } elseif ($control -is [System.Windows.Forms.Panel] -or $control -is [System.Windows.Forms.GroupBox]) {
      $control.BackColor = $palette.Surface
      $control.ForeColor = $palette.Text
    } elseif ($control -is [System.Windows.Forms.Label] -or $control -is [System.Windows.Forms.LinkLabel] -or $control -is [System.Windows.Forms.RadioButton] -or $control -is [System.Windows.Forms.CheckBox]) {
      $control.BackColor = [System.Drawing.Color]::Transparent
      $control.ForeColor = $palette.Text
    } elseif ($control -is [System.Windows.Forms.TextBox] -or $control -is [System.Windows.Forms.ComboBox] -or $control -is [System.Windows.Forms.NumericUpDown] -or $control -is [System.Windows.Forms.DateTimePicker] -or $control -is [System.Windows.Forms.ListBox]) {
      $control.BackColor = $palette.Surface
      $control.ForeColor = $palette.Text
    } elseif ($control -is [System.Windows.Forms.Button]) {
      $control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
      $control.FlatAppearance.BorderSize = 1
      $control.FlatAppearance.BorderColor = $palette.Border
      $control.BackColor = $palette.Surface
      $control.ForeColor = $palette.Text
    } elseif ($control -is [System.Windows.Forms.DataGridView]) {
      $control.BackgroundColor = $palette.Surface
      $control.GridColor = $palette.Border
      $control.DefaultCellStyle.BackColor = $palette.Surface
      $control.DefaultCellStyle.ForeColor = $palette.Text
      $control.DefaultCellStyle.SelectionBackColor = $palette.Accent
      $control.DefaultCellStyle.SelectionForeColor = $palette.AccentText
      $control.ColumnHeadersDefaultCellStyle.BackColor = $palette.Surface
      $control.ColumnHeadersDefaultCellStyle.ForeColor = $palette.Text
      $control.EnableHeadersVisualStyles = $false
    }

    foreach ($child in $control.Controls) { Apply-ThemeRecursive $child }
  }

  function Set-ButtonStyle([System.Windows.Forms.Button]$button, [string]$kind) {
    if (-not $button) { return }
    $palette = $script:ThemePalette
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 1
    $resolvedKind = ''
    if ($null -ne $kind) { $resolvedKind = ([string]$kind).ToLowerInvariant() }
    switch ($resolvedKind) {
      'primary' {
        $button.BackColor = $palette.Accent
        $button.ForeColor = $palette.AccentText
        $button.FlatAppearance.BorderColor = $palette.Accent
      }
      'danger' {
        $button.BackColor = $palette.Danger
        $button.ForeColor = $palette.DangerText
        $button.FlatAppearance.BorderColor = $palette.Danger
      }
      default {
        $button.BackColor = $palette.Surface
        $button.ForeColor = $palette.Text
        $button.FlatAppearance.BorderColor = $palette.Border
      }
    }
  }

  # ----------------------------
  # Build GUI
  # ----------------------------
  $form = New-Object System.Windows.Forms.Form
  $form.StartPosition = "CenterScreen"
  $form.Size = New-Object System.Drawing.Size(1120, 720)
  $form.MinimumSize = New-Object System.Drawing.Size(1040, 650)
  $form.Padding = New-Object System.Windows.Forms.Padding(8)
  $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

  $tabs = New-Object System.Windows.Forms.TabControl
  Set-Theme ([string]$script:Settings.uiTheme)
  $tabs.Dock = "Fill"

  $tabExport = New-Object System.Windows.Forms.TabPage
  $tabAttachmentHarvester = New-Object System.Windows.Forms.TabPage
  $tabViewEditor = New-Object System.Windows.Forms.TabPage
  $tabSettings = New-Object System.Windows.Forms.TabPage
  $tabDelete = New-Object System.Windows.Forms.TabPage
  $tabLogs = New-Object System.Windows.Forms.TabPage

  if ($script:IsExportFeatureEnabled) { [void]$tabs.TabPages.Add($tabExport) }
  if ($script:IsAttachmentHarvesterFeatureEnabled) { [void]$tabs.TabPages.Add($tabAttachmentHarvester) }
  if ($script:IsViewEditorFeatureEnabled) { [void]$tabs.TabPages.Add($tabViewEditor) }
  if ($script:IsDeleteFeatureEnabled) { [void]$tabs.TabPages.Add($tabDelete) }
  [void]$tabs.TabPages.Add($tabLogs)
  [void]$tabs.TabPages.Add($tabSettings)
  $form.Controls.Add($tabs)

  # --- Export tab layout
  $panelExport = New-Object System.Windows.Forms.Panel
  $panelExport.Dock = "Fill"
  $panelExport.AutoScroll = $true
  $panelExport.AutoScrollMinSize = New-Object System.Drawing.Size(940, 660)
  $tabExport.Controls.Add($panelExport)

  $lblTable = New-Object System.Windows.Forms.Label
  $lblTable.Location = New-Object System.Drawing.Point(20, 20)
  $lblTable.AutoSize = $true

  $cmbTable = New-Object System.Windows.Forms.ComboBox
  $cmbTable.Location = New-Object System.Drawing.Point(160, 16)
  $cmbTable.Size = New-Object System.Drawing.Size(560, 28)
  $cmbTable.DropDownStyle = "DropDown"

  $btnReloadTables = New-Object System.Windows.Forms.Button
  $btnReloadTables.Location = New-Object System.Drawing.Point(740, 14)
  $btnReloadTables.Size = New-Object System.Drawing.Size(180, 32)

  $lblFilter = New-Object System.Windows.Forms.Label
  $lblFilter.Location = New-Object System.Drawing.Point(20, 65)
  $lblFilter.AutoSize = $true

  $rbAll = New-Object System.Windows.Forms.RadioButton
  $rbAll.Location = New-Object System.Drawing.Point(160, 63)
  $rbAll.AutoSize = $true

  $rbBetween = New-Object System.Windows.Forms.RadioButton
  $rbBetween.Location = New-Object System.Drawing.Point(240, 63)
  $rbBetween.AutoSize = $true

  $lblStart = New-Object System.Windows.Forms.Label
  $lblStart.Location = New-Object System.Drawing.Point(160, 105)
  $lblStart.AutoSize = $true

  $dtStart = New-Object System.Windows.Forms.DateTimePicker
  $dtStart.Location = New-Object System.Drawing.Point(210, 102)
  $dtStart.Size = New-Object System.Drawing.Size(250, 28)
  $dtStart.Format = "Custom"
  $dtStart.CustomFormat = "yyyy-MM-dd HH:mm:ss"
  $dtStart.ShowUpDown = $true

  $lblEnd = New-Object System.Windows.Forms.Label
  $lblEnd.Location = New-Object System.Drawing.Point(480, 105)
  $lblEnd.AutoSize = $true

  $dtEnd = New-Object System.Windows.Forms.DateTimePicker
  $dtEnd.Location = New-Object System.Drawing.Point(525, 102)
  $dtEnd.Size = New-Object System.Drawing.Size(200, 28)
  $dtEnd.Format = "Custom"
  $dtEnd.CustomFormat = "yyyy-MM-dd HH:mm:ss"
  $dtEnd.ShowUpDown = $true

  $btnLast30Days = New-Object System.Windows.Forms.Button
  $btnLast30Days.Location = New-Object System.Drawing.Point(740, 100)
  $btnLast30Days.Size = New-Object System.Drawing.Size(180, 32)

  $lblDir = New-Object System.Windows.Forms.Label
  $lblDir.Location = New-Object System.Drawing.Point(20, 150)
  $lblDir.AutoSize = $true

  $txtDir = New-Object System.Windows.Forms.TextBox
  $txtDir.Location = New-Object System.Drawing.Point(160, 146)
  $txtDir.Size = New-Object System.Drawing.Size(560, 28)

  $btnBrowse = New-Object System.Windows.Forms.Button
  $btnBrowse.Location = New-Object System.Drawing.Point(740, 144)
  $btnBrowse.Size = New-Object System.Drawing.Size(180, 32)

  $lblExportMaxRows = New-Object System.Windows.Forms.Label
  $lblExportMaxRows.Location = New-Object System.Drawing.Point(20, 194)
  $lblExportMaxRows.AutoSize = $true

  $numExportMaxRows = New-Object System.Windows.Forms.NumericUpDown
  $numExportMaxRows.Location = New-Object System.Drawing.Point(160, 190)
  $numExportMaxRows.Size = New-Object System.Drawing.Size(170, 28)
  $numExportMaxRows.Minimum = 1
  $numExportMaxRows.Maximum = 1000000
  $numExportMaxRows.Value = 10000

  $lblExportMaxRowsHint = New-Object System.Windows.Forms.Label
  $lblExportMaxRowsHint.Location = New-Object System.Drawing.Point(340, 194)
  $lblExportMaxRowsHint.Size = New-Object System.Drawing.Size(580, 72)
  $lblExportMaxRowsHint.ForeColor = [System.Drawing.Color]::FromArgb(90,90,90)

  $lblOutputFormat = New-Object System.Windows.Forms.Label
  $lblOutputFormat.Location = New-Object System.Drawing.Point(20, 282)
  $lblOutputFormat.AutoSize = $true

  $cmbOutputFormat = New-Object System.Windows.Forms.ComboBox
  $cmbOutputFormat.Location = New-Object System.Drawing.Point(160, 278)
  $cmbOutputFormat.Size = New-Object System.Drawing.Size(220, 28)
  $cmbOutputFormat.DropDownStyle = "DropDownList"
  [void]$cmbOutputFormat.Items.Add("csv")
  [void]$cmbOutputFormat.Items.Add("json")
  [void]$cmbOutputFormat.Items.Add("xlsx")

  $chkOutputBom = New-Object System.Windows.Forms.CheckBox
  $chkOutputBom.Location = New-Object System.Drawing.Point(390, 280)
  $chkOutputBom.AutoSize = $true

  $btnExecute = New-Object System.Windows.Forms.Button
  $btnExecute.Location = New-Object System.Drawing.Point(740, 270)
  $btnExecute.Size = New-Object System.Drawing.Size(180, 42)

  $btnOpenFolder = New-Object System.Windows.Forms.Button
  $btnOpenFolder.Location = New-Object System.Drawing.Point(740, 318)
  $btnOpenFolder.Size = New-Object System.Drawing.Size(180, 42)

  $panelExport.Controls.AddRange(@(
    $lblTable, $cmbTable, $btnReloadTables,
    $lblFilter, $rbAll, $rbBetween,
    $lblStart, $dtStart, $lblEnd, $dtEnd, $btnLast30Days,
    $lblDir, $txtDir, $btnBrowse,
    $lblExportMaxRows, $numExportMaxRows, $lblExportMaxRowsHint,
    $lblOutputFormat, $cmbOutputFormat, $chkOutputBom,
    $btnOpenFolder, $btnExecute
  ))


  # --- Attachment Harvester tab layout
  $panelAttachment = New-Object System.Windows.Forms.Panel
  $panelAttachment.Dock = "Fill"
  $panelAttachment.AutoScroll = $true
  $panelAttachment.AutoScrollMinSize = New-Object System.Drawing.Size(940, 660)
  $tabAttachmentHarvester.Controls.Add($panelAttachment)

  $lblAttachmentTable = New-Object System.Windows.Forms.Label
  $lblAttachmentTable.Location = New-Object System.Drawing.Point(20, 20)
  $lblAttachmentTable.AutoSize = $true

  $cmbAttachmentTable = New-Object System.Windows.Forms.ComboBox
  $cmbAttachmentTable.Location = New-Object System.Drawing.Point(220, 16)
  $cmbAttachmentTable.Size = New-Object System.Drawing.Size(500, 28)
  $cmbAttachmentTable.DropDownStyle = "DropDown"

  $cmbAttachmentDateField = New-Object System.Windows.Forms.ComboBox
  $cmbAttachmentDateField.Location = New-Object System.Drawing.Point(740, 16)
  $cmbAttachmentDateField.Size = New-Object System.Drawing.Size(180, 28)
  $cmbAttachmentDateField.DropDownStyle = "DropDownList"
  [void]$cmbAttachmentDateField.Items.Add('sys_created_on')
  [void]$cmbAttachmentDateField.Items.Add('sys_updated_on')

  $lblAttachmentStart = New-Object System.Windows.Forms.Label
  $lblAttachmentStart.Location = New-Object System.Drawing.Point(20, 65)
  $lblAttachmentStart.AutoSize = $true

  $dtAttachmentStart = New-Object System.Windows.Forms.DateTimePicker
  $dtAttachmentStart.Location = New-Object System.Drawing.Point(220, 61)
  $dtAttachmentStart.Size = New-Object System.Drawing.Size(250, 28)
  $dtAttachmentStart.Format = "Custom"
  $dtAttachmentStart.CustomFormat = "yyyy-MM-dd HH:mm:ss"
  $dtAttachmentStart.ShowUpDown = $true

  $lblAttachmentEnd = New-Object System.Windows.Forms.Label
  $lblAttachmentEnd.Location = New-Object System.Drawing.Point(490, 65)
  $lblAttachmentEnd.AutoSize = $true

  $dtAttachmentEnd = New-Object System.Windows.Forms.DateTimePicker
  $dtAttachmentEnd.Location = New-Object System.Drawing.Point(560, 61)
  $dtAttachmentEnd.Size = New-Object System.Drawing.Size(160, 28)
  $dtAttachmentEnd.Format = "Custom"
  $dtAttachmentEnd.CustomFormat = "yyyy-MM-dd HH:mm:ss"
  $dtAttachmentEnd.ShowUpDown = $true

  $btnAttachmentLastRunToNow = New-Object System.Windows.Forms.Button
  $btnAttachmentLastRunToNow.Location = New-Object System.Drawing.Point(740, 59)
  $btnAttachmentLastRunToNow.Size = New-Object System.Drawing.Size(180, 32)

  $lblAttachmentDir = New-Object System.Windows.Forms.Label
  $lblAttachmentDir.Location = New-Object System.Drawing.Point(20, 108)
  $lblAttachmentDir.AutoSize = $true

  $txtAttachmentDir = New-Object System.Windows.Forms.TextBox
  $txtAttachmentDir.Location = New-Object System.Drawing.Point(220, 104)
  $txtAttachmentDir.Size = New-Object System.Drawing.Size(500, 28)

  $btnAttachmentBrowse = New-Object System.Windows.Forms.Button
  $btnAttachmentBrowse.Location = New-Object System.Drawing.Point(740, 102)
  $btnAttachmentBrowse.Size = New-Object System.Drawing.Size(180, 32)

  $chkAttachmentSubfolder = New-Object System.Windows.Forms.CheckBox
  $chkAttachmentSubfolder.Location = New-Object System.Drawing.Point(220, 145)
  $chkAttachmentSubfolder.AutoSize = $true

  $btnAttachmentExecute = New-Object System.Windows.Forms.Button
  $btnAttachmentExecute.Location = New-Object System.Drawing.Point(740, 145)
  $btnAttachmentExecute.Size = New-Object System.Drawing.Size(180, 42)

  $panelAttachment.Controls.AddRange(@(
    $lblAttachmentTable, $cmbAttachmentTable, $cmbAttachmentDateField,
    $lblAttachmentStart, $dtAttachmentStart, $lblAttachmentEnd, $dtAttachmentEnd, $btnAttachmentLastRunToNow,
    $lblAttachmentDir, $txtAttachmentDir, $btnAttachmentBrowse,
    $chkAttachmentSubfolder,
    $btnAttachmentExecute
  ))

  # --- Logs tab layout
  $panelLogs = New-Object System.Windows.Forms.Panel
  $panelLogs.Dock = "Fill"
  $panelLogs.Padding = New-Object System.Windows.Forms.Padding(12)
  $panelLogs.AutoScroll = $true
  $panelLogs.AutoScrollMinSize = New-Object System.Drawing.Size(940, 600)
  $tabLogs.Controls.Add($panelLogs)

  $lblLogDir = New-Object System.Windows.Forms.Label
  $lblLogDir.Location = New-Object System.Drawing.Point(20, 20)
  $lblLogDir.AutoSize = $true

  $txtLogDir = New-Object System.Windows.Forms.TextBox
  $txtLogDir.Location = New-Object System.Drawing.Point(220, 16)
  $txtLogDir.Size = New-Object System.Drawing.Size(500, 28)

  $btnLogBrowse = New-Object System.Windows.Forms.Button
  $btnLogBrowse.Location = New-Object System.Drawing.Point(740, 14)
  $btnLogBrowse.Size = New-Object System.Drawing.Size(180, 32)

  $lblLogSearch = New-Object System.Windows.Forms.Label
  $lblLogSearch.Location = New-Object System.Drawing.Point(20, 62)
  $lblLogSearch.AutoSize = $true

  $txtLogSearch = New-Object System.Windows.Forms.TextBox
  $txtLogSearch.Location = New-Object System.Drawing.Point(80, 58)
  $txtLogSearch.Size = New-Object System.Drawing.Size(180, 28)

  $btnLogCopy = New-Object System.Windows.Forms.Button
  $btnLogCopy.Location = New-Object System.Drawing.Point(270, 56)
  $btnLogCopy.Size = New-Object System.Drawing.Size(100, 32)

  $btnLogClear = New-Object System.Windows.Forms.Button
  $btnLogClear.Location = New-Object System.Drawing.Point(378, 56)
  $btnLogClear.Size = New-Object System.Drawing.Size(100, 32)

  $chkLogAutoScroll = New-Object System.Windows.Forms.CheckBox
  $chkLogAutoScroll.Location = New-Object System.Drawing.Point(490, 61)
  $chkLogAutoScroll.AutoSize = $true
  $chkLogAutoScroll.Checked = $true

  $script:txtLog = New-Object System.Windows.Forms.TextBox
  $script:txtLog.Multiline = $true
  $script:txtLog.ScrollBars = "Both"
  $script:txtLog.WordWrap = $false
  $script:txtLog.Location = New-Object System.Drawing.Point(20, 96)
  $script:txtLog.Size = New-Object System.Drawing.Size(900, 492)
  $script:txtLog.ReadOnly = $true
  $script:txtLog.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

  $txtLogDir.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
  $btnLogBrowse.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

  $panelLogs.Controls.AddRange(@(
    $lblLogDir, $txtLogDir, $btnLogBrowse,
    $lblLogSearch, $txtLogSearch, $btnLogCopy, $btnLogClear, $chkLogAutoScroll,
    $script:txtLog
  ))

  # --- DataBase View Editor tab layout
  $panelViewEditor = New-Object System.Windows.Forms.Panel
  $panelViewEditor.Dock = "Fill"
  $panelViewEditor.AutoScroll = $true
  $panelViewEditor.AutoScrollMinSize = New-Object System.Drawing.Size(940, 560)
  $tabViewEditor.Controls.Add($panelViewEditor)

  $lblViewName = New-Object System.Windows.Forms.Label
  $lblViewName.Location = New-Object System.Drawing.Point(20, 20)
  $lblViewName.AutoSize = $true

  $txtViewName = New-Object System.Windows.Forms.TextBox
  $txtViewName.Location = New-Object System.Drawing.Point(190, 16)
  $txtViewName.Size = New-Object System.Drawing.Size(330, 28)

  $lblViewLabel = New-Object System.Windows.Forms.Label
  $lblViewLabel.Location = New-Object System.Drawing.Point(540, 20)
  $lblViewLabel.AutoSize = $true

  $txtViewLabel = New-Object System.Windows.Forms.TextBox
  $txtViewLabel.Location = New-Object System.Drawing.Point(650, 16)
  $txtViewLabel.Size = New-Object System.Drawing.Size(270, 28)

  $lblBaseTable = New-Object System.Windows.Forms.Label
  $lblBaseTable.Location = New-Object System.Drawing.Point(20, 60)
  $lblBaseTable.AutoSize = $true

  $cmbBaseTable = New-Object System.Windows.Forms.ComboBox
  $cmbBaseTable.Location = New-Object System.Drawing.Point(190, 56)
  $cmbBaseTable.Size = New-Object System.Drawing.Size(520, 28)
  $cmbBaseTable.DropDownStyle = "DropDown"
  $cmbBaseTable.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
  $cmbBaseTable.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems

  $btnReloadColumns = New-Object System.Windows.Forms.Button
  $btnReloadColumns.Location = New-Object System.Drawing.Point(740, 54)
  $btnReloadColumns.Size = New-Object System.Drawing.Size(180, 32)

  $lblViewColumns = New-Object System.Windows.Forms.Label
  $lblViewColumns.Location = New-Object System.Drawing.Point(20, 100)
  $lblViewColumns.AutoSize = $true

  $clbViewColumns = New-Object System.Windows.Forms.ListBox
  $clbViewColumns.Location = New-Object System.Drawing.Point(190, 100)
  $clbViewColumns.Size = New-Object System.Drawing.Size(730, 120)
  $clbViewColumns.HorizontalScrollbar = $true

  $txtViewName.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
  $lblViewLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
  $txtViewLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
  $cmbBaseTable.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
  $btnReloadColumns.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
  $clbViewColumns.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

  $lblJoinDefinitions = New-Object System.Windows.Forms.Label
  $lblJoinDefinitions.Location = New-Object System.Drawing.Point(20, 230)
  $lblJoinDefinitions.AutoSize = $true

  $btnAddJoin = New-Object System.Windows.Forms.Button
  $btnAddJoin.Location = New-Object System.Drawing.Point(190, 226)
  $btnAddJoin.Size = New-Object System.Drawing.Size(170, 32)

  $btnRemoveJoin = New-Object System.Windows.Forms.Button
  $btnRemoveJoin.Location = New-Object System.Drawing.Point(370, 226)
  $btnRemoveJoin.Size = New-Object System.Drawing.Size(170, 32)

  $lblBasePrefix = New-Object System.Windows.Forms.Label
  $lblBasePrefix.Location = New-Object System.Drawing.Point(560, 232)
  $lblBasePrefix.AutoSize = $true

  $txtBasePrefix = New-Object System.Windows.Forms.TextBox
  $txtBasePrefix.Location = New-Object System.Drawing.Point(670, 228)
  $txtBasePrefix.Size = New-Object System.Drawing.Size(120, 28)
  $lblBasePrefix.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
  $txtBasePrefix.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

  $gridJoins = New-Object System.Windows.Forms.DataGridView
  $gridJoins.Location = New-Object System.Drawing.Point(190, 264)
  $gridJoins.Size = New-Object System.Drawing.Size(730, 220)
  $gridJoins.AllowUserToAddRows = $false
  $gridJoins.AllowUserToDeleteRows = $false
  $gridJoins.RowHeadersVisible = $false
  $gridJoins.SelectionMode = "FullRowSelect"
  $gridJoins.MultiSelect = $false
  $gridJoins.AutoSizeColumnsMode = "Fill"
  $gridJoins.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

  $colJoinTable = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
  $colJoinTable.Name = "JoinTable"
  $colJoinTable.FlatStyle = "Popup"
  $colJoinTable.DisplayStyle = "DropDownButton"
  $colJoinTable.FillWeight = 34

  $colJoinBaseColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
  $colJoinBaseColumn.Name = "JoinBaseColumn"
  $colJoinBaseColumn.FlatStyle = "Popup"
  $colJoinBaseColumn.DisplayStyle = "DropDownButton"
  $colJoinBaseColumn.FillWeight = 26

  $colJoinSource = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
  $colJoinSource.Name = "JoinSource"
  $colJoinSource.FlatStyle = "Popup"
  $colJoinSource.DisplayStyle = "DropDownButton"
  $colJoinSource.FillWeight = 20

  $colJoinTargetColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
  $colJoinTargetColumn.Name = "JoinTargetColumn"
  $colJoinTargetColumn.FlatStyle = "Popup"
  $colJoinTargetColumn.DisplayStyle = "DropDownButton"
  $colJoinTargetColumn.FillWeight = 20

  $colJoinPrefix = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colJoinPrefix.Name = "JoinPrefix"
  $colJoinPrefix.FillWeight = 14

  $colJoinLeftJoin = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
  $colJoinLeftJoin.Name = "LeftJoin"
  $colJoinLeftJoin.FillWeight = 10

  [void]$gridJoins.Columns.Add($colJoinTable)
  [void]$gridJoins.Columns.Add($colJoinSource)
  [void]$gridJoins.Columns.Add($colJoinBaseColumn)
  [void]$gridJoins.Columns.Add($colJoinTargetColumn)
  [void]$gridJoins.Columns.Add($colJoinPrefix)
  [void]$gridJoins.Columns.Add($colJoinLeftJoin)


  $btnCreateView = New-Object System.Windows.Forms.Button
  $btnCreateView.Location = New-Object System.Drawing.Point(740, 500)
  $btnCreateView.Size = New-Object System.Drawing.Size(180, 42)
  $btnCreateView.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right

  $lnkCreatedViewList = New-Object System.Windows.Forms.LinkLabel
  $lnkCreatedViewList.Location = New-Object System.Drawing.Point(190, 504)
  $lnkCreatedViewList.Size = New-Object System.Drawing.Size(540, 18)
  $lnkCreatedViewList.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
  $lnkCreatedViewList.Visible = $false
  $lnkCreatedViewList.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline

  $lnkCreatedViewDefinition = New-Object System.Windows.Forms.LinkLabel
  $lnkCreatedViewDefinition.Location = New-Object System.Drawing.Point(190, 526)
  $lnkCreatedViewDefinition.Size = New-Object System.Drawing.Size(540, 18)
  $lnkCreatedViewDefinition.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
  $lnkCreatedViewDefinition.Visible = $false
  $lnkCreatedViewDefinition.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline

  $panelViewEditor.Controls.AddRange(@(
    $lblViewName, $txtViewName,
    $lblViewLabel, $txtViewLabel,
    $lblBaseTable, $cmbBaseTable, $btnReloadColumns,
    $lblViewColumns, $clbViewColumns,
    $lblJoinDefinitions, $btnAddJoin, $btnRemoveJoin, $lblBasePrefix, $txtBasePrefix,
    $gridJoins,
    $lnkCreatedViewList, $lnkCreatedViewDefinition,
    $btnCreateView
  ))

  # --- Settings tab layout
  $panelSettings = New-Object System.Windows.Forms.Panel
  $panelSettings.Dock = "Fill"
  $panelSettings.AutoScroll = $true
  $panelSettings.AutoScrollMinSize = New-Object System.Drawing.Size(940, 420)
  $tabSettings.Controls.Add($panelSettings)

  $lblUiLang = New-Object System.Windows.Forms.Label
  $lblUiLang.Location = New-Object System.Drawing.Point(20, 20)
  $lblUiLang.AutoSize = $true

  $cmbLang = New-Object System.Windows.Forms.ComboBox
  $cmbLang.Location = New-Object System.Drawing.Point(220, 16)
  $cmbLang.Size = New-Object System.Drawing.Size(160, 28)
  $cmbLang.DropDownStyle = "DropDownList"
  [void]$cmbLang.Items.Add("ja")
  [void]$cmbLang.Items.Add("en")

  $lblTheme = New-Object System.Windows.Forms.Label
  $lblTheme.Location = New-Object System.Drawing.Point(420, 20)
  $lblTheme.AutoSize = $true

  $cmbTheme = New-Object System.Windows.Forms.ComboBox
  $cmbTheme.Location = New-Object System.Drawing.Point(560, 16)
  $cmbTheme.Size = New-Object System.Drawing.Size(160, 28)
  $cmbTheme.DropDownStyle = "DropDownList"
  [void]$cmbTheme.Items.Add("dark")
  [void]$cmbTheme.Items.Add("light")

  $lblInstance = New-Object System.Windows.Forms.Label
  $lblInstance.Location = New-Object System.Drawing.Point(20, 60)
  $lblInstance.AutoSize = $true

  $txtInstance = New-Object System.Windows.Forms.TextBox
  $txtInstance.Location = New-Object System.Drawing.Point(220, 56)
  $txtInstance.Size = New-Object System.Drawing.Size(500, 28)

  $lblBaseUrl = New-Object System.Windows.Forms.Label
  $lblBaseUrl.Location = New-Object System.Drawing.Point(220, 88)
  $lblBaseUrl.Size = New-Object System.Drawing.Size(700, 18)
  $lblBaseUrl.ForeColor = [System.Drawing.Color]::FromArgb(70,70,70)

  $lblAuthType = New-Object System.Windows.Forms.Label
  $lblAuthType.Location = New-Object System.Drawing.Point(20, 125)
  $lblAuthType.AutoSize = $true

  $rbUserPass = New-Object System.Windows.Forms.RadioButton
  $rbUserPass.Location = New-Object System.Drawing.Point(220, 123)
  $rbUserPass.AutoSize = $true

  $rbApiKey = New-Object System.Windows.Forms.RadioButton
  $rbApiKey.Location = New-Object System.Drawing.Point(420, 123)
  $rbApiKey.AutoSize = $true

  $lblUser = New-Object System.Windows.Forms.Label
  $lblUser.Location = New-Object System.Drawing.Point(20, 170)
  $lblUser.AutoSize = $true

  $txtUser = New-Object System.Windows.Forms.TextBox
  $txtUser.Location = New-Object System.Drawing.Point(220, 166)
  $txtUser.Size = New-Object System.Drawing.Size(260, 28)

  $lblPass = New-Object System.Windows.Forms.Label
  $lblPass.Location = New-Object System.Drawing.Point(20, 210)
  $lblPass.AutoSize = $true

  $txtPass = New-Object System.Windows.Forms.TextBox
  $txtPass.Location = New-Object System.Drawing.Point(220, 206)
  $txtPass.Size = New-Object System.Drawing.Size(360, 28)
  $txtPass.UseSystemPasswordChar = $true

  $btnTogglePass = New-Object System.Windows.Forms.Button
  $btnTogglePass.Location = New-Object System.Drawing.Point(600, 204)
  $btnTogglePass.Size = New-Object System.Drawing.Size(120, 32)

  $lblKey = New-Object System.Windows.Forms.Label
  $lblKey.Location = New-Object System.Drawing.Point(20, 250)
  $lblKey.AutoSize = $true

  $txtKey = New-Object System.Windows.Forms.TextBox
  $txtKey.Location = New-Object System.Drawing.Point(220, 246)
  $txtKey.Size = New-Object System.Drawing.Size(360, 28)
  $txtKey.UseSystemPasswordChar = $true

  $btnToggleKey = New-Object System.Windows.Forms.Button
  $btnToggleKey.Location = New-Object System.Drawing.Point(600, 244)
  $btnToggleKey.Size = New-Object System.Drawing.Size(120, 32)

  $lblSaveHint = New-Object System.Windows.Forms.Label
  $lblSaveHint.Location = New-Object System.Drawing.Point(20, 305)
  $lblSaveHint.AutoSize = $true
  $lblSaveHint.ForeColor = [System.Drawing.Color]::FromArgb(70,70,70)

  $lblTablesHint = New-Object System.Windows.Forms.Label
  $lblTablesHint.Location = New-Object System.Drawing.Point(20, 335)
  $lblTablesHint.Size = New-Object System.Drawing.Size(900, 60)
  $lblTablesHint.ForeColor = [System.Drawing.Color]::FromArgb(70,70,70)

  $lnkCopyright = New-Object System.Windows.Forms.LinkLabel
  $lnkCopyright.Location = New-Object System.Drawing.Point(20, 0)
  $lnkCopyright.AutoSize = $true
  $lnkCopyright.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
  $lnkCopyright.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline

  function Position-CopyrightLink {
    $top = $panelSettings.ClientSize.Height - $lnkCopyright.Height - 16
    if ($top -lt 16) { $top = 16 }
    $lnkCopyright.Location = New-Object System.Drawing.Point(20, $top)
  }

  $panelSettings.Controls.AddRange(@(
    $lblUiLang, $cmbLang, $lblTheme, $cmbTheme,
    $lblInstance, $txtInstance, $lblBaseUrl,
    $lblAuthType, $rbUserPass, $rbApiKey,
    $lblUser, $txtUser,
    $lblPass, $txtPass, $btnTogglePass,
    $lblKey,  $txtKey,  $btnToggleKey,
    $lblSaveHint, $lblTablesHint,
    $lnkCopyright
  ))

  # --- Delete tab layout
  $panelDelete = New-Object System.Windows.Forms.Panel
  $panelDelete.Dock = "Fill"
  $panelDelete.AutoScroll = $true
  $panelDelete.AutoScrollMinSize = New-Object System.Drawing.Size(940, 420)
  $tabDelete.Controls.Add($panelDelete)

  $lblDeleteTable = New-Object System.Windows.Forms.Label
  $lblDeleteTable.Location = New-Object System.Drawing.Point(20, 20)
  $lblDeleteTable.AutoSize = $true

  $cmbDeleteTable = New-Object System.Windows.Forms.ComboBox
  $cmbDeleteTable.Location = New-Object System.Drawing.Point(220, 16)
  $cmbDeleteTable.Size = New-Object System.Drawing.Size(500, 28)
  $cmbDeleteTable.DropDownStyle = "DropDown"

  $btnDeleteReloadTables = New-Object System.Windows.Forms.Button
  $btnDeleteReloadTables.Location = New-Object System.Drawing.Point(740, 14)
  $btnDeleteReloadTables.Size = New-Object System.Drawing.Size(180, 32)

  $lblDeleteMaxRetries = New-Object System.Windows.Forms.Label
  $lblDeleteMaxRetries.Location = New-Object System.Drawing.Point(20, 65)
  $lblDeleteMaxRetries.AutoSize = $true

  $numDeleteMaxRetries = New-Object System.Windows.Forms.NumericUpDown
  $numDeleteMaxRetries.Location = New-Object System.Drawing.Point(220, 62)
  $numDeleteMaxRetries.Size = New-Object System.Drawing.Size(140, 28)
  $numDeleteMaxRetries.Minimum = 1
  $numDeleteMaxRetries.Maximum = 999
  $numDeleteMaxRetries.Value = 99

  $lblDeleteDangerHint = New-Object System.Windows.Forms.Label
  $lblDeleteDangerHint.Location = New-Object System.Drawing.Point(20, 105)
  $lblDeleteDangerHint.Size = New-Object System.Drawing.Size(900, 24)
  $lblDeleteDangerHint.ForeColor = [System.Drawing.Color]::FromArgb(180,30,30)

  $lblDeleteUsageHint = New-Object System.Windows.Forms.Label
  $lblDeleteUsageHint.Location = New-Object System.Drawing.Point(20, 130)
  $lblDeleteUsageHint.Size = New-Object System.Drawing.Size(900, 40)
  $lblDeleteUsageHint.ForeColor = [System.Drawing.Color]::FromArgb(140,70,30)

  $lblDeleteAllowedInstances = New-Object System.Windows.Forms.Label
  $lblDeleteAllowedInstances.Location = New-Object System.Drawing.Point(20, 175)
  $lblDeleteAllowedInstances.AutoSize = $true

  $txtDeleteAllowedInstances = New-Object System.Windows.Forms.TextBox
  $txtDeleteAllowedInstances.Location = New-Object System.Drawing.Point(220, 196)
  $txtDeleteAllowedInstances.Size = New-Object System.Drawing.Size(700, 28)
  $txtDeleteAllowedInstances.ReadOnly = $true

  $lblDeleteAllowedInstancesHint = New-Object System.Windows.Forms.Label
  $lblDeleteAllowedInstancesHint.Location = New-Object System.Drawing.Point(20, 230)
  $lblDeleteAllowedInstancesHint.Size = New-Object System.Drawing.Size(900, 32)
  $lblDeleteAllowedInstancesHint.ForeColor = [System.Drawing.Color]::FromArgb(110,70,70)

  $lblDeleteProgress = New-Object System.Windows.Forms.Label
  $lblDeleteProgress.Location = New-Object System.Drawing.Point(20, 278)
  $lblDeleteProgress.AutoSize = $true

  $prgDelete = New-Object System.Windows.Forms.ProgressBar
  $prgDelete.Location = New-Object System.Drawing.Point(220, 275)
  $prgDelete.Size = New-Object System.Drawing.Size(500, 24)
  $prgDelete.Minimum = 0
  $prgDelete.Maximum = 100
  $prgDelete.Value = 0

  $lblDeleteProgressValue = New-Object System.Windows.Forms.Label
  $lblDeleteProgressValue.Location = New-Object System.Drawing.Point(740, 278)
  $lblDeleteProgressValue.Size = New-Object System.Drawing.Size(180, 24)

  $btnDeleteExecute = New-Object System.Windows.Forms.Button
  $btnDeleteExecute.Location = New-Object System.Drawing.Point(740, 333)
  $btnDeleteExecute.Size = New-Object System.Drawing.Size(180, 42)
  $btnDeleteExecute.Enabled = $false

  $panelDelete.Controls.AddRange(@(
    $lblDeleteTable, $cmbDeleteTable, $btnDeleteReloadTables,
    $lblDeleteMaxRetries, $numDeleteMaxRetries,
    $lblDeleteDangerHint, $lblDeleteUsageHint,
    $lblDeleteAllowedInstances, $txtDeleteAllowedInstances, $lblDeleteAllowedInstancesHint,
    $lblDeleteProgress, $prgDelete, $lblDeleteProgressValue,
    $btnDeleteExecute
  ))

  function Apply-Language {
    $form.Text = T "AppTitle"
    if ($script:IsExportFeatureEnabled) { $tabExport.Text = T "TabExport" }
    if ($script:IsAttachmentHarvesterFeatureEnabled) { $tabAttachmentHarvester.Text = T "TabAttachmentHarvester" }
    if ($script:IsViewEditorFeatureEnabled) { $tabViewEditor.Text = T "TabViewEditor" }
    $tabLogs.Text = T "TabLogs"
    $tabSettings.Text = T "TabSettings"
    if ($script:IsDeleteFeatureEnabled) { $tabDelete.Text = T "TabDelete" }

    $lblTable.Text = T "TargetTable"
    $btnReloadTables.Text = T "ReloadTables"
    $lblFilter.Text = T "EasyFilter"
    $rbAll.Text = T "FilterAll"
    $rbBetween.Text = T "FilterUpdatedBetween"
    $lblStart.Text = T "Start"
    $lblEnd.Text = T "End"
    $btnLast30Days.Text = T "Last30Days"
    $lblDir.Text = T "ExportDir"
    $lblExportMaxRows.Text = T "ExportMaxRows"
    $lblExportMaxRowsHint.Text = T "ExportMaxRowsCsvHint"
    $btnBrowse.Text = T "Browse"
    $btnExecute.Text = T "Execute"
    $lblOutputFormat.Text = T "OutputFormat"
    $chkOutputBom.Text = T "OutputBom"
    $btnOpenFolder.Text = T "OpenFolder"

    $lblLogDir.Text = T "LogOutputDir"
    $btnLogBrowse.Text = T "Browse"
    $lblLogSearch.Text = T "LogSearch"
    $btnLogCopy.Text = T "LogCopy"
    $btnLogClear.Text = T "LogClear"
    $chkLogAutoScroll.Text = T "LogAutoScroll"

    $lblAttachmentTable.Text = T "TargetTable"
    $lblAttachmentStart.Text = T "Start"
    $lblAttachmentEnd.Text = T "End"
    $btnAttachmentLastRunToNow.Text = T "AttachmentSetLastRunToNow"
    $lblAttachmentDir.Text = T "AttachmentDownloadDir"
    $btnAttachmentBrowse.Text = T "Browse"
    $chkAttachmentSubfolder.Text = T "AttachmentCreateSubfolderPerTable"
    $btnAttachmentExecute.Text = T "Execute"

    $lblDeleteTable.Text = T "DeleteTargetTable"
    $btnDeleteReloadTables.Text = T "ReloadTables"
    $lblDeleteMaxRetries.Text = T "DeleteMaxRetries"
    $lblDeleteDangerHint.Text = (T "DeleteStep1") + "  " + (T "DeleteDangerHint")
    $lblDeleteUsageHint.Text = (T "DeleteStep2") + "  " + (T "DeleteUsageHint")
        $lblDeleteAllowedInstances.Text = T "DeleteAllowedInstances"
    $lblDeleteAllowedInstancesHint.Text = T "DeleteAllowedInstancesHint"
    $lblDeleteProgress.Text = T "DeleteProgress"
    $btnDeleteExecute.Text = T "DeleteExecute"

    $lblViewName.Text = T "ViewName"
    $lblViewLabel.Text = T "ViewLabel"
    $lblBaseTable.Text = T "BaseTable"
    $btnReloadColumns.Text = T "ReloadColumns"
    $lblViewColumns.Text = T "ViewColumns"
    $lblJoinDefinitions.Text = T "JoinDefinitions"
    $btnAddJoin.Text = T "AddJoin"
    $btnRemoveJoin.Text = T "RemoveJoin"
    $lblBasePrefix.Text = T "BasePrefix"
    $btnCreateView.Text = T "CreateView"
    $colJoinTable.HeaderText = T "JoinTable"
    $colJoinSource.HeaderText = T "JoinSource"
    $colJoinBaseColumn.HeaderText = T "JoinBaseColumn"
    $colJoinTargetColumn.HeaderText = T "JoinTargetColumn"
    $colJoinPrefix.HeaderText = T "JoinPrefix"
    $colJoinLeftJoin.HeaderText = T "LeftJoin"
    if ($lnkCreatedViewList.Visible) {
      $lnkCreatedViewList.Text = "{0}: {1}" -f (T "CreatedViewListLink"), [string]$lnkCreatedViewList.Tag
    }
    if ($lnkCreatedViewDefinition.Visible) {
      $lnkCreatedViewDefinition.Text = "{0}: {1}" -f (T "CreatedViewDefinitionLink"), [string]$lnkCreatedViewDefinition.Tag
    }

    $lblUiLang.Text = T "UiLang"
    $lblTheme.Text = T "UiTheme"
    $lblInstance.Text = T "Instance"
    Update-BaseUrlLabel
    $lblAuthType.Text = T "AuthType"
    $rbUserPass.Text = T "AuthUserPass"
    $rbApiKey.Text = T "AuthApiKey"
    $lblUser.Text = T "UserId"
    $lblPass.Text = T "Password"
    $lblKey.Text  = T "ApiKey"
    if ($txtPass.UseSystemPasswordChar) {
      $btnTogglePass.Text = T "Show"
    } else {
      $btnTogglePass.Text = T "Hide"
    }
    if ($txtKey.UseSystemPasswordChar) {
      $btnToggleKey.Text = T "Show"
    } else {
      $btnToggleKey.Text = T "Hide"
    }

    $lblSaveHint.Text = T "SaveHint"
    $lblTablesHint.Text = T "TestTablesHint"
    $lnkCopyright.Text = T "CopyrightLink"
    $lnkCopyright.Links.Clear()
    [void]$lnkCopyright.Links.Add(0, $lnkCopyright.Text.Length, "https://www.ixam.net")
    Position-CopyrightLink
  }

  $lnkCopyright.add_LinkClicked({
    param($sender, $e)
    $target = [string]$e.Link.LinkData
    if ([string]::IsNullOrWhiteSpace($target)) { $target = "https://www.ixam.net" }
    Start-Process $target | Out-Null
  })

  $panelSettings.add_Resize({ Position-CopyrightLink })

  function Update-BaseUrlLabel {
    $base = Get-BaseUrl
    if ([string]::IsNullOrWhiteSpace($base)) {
      $lblBaseUrl.Text = "{0}: {1}" -f (T "BaseUrlLabel"), (T "BaseUrlNotSet")
    } else {
      $lblBaseUrl.Text = "{0}: {1}" -f (T "BaseUrlLabel"), $base
    }
  }

  function Update-AuthUI {
    $isUserPass = $rbUserPass.Checked
    $txtUser.Enabled = $isUserPass
    $txtPass.Enabled = $isUserPass
    $btnTogglePass.Enabled = $isUserPass
    $txtKey.Enabled = (-not $isUserPass)
    $btnToggleKey.Enabled = (-not $isUserPass)
  }

  function Update-FilterUI {
    $isBetween = $rbBetween.Checked
    $dtStart.Enabled = $isBetween
    $dtEnd.Enabled   = $isBetween
  }

  function Complete-GridCurrentEdit([System.Windows.Forms.DataGridView]$grid, [string]$gridName) {
    if ($null -eq $grid -or -not $grid.IsCurrentCellDirty) { return }
    $currentCell = $grid.CurrentCell
    if ($currentCell -and $currentCell -is [System.Windows.Forms.DataGridViewTextBoxCell]) {
      return
    }
    try {
      $context = [System.Windows.Forms.DataGridViewDataErrorContexts]::Commit
      [void]$grid.CommitEdit($context)
      [void]$grid.EndEdit($context)
    } catch {
      Add-Log ("{0} grid edit commit failed: {1}" -f $gridName, $_.Exception.Message)
    }
  }

  # ----------------------------
  # Fetch table list from ServiceNow
  # ----------------------------
  function Fetch-Tables {
    Add-Log (T "FetchingTables")
    try {
      $fields = "name,label"
      $limit = 5000
      $q = "nameISNOTEMPTY^sys_update_nameISNOTEMPTY"
      $path = "/api/now/table/sys_db_object?sysparm_fields=$fields&sysparm_limit=$limit&sysparm_query=$(UrlEncode $q)"
      $res = Invoke-SnowGet $path
      $results = if ($res -and ($res.PSObject.Properties.Name -contains "result")) { @($res.result) } else { @() }
      $list = New-Object System.Collections.Generic.List[object]
      foreach ($r in $results) {
        $name = [string]$r.name
        $label = [string]$r.label
        if (-not [string]::IsNullOrWhiteSpace($name)) {
          if ([string]::IsNullOrWhiteSpace($label)) { $label = $name }
          [void]$list.Add([pscustomobject]@{ name=$name; label=$label })
        }
      }
      $list = @($list | Sort-Object name)

      $script:Settings.cachedTables = @($list)
      $script:Settings.cachedTablesFetchedAt = (Get-Date).ToString("o")
      Request-SaveSettings

      $cmbTable.BeginUpdate()
      $cmbTable.Items.Clear()
      foreach ($t in @($list)) {
        [void]$cmbTable.Items.Add(("{0} - {1}" -f $t.name, $t.label))
      }
      $cmbTable.EndUpdate()
      Refresh-BaseTableItems

      $targetName = ([string]$script:Settings.selectedTableName).Trim()
      if (-not [string]::IsNullOrWhiteSpace($targetName)) {
        $candidate = $null
        foreach ($item in $cmbTable.Items) {
          $itemText = [string]$item
          if ($itemText.StartsWith($targetName + " - ")) { $candidate = $item; break }
        }
        if ($candidate) { $cmbTable.SelectedItem = $candidate } else { $cmbTable.Text = $targetName }
      }

      Add-Log ("{0}: {1}" -f (T "Done"), @($list).Count)
    } catch {
      Add-Log ("{0}: {1}" -f (T "Failed"), $_.Exception.Message)
    }
  }

  function Ensure-TablesLoaded {
    $cachedCount = @($script:Settings.cachedTables).Count
    $uiCount = @($cmbTable.Items).Count
    if ($cachedCount -gt 0 -or $uiCount -gt 0) { return }
    Fetch-Tables
  }

  function Update-CreatedViewLinks([string]$viewName, [string]$viewSysId) {
    $base = Get-BaseUrl
    if ([string]::IsNullOrWhiteSpace($base) -or [string]::IsNullOrWhiteSpace($viewName) -or [string]::IsNullOrWhiteSpace($viewSysId)) {
      $lnkCreatedViewList.Visible = $false
      $lnkCreatedViewDefinition.Visible = $false
      return
    }

    $viewListUrl = "{0}/u_{1}_list.do" -f $base, $viewName
    $viewDefUrl = "{0}/sys_db_view.do?sys_id={1}" -f $base, $viewSysId

    $lnkCreatedViewList.Tag = $viewListUrl
    $lnkCreatedViewList.Text = "{0}: {1}" -f (T "CreatedViewListLink"), $viewListUrl
    $lnkCreatedViewList.Links.Clear()
    [void]$lnkCreatedViewList.Links.Add(0, $lnkCreatedViewList.Text.Length, $viewListUrl)
    $lnkCreatedViewList.Visible = $true

    $lnkCreatedViewDefinition.Tag = $viewDefUrl
    $lnkCreatedViewDefinition.Text = "{0}: {1}" -f (T "CreatedViewDefinitionLink"), $viewDefUrl
    $lnkCreatedViewDefinition.Links.Clear()
    [void]$lnkCreatedViewDefinition.Links.Add(0, $lnkCreatedViewDefinition.Text.Length, $viewDefUrl)
    $lnkCreatedViewDefinition.Visible = $true
  }

  function Get-SelectedTableName {
    $text = ""
    if ($cmbTable.SelectedItem) {
      $text = [string]$cmbTable.SelectedItem
    } else {
      $text = [string]$cmbTable.Text
    }
    $idx = $text.IndexOf(" - ")
    if ($idx -gt 0) { return $text.Substring(0, $idx).Trim() }
    return $text.Trim()
  }

  function Get-SelectedBaseTableName {
    $text = ""
    if ($cmbBaseTable.SelectedItem) {
      $text = [string]$cmbBaseTable.SelectedItem
    } else {
      $text = [string]$cmbBaseTable.Text
    }
    $idx = $text.IndexOf(" - ")
    if ($idx -gt 0) { return $text.Substring(0, $idx).Trim() }
    return $text.Trim()
  }

  function Get-SelectedDeleteTableName {
    $text = ""
    if ($cmbDeleteTable.SelectedItem) {
      $text = [string]$cmbDeleteTable.SelectedItem
    } else {
      $text = [string]$cmbDeleteTable.Text
    }
    $idx = $text.IndexOf(" - ")
    if ($idx -gt 0) { return $text.Substring(0, $idx).Trim() }
    return $text.Trim()
  }

  function Get-SelectedAttachmentTableName {
    $text = ""
    if ($cmbAttachmentTable.SelectedItem) {
      $text = [string]$cmbAttachmentTable.SelectedItem
    } else {
      $text = [string]$cmbAttachmentTable.Text
    }
    $idx = $text.IndexOf(" - ")
    if ($idx -gt 0) { return $text.Substring(0, $idx).Trim() }
    return $text.Trim()
  }

  function Convert-DisplayTokenToName([string]$text) {
    $value = ([string]$text).Trim()
    if ([string]::IsNullOrWhiteSpace($value)) { return "" }
    $idx = $value.IndexOf(" - ")
    if ($idx -gt 0) { return $value.Substring(0, $idx).Trim() }
    return $value
  }

  function Build-TableDisplayText([string]$tableName, [string]$tableLabel) {
    if ([string]::IsNullOrWhiteSpace([string]$tableLabel)) { return [string]$tableName }
    return ("{0} - {1}" -f [string]$tableName, [string]$tableLabel)
  }

  function Build-ColumnDisplayText([string]$columnName, [string]$columnLabel) {
    if ([string]::IsNullOrWhiteSpace([string]$columnLabel)) { return [string]$columnName }
    return ("{0} - {1}" -f [string]$columnName, [string]$columnLabel)
  }

  function Resolve-DisplayTextFromItems([System.Windows.Forms.DataGridViewComboBoxCell]$cell, [string]$token) {
    $name = Convert-DisplayTokenToName $token
    if ([string]::IsNullOrWhiteSpace($name) -or $null -eq $cell) { return "" }
    foreach ($item in $cell.Items) {
      $text = [string]$item
      if ([string]::IsNullOrWhiteSpace($text)) { continue }
      if ($text -eq $name -or $text.StartsWith($name + " - ")) { return $text }
    }
    return $name
  }

  function Get-TruncateAllowedInstancePatterns {
    $raw = [string]$script:Settings.truncateAllowedInstances
    if ([string]::IsNullOrWhiteSpace($raw)) { $raw = "*dev*,*stg*" }
    $parts = @($raw.Split(',') | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parts.Count -lt 1) { $parts = @("*dev*", "*stg*") }
    return $parts
  }

  function Test-TruncateInstanceAllowed {
    $baseUrl = ([string](Get-BaseUrl)).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($baseUrl)) { return $false }

    $candidates = New-Object System.Collections.Generic.List[string]
    $candidates.Add($baseUrl) | Out-Null

    try {
      $uri = [System.Uri]$baseUrl
      if ($uri -and -not [string]::IsNullOrWhiteSpace($uri.Host)) {
        $candidates.Add($uri.Host.ToLowerInvariant()) | Out-Null
      }
    } catch {
      # ignore parse errors and use baseUrl only
    }

    $patterns = @(Get-TruncateAllowedInstancePatterns)
    foreach ($pattern in $patterns) {
      $normalized = ([string]$pattern).ToLowerInvariant()
      foreach ($candidate in $candidates) {
        if ($candidate -like $normalized) { return $true }
      }
    }
    return $false
  }

  function Set-DeleteProgress([int]$percent, [string]$text) {
    if ($percent -lt 0) { $percent = 0 }
    if ($percent -gt 100) { $percent = 100 }
    $prgDelete.Value = $percent
    if ([string]::IsNullOrWhiteSpace($text)) {
      $lblDeleteProgressValue.Text = ("{0}%" -f $percent)
    } else {
      $lblDeleteProgressValue.Text = $text
    }
  }

  function Refresh-DeleteExecuteButton {
    $table = Get-SelectedDeleteTableName
    $isAllowed = Test-TruncateInstanceAllowed
    $btnDeleteExecute.Enabled = ((-not [string]::IsNullOrWhiteSpace($table)) -and $isAllowed)
  }

  function Request-DeleteVerificationCode {
    $code = New-VerificationCode 4

    $prompt = New-Object System.Windows.Forms.Form
    $prompt.StartPosition = "CenterParent"
    $prompt.Size = New-Object System.Drawing.Size(420, 230)
    $prompt.FormBorderStyle = "FixedDialog"
    $prompt.MaximizeBox = $false
    $prompt.MinimizeBox = $false
    $prompt.Text = T "DeletePromptTitle"

    $lblMessage = New-Object System.Windows.Forms.Label
    $lblMessage.Location = New-Object System.Drawing.Point(20, 20)
    $lblMessage.Size = New-Object System.Drawing.Size(360, 24)
    $lblMessage.Text = T "DeletePromptMessage"

    $lblCode = New-Object System.Windows.Forms.Label
    $lblCode.Location = New-Object System.Drawing.Point(20, 55)
    $lblCode.Size = New-Object System.Drawing.Size(360, 28)
    $lblCode.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblCode.Text = [string]::Format((T "DeletePromptCodeLabel"), $code)

    $lblInput = New-Object System.Windows.Forms.Label
    $lblInput.Location = New-Object System.Drawing.Point(20, 95)
    $lblInput.Size = New-Object System.Drawing.Size(100, 24)
    $lblInput.Text = T "DeletePromptInputLabel"

    $txtInput = New-Object System.Windows.Forms.TextBox
    $txtInput.Location = New-Object System.Drawing.Point(120, 92)
    $txtInput.Size = New-Object System.Drawing.Size(120, 28)
    $txtInput.CharacterCasing = "Upper"

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Location = New-Object System.Drawing.Point(210, 140)
    $btnOk.Size = New-Object System.Drawing.Size(80, 32)
    $btnOk.Text = "OK"
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(300, 140)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 32)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $prompt.Controls.AddRange(@($lblMessage, $lblCode, $lblInput, $txtInput, $btnOk, $btnCancel))
    $prompt.AcceptButton = $btnOk
    $prompt.CancelButton = $btnCancel

    $result = $prompt.ShowDialog($form)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
      throw (T "DeletePromptCancelled")
    }

    $inputCode = ([string]$txtInput.Text).Trim().ToUpperInvariant()
    return [pscustomobject]@{
      ExpectedCode = $code
      InputCode = $inputCode
    }
  }

  function Refresh-BaseTableItems {
    $tableChoices = @()
    if ($script:Settings.cachedTables) {
      foreach ($t in @($script:Settings.cachedTables)) {
        $name = [string]$t.name
        $tableChoices += [pscustomobject]@{
          name = $name
          display = Build-TableDisplayText $name ([string]$t.label)
        }
      }
    }

    $cmbBaseTable.BeginUpdate()
    $cmbBaseTable.Items.Clear()
    foreach ($t in $tableChoices) {
      [void]$cmbBaseTable.Items.Add([string]$t.display)
    }
    $cmbBaseTable.EndUpdate()

    $colJoinTable.Items.Clear()
    if ($tableChoices.Count -gt 0) {
      foreach ($t in $tableChoices) {
        [void]$colJoinTable.Items.Add([string]$t.display)
      }
    }
    $cmbDeleteTable.BeginUpdate()
    $cmbDeleteTable.Items.Clear()
    $cmbAttachmentTable.BeginUpdate()
    $cmbAttachmentTable.Items.Clear()
    if ($script:Settings.cachedTables) {
      foreach ($t in @($script:Settings.cachedTables)) {
        [void]$cmbDeleteTable.Items.Add(("{0} - {1}" -f $t.name, $t.label))
        [void]$cmbAttachmentTable.Items.Add(("{0} - {1}" -f $t.name, $t.label))
      }
    }
    $cmbDeleteTable.EndUpdate()
    $cmbAttachmentTable.EndUpdate()

    $deleteTableName = ([string]$script:Settings.deleteTargetTable).Trim()
    if (-not [string]::IsNullOrWhiteSpace($deleteTableName)) {
      $deleteCandidate = $null
      foreach ($item in $cmbDeleteTable.Items) {
        $itemText = [string]$item
        if ($itemText.StartsWith($deleteTableName + " - ")) {
          $deleteCandidate = $item
          break
        }
      }
      if ($deleteCandidate) {
        $cmbDeleteTable.SelectedItem = $deleteCandidate
      } else {
        $cmbDeleteTable.Text = $deleteTableName
      }
    }

    $attachmentTableName = ([string]$script:Settings.attachmentSelectedTableName).Trim()
    if (-not [string]::IsNullOrWhiteSpace($attachmentTableName)) {
      $attachmentCandidate = $null
      foreach ($item in $cmbAttachmentTable.Items) {
        $itemText = [string]$item
        if ($itemText.StartsWith($attachmentTableName + " - ")) {
          $attachmentCandidate = $item
          break
        }
      }
      if ($attachmentCandidate) {
        $cmbAttachmentTable.SelectedItem = $attachmentCandidate
      } else {
        $cmbAttachmentTable.Text = $attachmentTableName
      }
    }

    Refresh-DeleteExecuteButton
  }

  function Get-JoinDefinitions {
    $defs = New-Object System.Collections.Generic.List[object]
    foreach ($row in $gridJoins.Rows) {
      if ($row.IsNewRow) { continue }
      try {
        $tableCell = $row.Cells[0].Value
        $sourceCell = $row.Cells[1].Value
        $baseCell = $row.Cells[2].Value
        $targetCell = $row.Cells[3].Value
        $prefixCell = $row.Cells[4].Value
        $leftJoinCell = $row.Cells[5].Value

        $joinSource = if ($null -eq $sourceCell) { "" } else { ([string]$sourceCell).Trim() }
        $joinTable = if ($null -eq $tableCell) { "" } else { Convert-DisplayTokenToName ([string]$tableCell) }
        $baseColumn = if ($null -eq $baseCell) { "" } else { Convert-DisplayTokenToName ([string]$baseCell) }
        $targetColumn = if ($null -eq $targetCell) { "" } else { Convert-DisplayTokenToName ([string]$targetCell) }
        $joinPrefix = if ($null -eq $prefixCell) { "" } else { ([string]$prefixCell).Trim() }

        $leftJoin = $false
        if ($leftJoinCell -is [bool]) {
          $leftJoin = [bool]$leftJoinCell
        } elseif ($leftJoinCell -is [System.Windows.Forms.CheckState]) {
          $leftJoin = ([System.Windows.Forms.CheckState]$leftJoinCell -eq [System.Windows.Forms.CheckState]::Checked)
        } elseif ($null -ne $leftJoinCell) {
          $text = ([string]$leftJoinCell).Trim()
          if (-not [string]::IsNullOrWhiteSpace($text)) {
            try { $leftJoin = [System.Convert]::ToBoolean($text) } catch { $leftJoin = $false }
          }
        }

        if ([string]::IsNullOrWhiteSpace($joinTable) -and [string]::IsNullOrWhiteSpace($baseColumn) -and [string]::IsNullOrWhiteSpace($targetColumn) -and [string]::IsNullOrWhiteSpace($joinPrefix) -and (-not $leftJoin)) { continue }
        [void]$defs.Add([pscustomobject]@{
          joinTable = $joinTable
          joinSource = $joinSource
          baseColumn = $baseColumn
          targetColumn = $targetColumn
          joinPrefix = $joinPrefix
          leftJoin = $leftJoin
        })
      } catch {
        Add-Log ("Skip invalid join row: {0}" -f $_.Exception.Message)
      }
    }
    return $defs.ToArray()
  }

  function Save-JoinDefinitionsToSettings {
    try {
      $defs = @(Get-JoinDefinitions)
      $script:Settings.viewEditorJoinsJson = ($defs | ConvertTo-Json -Depth 4 -Compress)
      Request-SaveSettings
    } catch {
      Add-Log ("Failed to save join definitions: {0}" -f $_.Exception.Message)
    }
  }

  function Split-JoinSettingTokens([object]$value) {
    if ($null -eq $value) { return @() }
    if ($value -is [bool]) { return @([string]$value) }

    $text = ([string]$value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }

    $lines = @($text -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($lines.Count -gt 1) {
      return @($lines | ForEach-Object { ([string]$_).Trim() })
    }

    return @($text -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
  }

  function Normalize-JoinDefinitionsForLoad([object]$rawJoinDefs) {
    $normalized = New-Object System.Collections.Generic.List[object]
    if ($null -eq $rawJoinDefs) { return @() }

    $candidates = @($rawJoinDefs)
    if (($rawJoinDefs -isnot [System.Array]) -and ($rawJoinDefs -is [System.Collections.IEnumerable]) -and ($rawJoinDefs -isnot [string])) {
      $tmp = @()
      foreach ($item in $rawJoinDefs) { $tmp += $item }
      $candidates = $tmp
    }

    foreach ($j in $candidates) {
      if ($null -eq $j) { continue }
      $props = @($j.PSObject.Properties.Name)
      if ($props.Count -eq 0) { continue }

      $joinTables = @(Split-JoinSettingTokens $j.joinTable)
      $joinSources = @(Split-JoinSettingTokens $j.joinSource)
      $baseColumns = @(Split-JoinSettingTokens $j.baseColumn)
      $targetColumns = @(Split-JoinSettingTokens $j.targetColumn)
      $joinPrefixes = @(Split-JoinSettingTokens $j.joinPrefix)
      $leftJoinTokens = @(Split-JoinSettingTokens $j.leftJoin)

      $rowCountCandidates = @($joinTables.Count, $joinSources.Count, $baseColumns.Count, $targetColumns.Count, $joinPrefixes.Count, $leftJoinTokens.Count)
      $rowCount = (@($rowCountCandidates | Measure-Object -Maximum)[0]).Maximum
      if ($rowCount -lt 1) { $rowCount = 1 }

      if ($rowCount -eq 1) {
        [void]$normalized.Add([pscustomobject]@{
          joinTable = if ($joinTables.Count -gt 0) { [string]$joinTables[0] } else { [string]$j.joinTable }
          joinSource = if ($joinSources.Count -gt 0) { [string]$joinSources[0] } elseif ($j.PSObject.Properties.Name -contains "joinSource") { [string]$j.joinSource } else { "__base__" }
          baseColumn = if ($baseColumns.Count -gt 0) { [string]$baseColumns[0] } else { [string]$j.baseColumn }
          targetColumn = if ($targetColumns.Count -gt 0) { [string]$targetColumns[0] } else { [string]$j.targetColumn }
          joinPrefix = if ($joinPrefixes.Count -gt 0) { [string]$joinPrefixes[0] } else { [string]$j.joinPrefix }
          leftJoin = if ($leftJoinTokens.Count -gt 0) { try { [System.Convert]::ToBoolean($leftJoinTokens[0]) } catch { $false } } elseif ($j.PSObject.Properties.Name -contains "leftJoin") { try { [System.Convert]::ToBoolean($j.leftJoin) } catch { $false } } else { $false }
        })
        continue
      }

      for ($i = 0; $i -lt $rowCount; $i++) {
        $source = ""
        if ($i -lt $joinSources.Count) {
          $source = [string]$joinSources[$i]
        } elseif ($i -eq 0) {
          $source = "__base__"
        } elseif (($i - 1) -lt $joinPrefixes.Count) {
          $source = [string]$joinPrefixes[$i - 1]
        }

        $leftJoin = $false
        if ($i -lt $leftJoinTokens.Count) {
          try { $leftJoin = [System.Convert]::ToBoolean($leftJoinTokens[$i]) } catch { $leftJoin = $false }
        }

        [void]$normalized.Add([pscustomobject]@{
          joinTable = if ($i -lt $joinTables.Count) { [string]$joinTables[$i] } else { "" }
          joinSource = $source
          baseColumn = if ($i -lt $baseColumns.Count) { [string]$baseColumns[$i] } else { "" }
          targetColumn = if ($i -lt $targetColumns.Count) { [string]$targetColumns[$i] } else { "" }
          joinPrefix = if ($i -lt $joinPrefixes.Count) { [string]$joinPrefixes[$i] } else { "" }
          leftJoin = $leftJoin
        })
      }
    }

    return $normalized.ToArray()
  }

  function Fetch-ColumnsForTable([string]$table) {
    if ([string]::IsNullOrWhiteSpace($table)) { return @() }
    $cacheKey = $table.Trim().ToLowerInvariant()
    if ($script:ColumnCache.ContainsKey($cacheKey)) { return @($script:ColumnCache[$cacheKey]) }

    $tableNames = New-Object System.Collections.Generic.List[string]
    [void]$tableNames.Add($table)
    $visited = @{}
    $currentTable = $table
    while (-not [string]::IsNullOrWhiteSpace($currentTable) -and -not $visited.ContainsKey($currentTable)) {
      $visited[$currentTable] = $true
      $objQuery = UrlEncode ("name={0}" -f $currentTable)
      $objPath = "/api/now/table/sys_db_object?sysparm_fields=name,super_class&sysparm_limit=1&sysparm_query=$objQuery"
      $objRes = Invoke-SnowGet $objPath
      $objResults = if ($objRes -and ($objRes.PSObject.Properties.Name -contains "result")) { @($objRes.result) } else { @() }
      $obj = if ((@($objResults)).Count -gt 0) { (@($objResults))[0] } else { $null }
      if (-not $obj) { break }

      $superSysId = ""
      if ($obj.super_class) {
        if ($obj.super_class -is [string]) {
          $superSysId = [string]$obj.super_class
        } elseif ($obj.super_class.PSObject.Properties.Name -contains "value") {
          $superSysId = [string]$obj.super_class.value
        }
      }
      if ([string]::IsNullOrWhiteSpace($superSysId)) { break }

      $superPath = "/api/now/table/sys_db_object/{0}?sysparm_fields=name" -f $superSysId
      $superRes = Invoke-SnowGet $superPath
      $superObj = if ($superRes -and ($superRes.PSObject.Properties.Name -contains "result")) { $superRes.result } else { $null }
      $superName = if ($superObj) { [string]$superObj.name } else { "" }
      if ([string]::IsNullOrWhiteSpace($superName)) { break }
      [void]$tableNames.Add($superName)
      $currentTable = $superName
    }

    $q = "nameIN{0}^elementISNOTEMPTY" -f (($tableNames | Select-Object -Unique) -join ",")
    $fields = "element,column_label"
    $path = "/api/now/table/sys_dictionary?sysparm_fields=$fields&sysparm_limit=5000&sysparm_query=$(UrlEncode $q)"
    $res = Invoke-SnowGet $path

    $results = if ($res -and ($res.PSObject.Properties.Name -contains "result")) { $res.result } else { @() }
    $list = @()
    foreach ($r in @($results)) {
      $name = [string]$r.element
      if ([string]::IsNullOrWhiteSpace($name)) { continue }
      $label = [string]$r.column_label
      if ([string]::IsNullOrWhiteSpace($label)) { $label = $name }
      $list += [pscustomobject]@{ name=$name; label=$label }
    }
    $sorted = @($list | Sort-Object name -Unique)
    $script:ColumnCache[$cacheKey] = @($sorted)
    return @($sorted)
  }

  function Build-ViewEditorColumnDisplay([string]$token, [string]$label, [string]$sourceTable, [string]$sourcePrefix) {
    $left = if ([string]::IsNullOrWhiteSpace($token)) { "" } else { $token }
    $right = if ([string]::IsNullOrWhiteSpace($label)) { $left } else { $label }
    if (-not [string]::IsNullOrWhiteSpace($sourcePrefix)) {
      return ("{0} - [{1}] {2}" -f $left, $sourcePrefix, $right)
    }
    if (-not [string]::IsNullOrWhiteSpace($sourceTable)) {
      return ("{0} - [{1}] {2}" -f $left, $sourceTable, $right)
    }
    return ("{0} - {1}" -f $left, $right)
  }

  function Get-SelectedViewFieldTokens {
    $tokens = New-Object System.Collections.Generic.List[string]
    foreach ($item in $clbViewColumns.Items) {
      $text = [string]$item
      if ([string]::IsNullOrWhiteSpace($text)) { continue }
      $idx = $text.IndexOf(" - ")
      $token = if ($idx -gt 0) { $text.Substring(0, $idx).Trim() } else { $text.Trim() }
      if (-not [string]::IsNullOrWhiteSpace($token) -and -not $tokens.Contains($token)) {
        [void]$tokens.Add($token)
      }
    }
    return $tokens.ToArray()
  }

  function Set-CheckedViewFieldTokens([string[]]$tokens) {
    # チェックUIは廃止。互換のため関数は残す（既存設定の読込時も何もしない）。
    return
  }

  function Update-ViewEditorColumnChoices {
    $previousChecked = @(Get-SelectedViewFieldTokens)
    if ($previousChecked.Count -eq 0 -and $script:Settings -and -not [string]::IsNullOrWhiteSpace([string]$script:Settings.viewEditorSelectedColumnsJson)) {
      try {
        $previousChecked = @([string]$script:Settings.viewEditorSelectedColumnsJson | ConvertFrom-Json)
      } catch {
      }
    }
    $scopes = New-Object System.Collections.Generic.List[object]
    $basePrefix = ([string]$txtBasePrefix.Text).Trim()
    if ([string]::IsNullOrWhiteSpace($basePrefix)) { $basePrefix = "t0" }

    $baseTable = Get-SelectedBaseTableName
    if (-not [string]::IsNullOrWhiteSpace($baseTable)) {
      try {
        foreach ($col in @(Fetch-ColumnsForTable $baseTable)) {
          $baseColumn = [string]$col.name
          [void]$scopes.Add([pscustomobject]@{
            token = $baseColumn
            display = Build-ViewEditorColumnDisplay $baseColumn ([string]$col.label) $baseTable ""
            sourceTable = $baseTable
            sourceColumn = $baseColumn
          })
        }
      } catch {
      }
    }

    for ($i = 0; $i -lt $gridJoins.Rows.Count; $i++) {
      $joinRow = $gridJoins.Rows[$i]
      if ($joinRow.IsNewRow) { continue }
      $joinTableCell = $joinRow.Cells[0].Value
      $joinTable = if ($null -eq $joinTableCell) { "" } else { Convert-DisplayTokenToName ([string]$joinTableCell) }
      if ([string]::IsNullOrWhiteSpace($joinTable)) { continue }
      $prefix = Get-JoinRowPrefix $i
      if ([string]::IsNullOrWhiteSpace($prefix)) { continue }

      try {
        foreach ($col in @(Fetch-ColumnsForTable $joinTable)) {
          $token = ("{0}_{1}" -f $prefix, [string]$col.name)
          [void]$scopes.Add([pscustomobject]@{
            token = $token
            display = Build-ViewEditorColumnDisplay $token ([string]$col.label) $joinTable $prefix
            sourceTable = $joinTable
            sourceColumn = [string]$col.name
          })
        }
      } catch {
      }
    }

    $uniqueScopes = @($scopes | Group-Object token | ForEach-Object { $_.Group[0] } | Sort-Object sourceTable, sourceColumn, token)

    $clbViewColumns.BeginUpdate()
    $clbViewColumns.Items.Clear()
    foreach ($scope in $uniqueScopes) {
      [void]$clbViewColumns.Items.Add([string]$scope.display)
    }
    $clbViewColumns.EndUpdate()

    if ($script:Settings) {
      $script:Settings.viewEditorSelectedColumnsJson = (@(Get-SelectedViewFieldTokens) | ConvertTo-Json -Compress)
      Request-SaveSettings
    }

  }

  function Get-JoinRowPrefix([int]$rowIndex) {
    if ($rowIndex -lt 0 -or $rowIndex -ge $gridJoins.Rows.Count) { return "" }
    $prefixCell = $gridJoins.Rows[$rowIndex].Cells[4].Value
    $prefix = if ($null -eq $prefixCell) { "" } else { ([string]$prefixCell).Trim() }
    if ([string]::IsNullOrWhiteSpace($prefix)) { $prefix = ("t{0}" -f ($rowIndex + 1)) }
    return $prefix
  }

  function Resolve-JoinSourceTable([int]$rowIndex, [string]$sourcePrefix) {
    $baseTable = Get-SelectedBaseTableName
    if ([string]::IsNullOrWhiteSpace($sourcePrefix) -or $sourcePrefix -eq "__base__") { return $baseTable }

    for ($i = 0; $i -lt $rowIndex; $i++) {
      if ((Get-JoinRowPrefix $i) -ne $sourcePrefix) { continue }
      $joinTableCell = $gridJoins.Rows[$i].Cells[0].Value
      $joinTable = if ($null -eq $joinTableCell) { "" } else { Convert-DisplayTokenToName ([string]$joinTableCell) }
      if (-not [string]::IsNullOrWhiteSpace($joinTable)) { return $joinTable }
    }
    return ""
  }

  function Populate-JoinSourcesForRow([int]$rowIndex) {
    if ($rowIndex -lt 0 -or $rowIndex -ge $gridJoins.Rows.Count) { return }
    $row = $gridJoins.Rows[$rowIndex]
    $sourceCell = [System.Windows.Forms.DataGridViewComboBoxCell]$row.Cells[1]
    $selectedSource = if ($null -eq $sourceCell.Value) { "" } else { [string]$sourceCell.Value }

    $sources = New-Object System.Collections.Generic.List[string]
    [void]$sources.Add("__base__")
    for ($i = 0; $i -lt $rowIndex; $i++) {
      $joinTableCell = $gridJoins.Rows[$i].Cells[0].Value
      $joinTable = if ($null -eq $joinTableCell) { "" } else { Convert-DisplayTokenToName ([string]$joinTableCell) }
      if ([string]::IsNullOrWhiteSpace($joinTable)) { continue }
      $prefix = Get-JoinRowPrefix $i
      if ([string]::IsNullOrWhiteSpace($prefix)) { continue }
      if (-not $sources.Contains($prefix)) { [void]$sources.Add($prefix) }
    }

    $sourceCell.Items.Clear()
    foreach ($s in $sources) { [void]$sourceCell.Items.Add($s) }

    if (-not [string]::IsNullOrWhiteSpace($selectedSource) -and $sourceCell.Items.Contains($selectedSource)) {
      $sourceCell.Value = $selectedSource
    } else {
      $sourceCell.Value = "__base__"
    }
  }

  function Populate-JoinColumnsForRow([int]$rowIndex) {
    if ($rowIndex -lt 0 -or $rowIndex -ge $gridJoins.Rows.Count) { return }
    $row = $gridJoins.Rows[$rowIndex]
    if ($null -eq $row) { return }

    Populate-JoinSourcesForRow $rowIndex

    $sourceCellValue = $row.Cells[1].Value
    $sourcePrefix = if ($null -eq $sourceCellValue) { "__base__" } else { ([string]$sourceCellValue).Trim() }
    if ([string]::IsNullOrWhiteSpace($sourcePrefix)) { $sourcePrefix = "__base__" }

    $baseTable = Resolve-JoinSourceTable $rowIndex $sourcePrefix
    $joinTableCell = $row.Cells[0].Value
    $joinTable = if ($null -eq $joinTableCell) { "" } else { Convert-DisplayTokenToName ([string]$joinTableCell) }

    $baseColumns = @()
    $joinColumns = @()
    if (-not [string]::IsNullOrWhiteSpace($baseTable)) { $baseColumns = @(Fetch-ColumnsForTable $baseTable) }
    if (-not [string]::IsNullOrWhiteSpace($joinTable)) { $joinColumns = @(Fetch-ColumnsForTable $joinTable) }

    $baseCell = [System.Windows.Forms.DataGridViewComboBoxCell]$row.Cells[2]
    $targetCell = [System.Windows.Forms.DataGridViewComboBoxCell]$row.Cells[3]
    $selectedBase = if ($null -eq $baseCell.Value) { "" } else { [string]$baseCell.Value }
    $selectedTarget = if ($null -eq $targetCell.Value) { "" } else { [string]$targetCell.Value }

    $baseCell.Items.Clear()
    foreach ($c in $baseColumns) {
      $name = [string]$c.name
      [void]$baseCell.Items.Add((Build-ColumnDisplayText $name ([string]$c.label)))
    }
    if (-not [string]::IsNullOrWhiteSpace($selectedBase)) { $baseCell.Value = Resolve-DisplayTextFromItems $baseCell $selectedBase }
    else { $baseCell.Value = $null }

    $targetCell.Items.Clear()
    foreach ($c in $joinColumns) {
      $name = [string]$c.name
      [void]$targetCell.Items.Add((Build-ColumnDisplayText $name ([string]$c.label)))
    }
    if (-not [string]::IsNullOrWhiteSpace($selectedTarget)) { $targetCell.Value = Resolve-DisplayTextFromItems $targetCell $selectedTarget }
    else { $targetCell.Value = $null }
  }




  function Build-JoinWhereClause([string]$leftPrefix, [string]$baseColumn, [string]$joinPrefix, [string]$joinColumn) {
    $left = if ([string]::IsNullOrWhiteSpace($leftPrefix)) { [string]$baseColumn } else { "{0}_{1}" -f [string]$leftPrefix, [string]$baseColumn }
    $right = if ([string]::IsNullOrWhiteSpace($joinPrefix)) { [string]$joinColumn } else { "{0}_{1}" -f [string]$joinPrefix, [string]$joinColumn }
    return ("{0}={1}" -f $left, $right)
  }

  function Test-ViewTableMetadata([psobject]$record, [string]$expectedPrefix, [string]$expectedWhereText, [bool]$expectedLeftJoin, [bool]$shouldCheckLeftJoin) {
    if ($null -eq $record) { return $false }

    if (-not [string]::IsNullOrWhiteSpace($expectedPrefix)) {
      $prefixOk = $false
      if ($record.PSObject.Properties.Name -contains "variable_prefix") {
        $prefixOk = ([string]$record.variable_prefix -eq $expectedPrefix)
      }
      if (-not $prefixOk -and ($record.PSObject.Properties.Name -contains "prefix")) {
        $prefixOk = ([string]$record.prefix -eq $expectedPrefix)
      }
      if (-not $prefixOk) { return $false }
    }

    if (-not [string]::IsNullOrWhiteSpace($expectedWhereText)) {
      $whereOk = $false
      if ($record.PSObject.Properties.Name -contains "where_clause") {
        $whereOk = ([string]$record.where_clause -eq $expectedWhereText)
      }
      if (-not $whereOk -and ($record.PSObject.Properties.Name -contains "where")) {
        $whereOk = ([string]$record.where -eq $expectedWhereText)
      }
      if (-not $whereOk) { return $false }
    }

    if ($shouldCheckLeftJoin) {
      if (-not ($record.PSObject.Properties.Name -contains "left_join")) { return $false }
      if ([System.Convert]::ToBoolean($record.left_join) -ne $expectedLeftJoin) { return $false }
    }

    return $true
  }

  function Save-ViewTableMetadata([string]$viewTableSysId, [string]$prefix, [string]$whereText, [bool]$leftJoin, [bool]$hasLeftJoin) {
    if ([string]::IsNullOrWhiteSpace($viewTableSysId)) { return $false }

    $payloads = @()
    $prefixCandidates = @(
      @{ variable_prefix = $prefix },
      @{ prefix = $prefix },
      @{ prefix = $prefix; variable_prefix = $prefix },
      @{}
    )
    $whereCandidates = @(
      @{ where_clause = $whereText },
      @{ where = $whereText },
      @{ where = $whereText; where_clause = $whereText },
      @{}
    )

    foreach ($pPayload in $prefixCandidates) {
      foreach ($wPayload in $whereCandidates) {
        $payload = @{}
        foreach ($k in $pPayload.Keys) {
          if (-not [string]::IsNullOrWhiteSpace([string]$pPayload[$k])) { $payload[$k] = $pPayload[$k] }
        }
        foreach ($k in $wPayload.Keys) {
          if (-not [string]::IsNullOrWhiteSpace([string]$wPayload[$k])) { $payload[$k] = $wPayload[$k] }
        }
        if ($hasLeftJoin) { $payload["left_join"] = $leftJoin }
        if ($payload.Count -gt 0) { $payloads += $payload }
      }
    }

    foreach ($payload in $payloads) {
      try {
        [void](Invoke-SnowPatch ("/api/now/table/sys_db_view_table/{0}" -f $viewTableSysId) $payload)

        $verifyPath = "/api/now/table/sys_db_view_table/{0}?sysparm_fields=prefix,variable_prefix,where,where_clause,left_join" -f $viewTableSysId
        $verifyRes = Invoke-SnowGet $verifyPath
        $verifyRecord = if ($verifyRes -and ($verifyRes.PSObject.Properties.Name -contains "result")) { $verifyRes.result } else { $null }
        if (Test-ViewTableMetadata $verifyRecord $prefix $whereText $leftJoin $hasLeftJoin) {
          return $true
        }
      } catch {
      }
    }
    return $false
  }

  function Try-CreateViewJoinRow([string]$sysId, [psobject]$joinDef, [string]$joinWhereClause, [string]$joinPrefix, [bool]$isLeftJoin, [int]$joinOrder) {
    $joinBody = @{
      view = $sysId
      table = [string]$joinDef.joinTable
      left_field = [string]$joinDef.baseColumn
      right_field = [string]$joinDef.targetColumn
      join_condition = $joinWhereClause
      variable_prefix = $joinPrefix
      left_join = $isLeftJoin
      order = $joinOrder
    }

    $saved = $false
    $joinRowId = ""
    try {
      $joinRes = Invoke-SnowPost "/api/now/table/sys_db_view_table" $joinBody
      if ($joinRes -and ($joinRes.PSObject.Properties.Name -contains "result") -and $joinRes.result) {
        $joinRowId = [string]$joinRes.result.sys_id
      }
      $saved = $true
    } catch {
      foreach ($leftField in @("left_field", "left_column", "field")) {
        foreach ($rightField in @("right_field", "right_column", "join_field")) {
          try {
            $fallbackBody = @{ view = $sysId; table = [string]$joinDef.joinTable; order = $joinOrder }
            $fallbackBody[$leftField] = [string]$joinDef.baseColumn
            $fallbackBody[$rightField] = [string]$joinDef.targetColumn
            $joinRes = Invoke-SnowPost "/api/now/table/sys_db_view_table" $fallbackBody
            if ($joinRes -and ($joinRes.PSObject.Properties.Name -contains "result") -and $joinRes.result) {
              $joinRowId = [string]$joinRes.result.sys_id
            }
            $saved = $true
            break
          } catch {
          }
        }
        if ($saved) { break }
      }
    }

    return [pscustomobject]@{ saved = $saved; rowId = $joinRowId }
  }


  function Fetch-ColumnsForBaseTable {
    $table = Get-SelectedBaseTableName
    if ([string]::IsNullOrWhiteSpace($table)) {
      [System.Windows.Forms.MessageBox]::Show((T "WarnBaseTable")) | Out-Null
      return
    }

    Add-Log ("{0} [{1}]" -f (T "FetchingColumns"), $table)
    Invoke-Async "Fetch-Columns" {
      param($state)
      $tableName = [string]$state
      $list = @(Fetch-ColumnsForTable $tableName)
      return [pscustomobject]@{ table = $tableName; count = @($list).Count }
    } {
      param($result)
      for ($i = 0; $i -lt $gridJoins.Rows.Count; $i++) {
        Populate-JoinColumnsForRow $i
      }
      Update-ViewEditorColumnChoices
      Add-Log ("{0}: {1}" -f (T "ColumnsFetched"), [int]$result.count)
    } $table
  }

  function Create-DatabaseView {
    $viewName = ([string]$txtViewName.Text).Trim()
    $viewLabel = ([string]$txtViewLabel.Text).Trim()
    $baseTable = Get-SelectedBaseTableName
    $joinDefs = @(Get-JoinDefinitions)

    $validation = Validate-ViewInput -ViewName $viewName -ViewLabel $viewLabel -BaseTable $baseTable -JoinDefinitions @($joinDefs) -GetText ${function:T}
    if (-not $validation.IsValid) {
      [System.Windows.Forms.MessageBox]::Show([string]$validation.Errors[0]) | Out-Null
      return
    }

    $basePrefix = ([string]$txtBasePrefix.Text).Trim()
    if ([string]::IsNullOrWhiteSpace($basePrefix)) { $basePrefix = "t0" }
    $selectedColumns = @(Get-SelectedViewFieldTokens)

    Add-Log ("Creating DB view: {0}, base={1}, joins={2}" -f $viewName, $baseTable, $joinDefs.Count)
    $ctx = [pscustomobject]@{ viewName=$viewName; viewLabel=$viewLabel; baseTable=$baseTable; basePrefix=$basePrefix; selectedColumns=@($selectedColumns); joinDefs=@($joinDefs) }
    Invoke-Async "Create-DatabaseView" {
      param($state)
      return Invoke-CreateViewUseCase -Context $state -InvokeSnowPost ${function:Invoke-SnowPost} -InvokeSnowPatch ${function:Invoke-SnowPatch} -InvokeSnowGet ${function:Invoke-SnowGet} -UrlEncode ${function:UrlEncode} -SaveViewTableMetadata ${function:Save-ViewTableMetadata} -BuildJoinWhereClause ${function:Build-JoinWhereClause} -TryCreateViewJoinRow ${function:Try-CreateViewJoinRow}
    } {
      param($result)
      if ([string]::IsNullOrWhiteSpace([string]$result.sysId)) {
        Add-Log ("{0}: {1}" -f (T "ViewCreateFailed"), [string]$result.viewName)
        [System.Windows.Forms.MessageBox]::Show((T "ViewCreateFailed")) | Out-Null
        return
      }
      if (-not [bool]$result.joinsSaved) {
        Add-Log (T "ViewJoinFallback")
        [System.Windows.Forms.MessageBox]::Show((T "ViewJoinFallback")) | Out-Null
      }
      Update-CreatedViewLinks ([string]$result.viewName) ([string]$result.sysId)
      Add-Log ("{0}: {1}" -f (T "ViewCreated"), [string]$result.viewName)
      [System.Windows.Forms.MessageBox]::Show(("{0}`r`n{1}" -f (T "ViewCreated"), [string]$result.viewName)) | Out-Null
    } $ctx
  }

  # ----------------------------
  # Export
  # ----------------------------
  function Build-QueryString {
    if ($rbAll.Checked) { return "" }

    $start = $dtStart.Value
    $end = $dtEnd.Value
    if ($end -lt $start) { $tmp = $start; $start = $end; $end = $tmp }

    $q = "sys_updated_onBETWEENjavascript:gs.dateGenerate('{0}','{1}')@javascript:gs.dateGenerate('{2}','{3}')" -f `
      $start.ToString("yyyy-MM-dd"), $start.ToString("HH:mm:ss"),
      $end.ToString("yyyy-MM-dd"),   $end.ToString("HH:mm:ss")
    return $q
  }

  function Get-AttachmentLastRunMap {
    $map = @{}
    if ($script:Settings -and ($script:Settings.PSObject.Properties.Name -contains 'attachmentHarvesterLastRunMap')) {
      $src = $script:Settings.attachmentHarvesterLastRunMap
      if ($src -is [hashtable]) {
        $map = $src
      } elseif ($src) {
        foreach ($p in $src.PSObject.Properties) {
          $map[$p.Name] = [string]$p.Value
        }
      }
    }
    return $map
  }

  function Set-AttachmentLastRunRangeFromHistory {
    $table = Get-SelectedAttachmentTableName
    $dateField = [string]$cmbAttachmentDateField.SelectedItem
    if ([string]::IsNullOrWhiteSpace($table) -or [string]::IsNullOrWhiteSpace($dateField)) { return }

    $key = "{0}:{1}" -f $table, $dateField
    $map = Get-AttachmentLastRunMap
    $now = Get-Date
    if ($map.ContainsKey($key)) {
      try {
        $dtAttachmentStart.Value = [datetime]::ParseExact([string]$map[$key], 'yyyy-MM-dd HH:mm:ss', $null)
      } catch {
        $dtAttachmentStart.Value = $now.AddDays(-30)
      }
      $dtAttachmentEnd.Value = $now
      return
    }

    $result = [System.Windows.Forms.MessageBox]::Show(
      '前回取得時間の記録がないため過去30日を自動設定します',
      'Attachment Harvester',
      [System.Windows.Forms.MessageBoxButtons]::OKCancel,
      [System.Windows.Forms.MessageBoxIcon]::Information
    )
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
      $dtAttachmentStart.Value = $now.AddDays(-30)
      $dtAttachmentEnd.Value = $now
    }
  }

  function Harvest-Attachments {
    $table = Get-SelectedAttachmentTableName
    $dateField = [string]$cmbAttachmentDateField.SelectedItem
    $baseDir = [string]$txtAttachmentDir.Text
    if ([string]::IsNullOrWhiteSpace($baseDir)) {
      $baseDir = [string]$script:Settings.attachmentDownloadDirectory
    }
    if ([string]::IsNullOrWhiteSpace($baseDir)) {
      $baseDir = $DefaultAttachmentDir
    }
    if (-not (Test-Path $baseDir)) { [void](New-Item -ItemType Directory -Path $baseDir -Force) }
    $txtAttachmentDir.Text = $baseDir

    $validation = Validate-AttachmentHarvesterInput -BaseUrl (Get-BaseUrl) -Table $table -DownloadDirectory $baseDir -Settings $script:Settings -UnprotectSecret ${function:Unprotect-Secret} -GetText ${function:T}
    if (-not $validation.IsValid) {
      [System.Windows.Forms.MessageBox]::Show([string]$validation.Errors[0]) | Out-Null
      return
    }

    $start = $dtAttachmentStart.Value
    $end = $dtAttachmentEnd.Value
    if ($start -gt $end) { $tmp = $start; $start = $end; $end = $tmp }

    $script:Settings.attachmentDownloadDirectory = $baseDir
    $script:Settings.attachmentCreateSubfolderPerTable = [bool]$chkAttachmentSubfolder.Checked
    $script:Settings.attachmentFilterDateField = $dateField
    $script:Settings.attachmentStartDateTime = $start.ToString('yyyy-MM-dd HH:mm:ss')
    $script:Settings.attachmentEndDateTime = $end.ToString('yyyy-MM-dd HH:mm:ss')
    $script:Settings.attachmentSelectedTableName = $table
    Request-SaveSettings

    Add-AttachmentLog ("Target table: {0}" -f $table)
    Add-AttachmentLog ("DateField: {0}" -f $dateField)
    Add-AttachmentLog ("Date range: {0} - {1}" -f $start.ToString('yyyy-MM-dd HH:mm:ss'), $end.ToString('yyyy-MM-dd HH:mm:ss'))
    Add-AttachmentLog ("Save directory: {0}" -f $baseDir)

    $ctx = [pscustomobject]@{
      table = $table
      dateField = $dateField
      startDateTime = $start
      endDateTime = $end
      downloadDirectory = $baseDir
      createSubfolderPerTable = [bool]$chkAttachmentSubfolder.Checked
    }

    Invoke-Async "Attachment-Harvester" {
      param($state)
      return Invoke-AttachmentHarvesterUseCase -Context $state -InvokeSnowGet ${function:Invoke-SnowGet} -UrlEncode ${function:UrlEncode} -DownloadAttachmentBytes ${function:Invoke-SnowDownloadAttachmentBytes} -WriteLog ${function:Add-AttachmentLog}
    } {
      param($result)
      Add-AttachmentLog ("saved={0}, skipped={1}, failed={2}" -f [int]$result.Saved, [int]$result.Skipped, [int]$result.Failed)
      if ([bool]$result.Success) {
        $map = Get-AttachmentLastRunMap
        $map["{0}:{1}" -f $table, $dateField] = $end.ToString('yyyy-MM-dd HH:mm:ss')
        $script:Settings.attachmentHarvesterLastRunMap = $map
        Request-SaveSettings
      }
      [System.Windows.Forms.MessageBox]::Show(("Attachment Harvester complete`r`nSaved: {0}`r`nSkipped: {1}`r`nFailed: {2}" -f [int]$result.Saved, [int]$result.Skipped, [int]$result.Failed)) | Out-Null
    } $ctx
  }

  function Remove-AllTableRecords {
    $table = Get-SelectedDeleteTableName
    $maxRetries = [int]$numDeleteMaxRetries.Value

    $verification = Request-DeleteVerificationCode
    $expectedCode = [string]$verification.ExpectedCode
    $actualCode = [string]$verification.InputCode

    $validation = Validate-TruncateInput -Table $table -MaxRetries $maxRetries -ExpectedCode $expectedCode -InputCode $actualCode -GetText ${function:T} -IsInstanceAllowed (Test-TruncateInstanceAllowed)
    if (-not $validation.IsValid) {
      throw [string]$validation.Errors[0]
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
      ([string]::Format((T "DeleteConfirmMessage"), $table, $expectedCode, (Get-BaseUrl), [int]$maxRetries)),
      (T "DeleteConfirmTitle"),
      [System.Windows.Forms.MessageBoxButtons]::YesNo,
      [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
      return
    }

    $result = Invoke-TruncateUseCase -Context ([pscustomobject]@{ table = $table; maxRetries = $maxRetries }) -InvokeSnowGet ${function:Invoke-SnowGet} -InvokeSnowDelete ${function:Invoke-SnowDelete} -InvokeSnowBatchDelete ${function:Invoke-SnowBatchDelete} -WriteLog ${function:Add-Log} -SetProgress ${function:Set-DeleteProgress} -GetText ${function:T}

    if ($result.Status -eq "NoRecord") {
      [System.Windows.Forms.MessageBox]::Show((T "DeleteNoRecord")) | Out-Null
      return
    }
    if ($result.Status -eq "Done") {
      [System.Windows.Forms.MessageBox]::Show((T "DeleteDone")) | Out-Null
      return
    }

    [System.Windows.Forms.MessageBox]::Show((T "DeleteStopped")) | Out-Null
  }

  function Export-Table {
    $table = Get-SelectedTableName

    $validation = Validate-ExportInput -BaseUrl (Get-BaseUrl) -Table $table -Settings $script:Settings -UnprotectSecret ${function:Unprotect-Secret} -GetText ${function:T}
    if (-not $validation.IsValid) {
      [System.Windows.Forms.MessageBox]::Show([string]$validation.Errors[0]) | Out-Null
      return
    }

    $exportDir = Ensure-ExportDir $txtDir.Text
    $script:Settings.exportDirectory = $exportDir
    Request-SaveSettings

    $query = Build-QueryString

    $pageSizeVal = $script:Settings.pageSize
    if ($null -eq $pageSizeVal) { $pageSizeVal = 1000 }
    $pageSize = [int]$pageSizeVal
    if ($pageSize -lt 100) { $pageSize = 100 }
    if ($pageSize -gt 5000) { $pageSize = 5000 }

    $maxRowsVal = [int]$numExportMaxRows.Value
    if ($maxRowsVal -lt 1) { $maxRowsVal = 1 }

    Add-Log (T "Exporting")
    Add-Log ("table={0}, pageSize={1}, maxRows={2}" -f $table, $pageSize, $maxRowsVal)
    Add-Log ("outputFormat={0}" -f [string]$script:Settings.outputFormat)

    $fieldsVal = $script:Settings.exportFields
    if ($null -eq $fieldsVal) { $fieldsVal = "" }
    $fields = ([string]$fieldsVal).Trim()

    $formatVal = [string]$script:Settings.outputFormat
    if ([string]::IsNullOrWhiteSpace($formatVal)) { $formatVal = "csv" }
    $format = $formatVal.Trim().ToLowerInvariant()
    if ((@("csv","json","xlsx") -notcontains $format)) { $format = "csv" }

    $outputEncoding = "utf-8"
    $outputBom = $true
    try { $outputBom = [bool]$script:Settings.outputBom } catch { $outputBom = $true }
    if ((@("csv", "json") -notcontains $format)) { $outputBom = $false }

    Add-Log ("outputEncoding={0}, outputBom={1}" -f $outputEncoding, $outputBom)
    if (-not [string]::IsNullOrWhiteSpace($query)) { Add-Log ("query={0}" -f $query) }

    $stamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $suffix = if ($rbBetween.Checked) {
      ("_{0}-{1}" -f $dtStart.Value.ToString("yyyyMMddHHmmss"), $dtEnd.Value.ToString("yyyyMMddHHmmss"))
    } else { "" }
    $ext = switch ($format) {
      "json" { "json" }
      "xlsx" { "xlsx" }
      default { "csv" }
    }
    $file = Join-Path $exportDir ("{0}{1}_{2}.{3}" -f $table, $suffix, $stamp, $ext)

    $ctx = [pscustomobject]@{ table=$table; pageSize=$pageSize; maxRows=$maxRowsVal; query=$query; fields=$fields; format=$format; file=$file; outputEncoding=$outputEncoding; outputBom=$outputBom }

    Invoke-Async "Export-Table" {
      param($state)
      return Invoke-ExportUseCase -Context $state -InvokeSnowGet ${function:Invoke-SnowGet} -UrlEncode ${function:UrlEncode}
    } {
      param($result)
      if ([int]$result.total -eq 0) {
        Add-Log "0 records."
        [System.Windows.Forms.MessageBox]::Show("0 records.") | Out-Null
        return
      }
      $resultFiles = @()
      if ($result -and ($result.PSObject.Properties.Name -contains "files")) {
        $resultFiles = @($result.files | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
      }
      if ($resultFiles.Count -lt 1 -and -not [string]::IsNullOrWhiteSpace([string]$result.file)) {
        $resultFiles = @([string]$result.file)
      }

      $donePathText = [string]::Join(", ", $resultFiles)
      Add-Log ("{0}: {1}" -f (T "Done"), $donePathText)
      if ($result -and ($result.PSObject.Properties.Name -contains "outputEncoding")) {
        Add-Log ("encodedAs={0}, bom={1}" -f [string]$result.outputEncoding, [bool]$result.outputBom)
      }
      $fileLines = [string]::Join("`r`n", $resultFiles)
      [System.Windows.Forms.MessageBox]::Show(("OK`r`n{0}`r`nRecords: {1}" -f $fileLines, [int]$result.total)) | Out-Null
    } $ctx
  }

  # ----------------------------
  # Initialize from settings
  # ----------------------------
  $cmbLang.SelectedItem = [string]$script:Settings.uiLanguage
  $initialTheme = ([string]$script:Settings.uiTheme).Trim().ToLowerInvariant()
  if (@("dark","light") -notcontains $initialTheme) { $initialTheme = "dark" }
  $cmbTheme.SelectedItem = $initialTheme
  if (-not $cmbLang.SelectedItem) { $cmbLang.SelectedItem = "ja" }

  $txtInstance.Text = [string]$script:Settings.instanceName
  $txtUser.Text = [string]$script:Settings.userId

  if ([string]::IsNullOrWhiteSpace([string]$script:Settings.exportDirectory)) {
    $txtDir.Text = $DefaultExportDir
  } else {
    $txtDir.Text = [string]$script:Settings.exportDirectory
  }

  if ([string]::IsNullOrWhiteSpace([string]$script:Settings.attachmentDownloadDirectory)) {
    $txtAttachmentDir.Text = $DefaultAttachmentDir
  } else {
    $txtAttachmentDir.Text = [string]$script:Settings.attachmentDownloadDirectory
  }

  if ([string]::IsNullOrWhiteSpace([string]$script:Settings.logOutputDirectory)) {
    $txtLogDir.Text = $DefaultLogDir
  } else {
    $txtLogDir.Text = [string]$script:Settings.logOutputDirectory
  }

  if ([string]$script:Settings.filterMode -eq "updated_between") { $rbBetween.Checked = $true } else { $rbAll.Checked = $true }

  $initialOutputFormat = ([string]$script:Settings.outputFormat).Trim().ToLowerInvariant()
  if ((@("csv","json","xlsx") -notcontains $initialOutputFormat)) { $initialOutputFormat = "csv" }
  $cmbOutputFormat.SelectedItem = $initialOutputFormat

  $initialOutputBom = $true
  try { $initialOutputBom = [bool]$script:Settings.outputBom } catch { $initialOutputBom = $true }
  $chkOutputBom.Checked = $initialOutputBom

  function Update-OutputOptionState {
    $selectedFormat = ([string]$cmbOutputFormat.SelectedItem).Trim().ToLowerInvariant()
    $supportsBom = (@("csv", "json") -contains $selectedFormat)
    $chkOutputBom.Enabled = $supportsBom
    if (-not $supportsBom) {
      $chkOutputBom.Checked = $false
    }
  }
  Update-OutputOptionState

  $initialExportMaxRows = 10000
  try { $initialExportMaxRows = [int]$script:Settings.exportMaxRows } catch { $initialExportMaxRows = 10000 }
  if ($initialExportMaxRows -lt [int]$numExportMaxRows.Minimum -or $initialExportMaxRows -gt [int]$numExportMaxRows.Maximum) { $initialExportMaxRows = 10000 }
  $numExportMaxRows.Value = $initialExportMaxRows

  $txtDeleteAllowedInstances.Text = [string]::Join(",", @(Get-TruncateAllowedInstancePatterns))

  $initialDeleteMaxRetries = 99
  try { $initialDeleteMaxRetries = [int]$script:Settings.deleteMaxRetries } catch { $initialDeleteMaxRetries = 99 }
  if ($initialDeleteMaxRetries -lt 1 -or $initialDeleteMaxRetries -gt 999) { $initialDeleteMaxRetries = 99 }
  $numDeleteMaxRetries.Value = $initialDeleteMaxRetries

  try { $dtStart.Value = [datetime]::Parse([string]$script:Settings.startDateTime) } catch { }
  try { $dtEnd.Value   = [datetime]::Parse([string]$script:Settings.endDateTime) } catch { }
  try { $dtAttachmentStart.Value = [datetime]::Parse([string]$script:Settings.attachmentStartDateTime) } catch { }
  try { $dtAttachmentEnd.Value = [datetime]::Parse([string]$script:Settings.attachmentEndDateTime) } catch { }

  $initialAttachmentDateField = ([string]$script:Settings.attachmentFilterDateField).Trim()
  if ((@('sys_created_on','sys_updated_on') -notcontains $initialAttachmentDateField)) { $initialAttachmentDateField = 'sys_updated_on' }
  $cmbAttachmentDateField.SelectedItem = $initialAttachmentDateField
  $chkAttachmentSubfolder.Checked = [bool]$script:Settings.attachmentCreateSubfolderPerTable

  if ((Resolve-CoreSnowAuthType -AuthType $script:Settings.authType) -eq "apikey") { $rbApiKey.Checked = $true } else { $rbUserPass.Checked = $true }

  $txtPass.Text = Unprotect-Secret ([string]$script:Settings.passwordEnc)
  $txtKey.Text  = Unprotect-Secret ([string]$script:Settings.apiKeyEnc)

  if (@($script:Settings.cachedTables).Count -gt 0) {
    $cmbTable.BeginUpdate()
    $cmbTable.Items.Clear()
    foreach ($t in @($script:Settings.cachedTables)) {
      [void]$cmbTable.Items.Add(("{0} - {1}" -f $t.name, $t.label))
    }
    $cmbTable.EndUpdate()
    Refresh-BaseTableItems
  }

  $initialAttachmentTableName = ([string]$script:Settings.attachmentSelectedTableName).Trim()
  if (-not [string]::IsNullOrWhiteSpace($initialAttachmentTableName) -and @($cmbAttachmentTable.Items).Count -eq 0) {
    $cmbAttachmentTable.Text = $initialAttachmentTableName
  }

  $initialTableName = ([string]$script:Settings.selectedTableName).Trim()
  if (-not [string]::IsNullOrWhiteSpace($initialTableName)) {
    $candidate = $null
    foreach ($item in $cmbTable.Items) {
      $itemText = [string]$item
      if ($itemText.StartsWith($initialTableName + " - ")) {
        $candidate = $item
        break
      }
    }
    if ($candidate) {
      $cmbTable.SelectedItem = $candidate
    } else {
      $cmbTable.Text = $initialTableName
    }
  }

  $txtViewName.Text = [string]$script:Settings.viewEditorViewName
  $txtViewLabel.Text = [string]$script:Settings.viewEditorViewLabel

  $initialBaseTableName = ([string]$script:Settings.viewEditorBaseTable).Trim()
  if (-not [string]::IsNullOrWhiteSpace($initialBaseTableName)) {
    $baseCandidate = $null
    foreach ($item in $cmbBaseTable.Items) {
      $itemText = [string]$item
      if ($itemText.StartsWith($initialBaseTableName + " - ")) {
        $baseCandidate = $item
        break
      }
    }
    if ($baseCandidate) {
      $cmbBaseTable.SelectedItem = $baseCandidate
    } else {
      $cmbBaseTable.Text = $initialBaseTableName
    }
  }

  $txtBasePrefix.Text = [string]$script:Settings.viewEditorBasePrefix
  if ([string]::IsNullOrWhiteSpace($txtBasePrefix.Text)) { $txtBasePrefix.Text = "t0" }

  try {
    $joinsText = [string]$script:Settings.viewEditorJoinsJson
    if (-not [string]::IsNullOrWhiteSpace($joinsText)) {
      $loadedJoinDefs = @(Normalize-JoinDefinitionsForLoad ($joinsText | ConvertFrom-Json))
      foreach ($j in $loadedJoinDefs) {
        if ($null -eq $j) { continue }
        $rowIndex = $gridJoins.Rows.Add()
        if ($rowIndex -lt 0) { continue }
        $gridJoins.Rows[$rowIndex].Cells[1].Value = "__base__"
        $gridJoins.Rows[$rowIndex].Cells[5].Value = $false
        $joinTableCell = [System.Windows.Forms.DataGridViewComboBoxCell]$gridJoins.Rows[$rowIndex].Cells[0]
        $gridJoins.Rows[$rowIndex].Cells[0].Value = Resolve-DisplayTextFromItems $joinTableCell ([string]$j.joinTable)
        Populate-JoinColumnsForRow $rowIndex
        if ($j.PSObject.Properties.Name -contains "joinSource") { $gridJoins.Rows[$rowIndex].Cells[1].Value = [string]$j.joinSource }
        else { $gridJoins.Rows[$rowIndex].Cells[1].Value = "__base__" }
        Populate-JoinColumnsForRow $rowIndex
        $baseCell = [System.Windows.Forms.DataGridViewComboBoxCell]$gridJoins.Rows[$rowIndex].Cells[2]
        $targetCell = [System.Windows.Forms.DataGridViewComboBoxCell]$gridJoins.Rows[$rowIndex].Cells[3]
        $gridJoins.Rows[$rowIndex].Cells[2].Value = Resolve-DisplayTextFromItems $baseCell ([string]$j.baseColumn)
        $gridJoins.Rows[$rowIndex].Cells[3].Value = Resolve-DisplayTextFromItems $targetCell ([string]$j.targetColumn)
        $gridJoins.Rows[$rowIndex].Cells[4].Value = [string]$j.joinPrefix
        if ($j.PSObject.Properties.Name -contains "leftJoin") { $gridJoins.Rows[$rowIndex].Cells[5].Value = [System.Convert]::ToBoolean($j.leftJoin) }
      }
    }
  } catch {
  }

  Update-ViewEditorColumnChoices
  try {
    $selectedColsText = [string]$script:Settings.viewEditorSelectedColumnsJson
    if (-not [string]::IsNullOrWhiteSpace($selectedColsText)) {
      $loadedColumns = @($selectedColsText | ConvertFrom-Json)
      if ($loadedColumns.Count -gt 0) { Set-CheckedViewFieldTokens $loadedColumns }
    }
  } catch {
  }

  Update-AuthUI
  Update-FilterUI
  Apply-Language
  Set-DeleteProgress 0 "0%"
  Refresh-DeleteExecuteButton

  # ----------------------------
  # Wire events for auto-save
  # ----------------------------
  $cmbLang.add_SelectedIndexChanged({
    $script:Settings.uiLanguage = [string]$cmbLang.SelectedItem
    Request-SaveSettings
    Apply-Language
  })

  $txtInstance.add_TextChanged({
    $script:Settings.instanceName = $txtInstance.Text
    Request-SaveSettings
    Update-BaseUrlLabel
    Refresh-DeleteExecuteButton
  })

  $rbUserPass.add_CheckedChanged({
    if ($rbUserPass.Checked) {
      $script:Settings.authType = "userpass"
      Request-SaveSettings
      Update-AuthUI
    }
  })
  $rbApiKey.add_CheckedChanged({
    if ($rbApiKey.Checked) {
      $script:Settings.authType = "apikey"
      Request-SaveSettings
      Update-AuthUI
    }
  })

  $txtUser.add_TextChanged({
    $script:Settings.userId = $txtUser.Text
    Request-SaveSettings
  })

  $txtPass.add_TextChanged({
    $script:Settings.passwordEnc = Protect-Secret $txtPass.Text
    Request-SaveSettings
  })

  $txtKey.add_TextChanged({
    $script:Settings.apiKeyEnc = Protect-Secret $txtKey.Text
    Request-SaveSettings
  })

  $rbAll.add_CheckedChanged({
    if ($rbAll.Checked) {
      $script:Settings.filterMode = "all"
      Request-SaveSettings
      Update-FilterUI
    }
  })
  $rbBetween.add_CheckedChanged({
    if ($rbBetween.Checked) {
      $script:Settings.filterMode = "updated_between"
      Request-SaveSettings
      Update-FilterUI
    }
  })

  $dtStart.add_ValueChanged({
    $script:Settings.startDateTime = $dtStart.Value.ToString("yyyy-MM-dd HH:mm:ss")
    Request-SaveSettings
  })
  $dtEnd.add_ValueChanged({
    $script:Settings.endDateTime = $dtEnd.Value.ToString("yyyy-MM-dd HH:mm:ss")
    Request-SaveSettings
  })

  $cmbTable.add_SelectedIndexChanged({
    $script:Settings.selectedTableName = Get-SelectedTableName
    Request-SaveSettings
  })

  $cmbTable.add_TextChanged({
    $script:Settings.selectedTableName = Get-SelectedTableName
    Request-SaveSettings
  })

  $cmbAttachmentTable.add_SelectedIndexChanged({
    $script:Settings.attachmentSelectedTableName = Get-SelectedAttachmentTableName
    Request-SaveSettings
  })

  $cmbAttachmentTable.add_TextChanged({
    $script:Settings.attachmentSelectedTableName = Get-SelectedAttachmentTableName
    Request-SaveSettings
  })

  $cmbAttachmentDateField.add_SelectedIndexChanged({
    $script:Settings.attachmentFilterDateField = [string]$cmbAttachmentDateField.SelectedItem
    Request-SaveSettings
  })

  $dtAttachmentStart.add_ValueChanged({
    $script:Settings.attachmentStartDateTime = $dtAttachmentStart.Value.ToString('yyyy-MM-dd HH:mm:ss')
    Request-SaveSettings
  })

  $dtAttachmentEnd.add_ValueChanged({
    $script:Settings.attachmentEndDateTime = $dtAttachmentEnd.Value.ToString('yyyy-MM-dd HH:mm:ss')
    Request-SaveSettings
  })

  $txtAttachmentDir.add_TextChanged({
    $script:Settings.attachmentDownloadDirectory = $txtAttachmentDir.Text
    Request-SaveSettings
  })

  $chkAttachmentSubfolder.add_CheckedChanged({
    $script:Settings.attachmentCreateSubfolderPerTable = [bool]$chkAttachmentSubfolder.Checked
    Request-SaveSettings
  })

  $cmbDeleteTable.add_SelectedIndexChanged({
    $script:Settings.deleteTargetTable = Get-SelectedDeleteTableName
    Request-SaveSettings
    Refresh-DeleteExecuteButton
  })

  $cmbDeleteTable.add_TextChanged({
    $script:Settings.deleteTargetTable = Get-SelectedDeleteTableName
    Request-SaveSettings
    Refresh-DeleteExecuteButton
  })

  $numDeleteMaxRetries.add_ValueChanged({
    $script:Settings.deleteMaxRetries = [int]$numDeleteMaxRetries.Value
    Request-SaveSettings
  })

  $numExportMaxRows.add_ValueChanged({
    $script:Settings.exportMaxRows = [int]$numExportMaxRows.Value
    Request-SaveSettings
  })

  $txtDir.add_TextChanged({
    $script:Settings.exportDirectory = $txtDir.Text
    Request-SaveSettings
  })

  $txtLogDir.add_TextChanged({
    $script:Settings.logOutputDirectory = $txtLogDir.Text
    Request-SaveSettings
  })

  $tabLogs.add_Enter({
    Scroll-LogsToBottom
  })

  $script:txtLog.add_TextChanged({
    if ($chkLogAutoScroll.Checked) { Scroll-LogsToBottom }
  })

  $txtViewName.add_TextChanged({
    $script:Settings.viewEditorViewName = $txtViewName.Text
    Request-SaveSettings
  })

  $txtViewLabel.add_TextChanged({
    $script:Settings.viewEditorViewLabel = $txtViewLabel.Text
    Request-SaveSettings
  })

  $txtBasePrefix.add_TextChanged({
    $script:Settings.viewEditorBasePrefix = $txtBasePrefix.Text
    Request-SaveSettings
    Update-ViewEditorColumnChoices
  })

  $cmbBaseTable.add_SelectedIndexChanged({
    $script:Settings.viewEditorBaseTable = Get-SelectedBaseTableName
    Request-SaveSettings
    for ($i = 0; $i -lt $gridJoins.Rows.Count; $i++) {
      Populate-JoinColumnsForRow $i
    }
    Update-ViewEditorColumnChoices
  })

  $cmbBaseTable.add_TextChanged({
    $script:Settings.viewEditorBaseTable = Get-SelectedBaseTableName
    Request-SaveSettings
    for ($i = 0; $i -lt $gridJoins.Rows.Count; $i++) {
      Populate-JoinColumnsForRow $i
    }
    Update-ViewEditorColumnChoices
  })

  $btnReloadColumns.add_Click({ Fetch-ColumnsForBaseTable })

  $btnAddJoin.add_Click({
    $rowIndex = $gridJoins.Rows.Add()
    if ($rowIndex -ge 0) {
      $gridJoins.Rows[$rowIndex].Cells[1].Value = "__base__"
      $gridJoins.Rows[$rowIndex].Cells[5].Value = $false
      Populate-JoinColumnsForRow $rowIndex
      $gridJoins.Rows[$rowIndex].Cells[1].Value = "__base__"
      $gridJoins.Rows[$rowIndex].Cells[4].Value = ("t{0}" -f ($rowIndex + 1))
      $gridJoins.Rows[$rowIndex].Cells[5].Value = $false
      Update-ViewEditorColumnChoices
      Save-JoinDefinitionsToSettings
    }
  })

  $btnRemoveJoin.add_Click({
    if ($gridJoins.SelectedRows.Count -gt 0) {
      $gridJoins.Rows.Remove($gridJoins.SelectedRows[0])
      Update-ViewEditorColumnChoices
      Save-JoinDefinitionsToSettings
    }
  })




  $gridJoins.add_EditingControlShowing({
    param($sender, $e)
    if ($gridJoins.CurrentCell -and ($gridJoins.CurrentCell.ColumnIndex -eq 0 -or $gridJoins.CurrentCell.ColumnIndex -eq 2 -or $gridJoins.CurrentCell.ColumnIndex -eq 3)) {
      $combo = $e.Control -as [System.Windows.Forms.ComboBox]
      if ($combo) {
        $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
        $combo.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
        $combo.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
      }
    }
  })

  $gridJoins.add_CellValueChanged({
    param($sender, $e)
    if ($e.RowIndex -ge 0) {
      if ($e.ColumnIndex -eq 0 -or $e.ColumnIndex -eq 1 -or $e.ColumnIndex -eq 4) {
        for ($i = $e.RowIndex; $i -lt $gridJoins.Rows.Count; $i++) {
          Populate-JoinColumnsForRow $i
        }
        Update-ViewEditorColumnChoices
      }
    }
    Save-JoinDefinitionsToSettings
  })
  $gridJoins.add_RowsRemoved({
    Update-ViewEditorColumnChoices
    Save-JoinDefinitionsToSettings
  })
  $gridJoins.add_CurrentCellDirtyStateChanged({
    Complete-GridCurrentEdit $gridJoins "Join"
  })


  $gridJoins.add_DataError({
    param($sender, $e)
    $e.ThrowException = $false
    Add-Log ("Join grid input error: {0}" -f $e.Exception.Message)
  })



  $btnCreateView.add_Click({ Create-DatabaseView })

  $lnkCreatedViewList.add_LinkClicked({
    param($sender, $e)
    $target = [string]$e.Link.LinkData
    if (-not [string]::IsNullOrWhiteSpace($target)) {
      Start-Process $target | Out-Null
    }
  })
  $lnkCreatedViewDefinition.add_LinkClicked({
    param($sender, $e)
    $target = [string]$e.Link.LinkData
    if (-not [string]::IsNullOrWhiteSpace($target)) {
      Start-Process $target | Out-Null
    }
  })

  $tabs.add_SelectedIndexChanged({
    if (($script:IsAttachmentHarvesterFeatureEnabled -and $tabs.SelectedTab -eq $tabAttachmentHarvester) -or ($script:IsViewEditorFeatureEnabled -and $tabs.SelectedTab -eq $tabViewEditor) -or ($script:IsDeleteFeatureEnabled -and $tabs.SelectedTab -eq $tabDelete)) {
      Ensure-TablesLoaded
    }
  })

  $cmbOutputFormat.add_SelectedIndexChanged({
    $script:Settings.outputFormat = [string]$cmbOutputFormat.SelectedItem
    Update-OutputOptionState
    if ($chkOutputBom.Enabled) {
      $script:Settings.outputBom = [bool]$chkOutputBom.Checked
    } else {
      $script:Settings.outputBom = $false
    }
    Request-SaveSettings
  })

  $chkOutputBom.add_CheckedChanged({
    if ($chkOutputBom.Enabled) {
      $script:Settings.outputBom = [bool]$chkOutputBom.Checked
    } else {
      $script:Settings.outputBom = $false
    }
    Request-SaveSettings
  })

  $btnTogglePass.add_Click({
    $txtPass.UseSystemPasswordChar = -not $txtPass.UseSystemPasswordChar
    if ($txtPass.UseSystemPasswordChar) {
      $btnTogglePass.Text = T "Show"
    } else {
      $btnTogglePass.Text = T "Hide"
    }
  })
  $btnToggleKey.add_Click({
    $txtKey.UseSystemPasswordChar = -not $txtKey.UseSystemPasswordChar
    if ($txtKey.UseSystemPasswordChar) {
      $btnToggleKey.Text = T "Show"
    } else {
      $btnToggleKey.Text = T "Hide"
    }
  })

  $btnBrowse.add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = (T "ExportDir")
    if (Test-Path $txtDir.Text) {
      $dlg.SelectedPath = $txtDir.Text
    } else {
      $dlg.SelectedPath = $DefaultExportDir
    }
    if ($dlg.ShowDialog() -eq "OK") { $txtDir.Text = $dlg.SelectedPath }
  })

  $btnAttachmentBrowse.add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = (T "AttachmentDownloadDir")
    if (Test-Path $txtAttachmentDir.Text) {
      $dlg.SelectedPath = $txtAttachmentDir.Text
    } else {
      $dlg.SelectedPath = $DefaultAttachmentDir
    }
    if ($dlg.ShowDialog() -eq "OK") { $txtAttachmentDir.Text = $dlg.SelectedPath }
  })

  $btnLogBrowse.add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = (T "LogOutputDir")
    if (Test-Path $txtLogDir.Text) {
      $dlg.SelectedPath = $txtLogDir.Text
    } else {
      $dlg.SelectedPath = $DefaultLogDir
    }
    if ($dlg.ShowDialog() -eq "OK") { $txtLogDir.Text = $dlg.SelectedPath }
  })

  $btnLogCopy.add_Click({
    if (-not [string]::IsNullOrEmpty($script:txtLog.Text)) {
      [System.Windows.Forms.Clipboard]::SetText($script:txtLog.Text)
      Add-Log (T "LogCopied")
    }
  })

  $btnLogClear.add_Click({
    $script:txtLog.Clear()
  })

  $txtLogSearch.add_TextChanged({
    $needle = ([string]$txtLogSearch.Text).Trim()
    if ([string]::IsNullOrWhiteSpace($needle)) { return }
    $idx = $script:txtLog.Text.LastIndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase)
    if ($idx -ge 0) {
      $script:txtLog.SelectionStart = $idx
      $script:txtLog.SelectionLength = $needle.Length
      $script:txtLog.ScrollToCaret()
      $script:txtLog.Focus()
    }
  })

  $cmbTheme.add_SelectedIndexChanged({
    $script:Settings.uiTheme = [string]$cmbTheme.SelectedItem
    Request-SaveSettings
    Set-Theme $script:Settings.uiTheme
    Apply-ThemeRecursive $form
    Set-ButtonStyle $btnExecute "primary"
    Set-ButtonStyle $btnAttachmentExecute "primary"
    Set-ButtonStyle $btnCreateView "primary"
    Set-ButtonStyle $btnDeleteExecute "danger"
  })

  $btnLast30Days.add_Click({
    $now = Get-Date
    $dtStart.Value = $now.AddDays(-30)
    $dtEnd.Value = $now
    $rbBetween.Checked = $true
  })

  $btnAttachmentLastRunToNow.add_Click({ Set-AttachmentLastRunRangeFromHistory })

  $btnOpenFolder.add_Click({
    $dir = Ensure-ExportDir $txtDir.Text
    Start-Process explorer.exe $dir | Out-Null
  })

  $btnReloadTables.add_Click({ Fetch-Tables })
  if ($script:IsDeleteFeatureEnabled) {
    $btnDeleteReloadTables.add_Click({ Fetch-Tables })
  }
  if ($script:IsExportFeatureEnabled) {
    $btnExecute.add_Click({ Export-Table })
  }
  if ($script:IsAttachmentHarvesterFeatureEnabled) {
    $btnAttachmentExecute.add_Click({ Harvest-Attachments })
  }
  if ($script:IsDeleteFeatureEnabled) {
    $btnDeleteExecute.add_Click({
      try {
        Remove-AllTableRecords
      } catch {
        Add-Log ("{0}: {1}" -f (T "Failed"), $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message) | Out-Null
      }
    })
  }

  # First-run export/log dir
  try { [void](Ensure-ExportDir $txtDir.Text) } catch { }
  try { [void](Ensure-LogDir $txtLogDir.Text) } catch { }
  Scroll-LogsToBottom

  $form.add_FormClosing({
    Complete-GridCurrentEdit $gridJoins "Join"
    Save-JoinDefinitionsToSettings
    $script:Settings.viewEditorSelectedColumnsJson = (@(Get-SelectedViewFieldTokens) | ConvertTo-Json -Compress)
    Request-SaveSettings -Immediate
  })

  foreach ($disabledMessage in @($script:DisabledFeatureMessages)) {
    Add-Log $disabledMessage
  }

  Apply-ThemeRecursive $form
  Set-ButtonStyle $btnExecute "primary"
  Set-ButtonStyle $btnAttachmentExecute "primary"
  Set-ButtonStyle $btnCreateView "primary"
  Set-ButtonStyle $btnDeleteExecute "danger"

  Add-Log "Ready."
  Add-Log "Notice: MIT License / https://www.ixam.net"
  Add-Log "Disclaimer: Not affiliated with or guaranteed by ServiceNow."
  [void]$form.ShowDialog()

} catch {
  try {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    [System.Windows.Forms.MessageBox]::Show($_.Exception.ToString(), "PS1SNOWUtilities Error") | Out-Null
  } catch {
    # last resort
    Write-Error $_
  }
}
