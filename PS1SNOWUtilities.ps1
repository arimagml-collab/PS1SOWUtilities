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

  # ----------------------------
  # i18n
  # ----------------------------
  $I18N = @{
    "ja" = @{
      AppTitle="PS1 SNOW Utilities"
      TabExport="Export"
      TabViewer="Viewer"
      TabSettings="設定"
      TargetTable="Target Table"
      ReloadTables="テーブル再取得"
      EasyFilter="イージーフィルタ"
      FilterAll="All"
      FilterUpdatedBetween="sys_updated_on 開始～終了"
      Start="開始"
      End="終了"
      Last30Days="過去30日"
      ExportDir="エクスポートDirectory"
      Browse="参照..."
      Execute="実行"
      OutputFormat="出力形式"
      FormatCsv="CSV"
      FormatJson="JSON"
      FormatXlsx="Excel (.xlsx)"
      Log="ログ"
      ViewerTable="閲覧テーブル"
      ViewerLoad="データ読込"
      ViewerApplyFilter="検索適用"
      ViewerClearFilter="検索クリア"
      ViewerFilter="フィールド検索"
      ViewerRows="表示件数"
      ViewerNoData="表示データがありません。先にデータを読込してください。"
      UiLang="UI言語"
      Instance="Servicenowインスタンス名"
      AuthType="認証方式"
      AuthUserPass="ユーザID＋パスワード"
      AuthApiKey="APIキー"
      UserId="ユーザID"
      Password="パスワード"
      ApiKey="APIキー"
      Show="表示"
      Hide="隠す"
      SaveHint="入力は自動保存されます（settings.json）"
      TestTablesHint="※テーブル一覧は sys_db_object を参照します（権限により取得できない場合あり）"
      WarnInstance="インスタンス名が未設定です。"
      WarnAuth="認証情報が不足しています。"
      WarnTable="テーブルが未選択です。"
      FetchingTables="テーブル一覧を取得中..."
      Exporting="エクスポート中..."
      Done="完了"
      Failed="失敗"
      OpenFolder="フォルダを開く"
      TableFetchFallback="テーブル一覧を取得できないため、Target Tableを手動入力してください。"
      CopyrightLink="Copyright (c) ixam.net"
    }
    "en" = @{
      AppTitle="PS1 SNOW Utilities"
      TabExport="Export"
      TabViewer="Viewer"
      TabSettings="Settings"
      TargetTable="Target Table"
      ReloadTables="Reload Tables"
      EasyFilter="Easy Filter"
      FilterAll="All"
      FilterUpdatedBetween="sys_updated_on Between"
      Start="Start"
      End="End"
      Last30Days="Last 30 Days"
      ExportDir="Export Directory"
      Browse="Browse..."
      Execute="Execute"
      OutputFormat="Output Format"
      FormatCsv="CSV"
      FormatJson="JSON"
      FormatXlsx="Excel (.xlsx)"
      Log="Log"
      ViewerTable="Viewer Table"
      ViewerLoad="Load Data"
      ViewerApplyFilter="Apply Search"
      ViewerClearFilter="Clear Search"
      ViewerFilter="Field Search"
      ViewerRows="Rows"
      ViewerNoData="No data loaded. Please load data first."
      UiLang="UI Language"
      Instance="ServiceNow Instance"
      AuthType="Authentication"
      AuthUserPass="User + Password"
      AuthApiKey="API Key"
      UserId="User ID"
      Password="Password"
      ApiKey="API Key"
      Show="Show"
      Hide="Hide"
      SaveHint="Inputs are auto-saved (settings.json)."
      TestTablesHint="Note: table list is read from sys_db_object (may fail depending on ACL)."
      WarnInstance="Instance is empty."
      WarnAuth="Authentication info is incomplete."
      WarnTable="No table selected."
      FetchingTables="Fetching table list..."
      Exporting="Exporting..."
      Done="Done"
      Failed="Failed"
      OpenFolder="Open Folder"
      TableFetchFallback="Could not fetch table list. Please type Target Table manually."
      CopyrightLink="Copyright (c) ixam.net"
    }
  }

  function T([string]$key) {
    $lang = "ja"
    if ($script:Settings -and $script:Settings.uiLanguage) { $lang = [string]$script:Settings.uiLanguage }
    if ($I18N.ContainsKey($lang) -and $I18N[$lang].ContainsKey($key)) { return $I18N[$lang][$key] }
    return $key
  }

  # ----------------------------
  # Secret protect/unprotect (DPAPI CurrentUser)
  # ----------------------------
  function Protect-Secret([string]$plain) {
    if ([string]::IsNullOrWhiteSpace($plain)) { return "" }
    $sec = ConvertTo-SecureString $plain -AsPlainText -Force
    return (ConvertFrom-SecureString $sec)
  }
  function Unprotect-Secret([string]$enc) {
    if ([string]::IsNullOrWhiteSpace($enc)) { return "" }
    try {
      $sec = ConvertTo-SecureString $enc
      $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
      try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
      finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    } catch {
      return ""
    }
  }

  # ----------------------------
  # Settings load/save (PSCustomObject)
  # ----------------------------
  function New-DefaultSettings {
    $o = [pscustomobject]@{
      uiLanguage = "ja"
      instanceName = ""
      authType = "userpass"      # userpass | apikey
      userId = ""
      passwordEnc = ""
      apiKeyEnc = ""
      exportDirectory = ""
      filterMode = "all"         # all | updated_between
      startDateTime = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss")
      endDateTime   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
      cachedTables = @()
      cachedTablesFetchedAt = ""
      selectedTableName = ""
      exportFields = ""          # optional: comma separated sysparm_fields
      pageSize = 1000
      outputFormat = "csv"       # csv | json | xlsx
    }
    return $o
  }

  function Load-Settings {
    $def = New-DefaultSettings
    if (Test-Path $SettingsPath) {
      try {
        $json = Get-Content $SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($p in $def.PSObject.Properties.Name) {
          if ($json -and ($json.PSObject.Properties.Name -contains $p) -and $null -ne $json.$p) {
            $def.$p = $json.$p
          }
        }
      } catch {
        # ignore and use default
      }
    }
    return $def
  }

  function Save-Settings {
    try {
      $out = ($script:Settings | ConvertTo-Json -Depth 8)
      Set-Content -Path $SettingsPath -Value $out -Encoding UTF8
    } catch {
      # ignore
    }
  }

  $script:Settings = Load-Settings

  # ----------------------------
  # ServiceNow REST helper
  # ----------------------------
  function UrlEncode([string]$s) {
    return [System.Uri]::EscapeDataString($s)
  }

  function Get-BaseUrl {
    $instVal = $script:Settings.instanceName
    if ($null -eq $instVal) { $instVal = "" }
    $inst = ([string]$instVal).Trim()
    if ([string]::IsNullOrWhiteSpace($inst)) { return "" }

    if ($inst -match '^https?://') { return $inst.TrimEnd('/') }
    if ($inst -match '\.service-now\.com$') { return ("https://{0}" -f $inst).TrimEnd('/') }
    return ("https://{0}.service-now.com" -f $inst).TrimEnd('/')
  }

  function New-SnowHeaders {
    $headers = @{
      "Accept" = "application/json"
      "Content-Type" = "application/json"
    }
    if ($script:Settings.authType -eq "apikey") {
      $key = Unprotect-Secret ([string]$script:Settings.apiKeyEnc
      )
      if (-not [string]::IsNullOrWhiteSpace($key)) {
        # Default: Bearer token. If your org uses another scheme, edit here.
        $headers["Authorization"] = "Bearer $key"
      }
    }
    return $headers
  }

  function Invoke-SnowGet([string]$pathAndQuery) {
    $base = Get-BaseUrl
    if ([string]::IsNullOrWhiteSpace($base)) { throw (T "WarnInstance") }

    $uri = $base + $pathAndQuery
    $headers = New-SnowHeaders

    if ($script:Settings.authType -eq "userpass") {
      $user = ([string]$script:Settings.userId).Trim()
      $pass = Unprotect-Secret ([string]$script:Settings.passwordEnc)
      if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) { throw (T "WarnAuth") }
      $sec = ConvertTo-SecureString $pass -AsPlainText -Force
      $cred = New-Object System.Management.Automation.PSCredential($user, $sec)
      return Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -Credential $cred -TimeoutSec 120
    } else {
      return Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -TimeoutSec 120
    }
  }

  # ----------------------------
  # UI helpers
  # ----------------------------
  function Add-Log([string]$msg) {
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $script:txtLog.AppendText("[$ts] $msg`r`n")
    $script:txtLog.SelectionStart = $script:txtLog.TextLength
    $script:txtLog.ScrollToCaret()
  }

  function Ensure-ExportDir([string]$dir) {
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = $DefaultExportDir }
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    return $dir
  }

  # ----------------------------
  # Build GUI
  # ----------------------------
  $form = New-Object System.Windows.Forms.Form
  $form.StartPosition = "CenterScreen"
  $form.Size = New-Object System.Drawing.Size(980, 720)
  $form.MinimumSize = New-Object System.Drawing.Size(900, 650)

  $tabs = New-Object System.Windows.Forms.TabControl
  $tabs.Dock = "Fill"

  $tabExport = New-Object System.Windows.Forms.TabPage
  $tabViewer = New-Object System.Windows.Forms.TabPage
  $tabSettings = New-Object System.Windows.Forms.TabPage

  [void]$tabs.TabPages.Add($tabExport)
  [void]$tabs.TabPages.Add($tabViewer)
  [void]$tabs.TabPages.Add($tabSettings)
  $form.Controls.Add($tabs)

  # --- Export tab layout
  $panelExport = New-Object System.Windows.Forms.Panel
  $panelExport.Dock = "Fill"
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
  $lblStart.Location = New-Object System.Drawing.Point(160, 95)
  $lblStart.AutoSize = $true

  $dtStart = New-Object System.Windows.Forms.DateTimePicker
  $dtStart.Location = New-Object System.Drawing.Point(210, 92)
  $dtStart.Size = New-Object System.Drawing.Size(250, 28)
  $dtStart.Format = "Custom"
  $dtStart.CustomFormat = "yyyy-MM-dd HH:mm:ss"
  $dtStart.ShowUpDown = $true

  $lblEnd = New-Object System.Windows.Forms.Label
  $lblEnd.Location = New-Object System.Drawing.Point(480, 95)
  $lblEnd.AutoSize = $true

  $dtEnd = New-Object System.Windows.Forms.DateTimePicker
  $dtEnd.Location = New-Object System.Drawing.Point(525, 92)
  $dtEnd.Size = New-Object System.Drawing.Size(200, 28)
  $dtEnd.Format = "Custom"
  $dtEnd.CustomFormat = "yyyy-MM-dd HH:mm:ss"
  $dtEnd.ShowUpDown = $true

  $btnLast30Days = New-Object System.Windows.Forms.Button
  $btnLast30Days.Location = New-Object System.Drawing.Point(740, 90)
  $btnLast30Days.Size = New-Object System.Drawing.Size(180, 32)

  $lblDir = New-Object System.Windows.Forms.Label
  $lblDir.Location = New-Object System.Drawing.Point(20, 140)
  $lblDir.AutoSize = $true

  $txtDir = New-Object System.Windows.Forms.TextBox
  $txtDir.Location = New-Object System.Drawing.Point(160, 136)
  $txtDir.Size = New-Object System.Drawing.Size(560, 28)

  $btnBrowse = New-Object System.Windows.Forms.Button
  $btnBrowse.Location = New-Object System.Drawing.Point(740, 134)
  $btnBrowse.Size = New-Object System.Drawing.Size(180, 32)

  $lblOutputFormat = New-Object System.Windows.Forms.Label
  $lblOutputFormat.Location = New-Object System.Drawing.Point(20, 184)
  $lblOutputFormat.AutoSize = $true

  $cmbOutputFormat = New-Object System.Windows.Forms.ComboBox
  $cmbOutputFormat.Location = New-Object System.Drawing.Point(160, 180)
  $cmbOutputFormat.Size = New-Object System.Drawing.Size(220, 28)
  $cmbOutputFormat.DropDownStyle = "DropDownList"
  [void]$cmbOutputFormat.Items.Add("csv")
  [void]$cmbOutputFormat.Items.Add("json")
  [void]$cmbOutputFormat.Items.Add("xlsx")

  $btnExecute = New-Object System.Windows.Forms.Button
  $btnExecute.Location = New-Object System.Drawing.Point(740, 180)
  $btnExecute.Size = New-Object System.Drawing.Size(180, 42)

  $btnOpenFolder = New-Object System.Windows.Forms.Button
  $btnOpenFolder.Location = New-Object System.Drawing.Point(540, 220)
  $btnOpenFolder.Size = New-Object System.Drawing.Size(180, 42)

  $grpLog = New-Object System.Windows.Forms.GroupBox
  $grpLog.Location = New-Object System.Drawing.Point(20, 275)
  $grpLog.Size = New-Object System.Drawing.Size(900, 360)

  $script:txtLog = New-Object System.Windows.Forms.TextBox
  $script:txtLog.Multiline = $true
  $script:txtLog.ScrollBars = "Vertical"
  $script:txtLog.Dock = "Fill"
  $script:txtLog.ReadOnly = $true
  $grpLog.Controls.Add($script:txtLog)

  $panelExport.Controls.AddRange(@(
    $lblTable, $cmbTable, $btnReloadTables,
    $lblFilter, $rbAll, $rbBetween,
    $lblStart, $dtStart, $lblEnd, $dtEnd, $btnLast30Days,
    $lblDir, $txtDir, $btnBrowse,
    $lblOutputFormat, $cmbOutputFormat,
    $btnOpenFolder, $btnExecute,
    $grpLog
  ))

  # --- Viewer tab layout
  $panelViewer = New-Object System.Windows.Forms.Panel
  $panelViewer.Dock = "Fill"
  $tabViewer.Controls.Add($panelViewer)

  $lblViewerTable = New-Object System.Windows.Forms.Label
  $lblViewerTable.Location = New-Object System.Drawing.Point(20, 20)
  $lblViewerTable.AutoSize = $true

  $cmbViewerTable = New-Object System.Windows.Forms.ComboBox
  $cmbViewerTable.Location = New-Object System.Drawing.Point(160, 16)
  $cmbViewerTable.Size = New-Object System.Drawing.Size(420, 28)
  $cmbViewerTable.DropDownStyle = "DropDown"

  $lblViewerRows = New-Object System.Windows.Forms.Label
  $lblViewerRows.Location = New-Object System.Drawing.Point(600, 20)
  $lblViewerRows.AutoSize = $true

  $numViewerRows = New-Object System.Windows.Forms.NumericUpDown
  $numViewerRows.Location = New-Object System.Drawing.Point(660, 16)
  $numViewerRows.Size = New-Object System.Drawing.Size(100, 28)
  $numViewerRows.Minimum = 10
  $numViewerRows.Maximum = 5000
  $numViewerRows.Value = 300

  $btnViewerLoad = New-Object System.Windows.Forms.Button
  $btnViewerLoad.Location = New-Object System.Drawing.Point(780, 14)
  $btnViewerLoad.Size = New-Object System.Drawing.Size(140, 32)

  $grpViewerFilter = New-Object System.Windows.Forms.GroupBox
  $grpViewerFilter.Location = New-Object System.Drawing.Point(20, 60)
  $grpViewerFilter.Size = New-Object System.Drawing.Size(900, 180)

  $flowViewerFilter = New-Object System.Windows.Forms.FlowLayoutPanel
  $flowViewerFilter.Dock = "Fill"
  $flowViewerFilter.AutoScroll = $true
  $flowViewerFilter.WrapContents = $true
  $flowViewerFilter.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $grpViewerFilter.Controls.Add($flowViewerFilter)

  $btnViewerApplyFilter = New-Object System.Windows.Forms.Button
  $btnViewerApplyFilter.Location = New-Object System.Drawing.Point(740, 246)
  $btnViewerApplyFilter.Size = New-Object System.Drawing.Size(85, 32)

  $btnViewerClearFilter = New-Object System.Windows.Forms.Button
  $btnViewerClearFilter.Location = New-Object System.Drawing.Point(835, 246)
  $btnViewerClearFilter.Size = New-Object System.Drawing.Size(85, 32)

  $gridViewer = New-Object System.Windows.Forms.DataGridView
  $gridViewer.Location = New-Object System.Drawing.Point(20, 286)
  $gridViewer.Size = New-Object System.Drawing.Size(900, 350)
  $gridViewer.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
  $gridViewer.ReadOnly = $true
  $gridViewer.AllowUserToAddRows = $false
  $gridViewer.AllowUserToDeleteRows = $false
  $gridViewer.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
  $gridViewer.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells

  $panelViewer.Controls.AddRange(@(
    $lblViewerTable, $cmbViewerTable,
    $lblViewerRows, $numViewerRows, $btnViewerLoad,
    $grpViewerFilter, $btnViewerApplyFilter, $btnViewerClearFilter,
    $gridViewer
  ))

  $script:viewerDataRows = @()
  $script:viewerFilterInputs = @{}
  $script:isSyncingTableSelection = $false

  function Get-ViewerSelectedTableName {
    $text = ""
    if ($cmbViewerTable.SelectedItem) {
      $text = [string]$cmbViewerTable.SelectedItem
    } else {
      $text = [string]$cmbViewerTable.Text
    }
    $idx = $text.IndexOf(" - ")
    if ($idx -gt 0) { return $text.Substring(0, $idx).Trim() }
    return $text.Trim()
  }

  function Get-ViewerRowsAsDataTable([object[]]$rows) {
    $dt = New-Object System.Data.DataTable
    if ($null -eq $rows -or $rows.Count -eq 0) { return $dt }

    $colNameSet = New-Object System.Collections.Generic.HashSet[string]
    foreach ($obj in $rows) {
      foreach ($p in $obj.PSObject.Properties) { [void]$colNameSet.Add($p.Name) }
    }
    $cols = @($colNameSet) | Sort-Object

    foreach ($c in $cols) { [void]$dt.Columns.Add([string]$c, [string]) }

    foreach ($obj in $rows) {
      $dr = $dt.NewRow()
      foreach ($c in $cols) {
        $val = $null
        try { $val = $obj.$c } catch { $val = $null }
        if ($null -eq $val) { $dr[$c] = "" }
        else { $dr[$c] = [string]$val }
      }
      [void]$dt.Rows.Add($dr)
    }
    return $dt
  }

  function Rebuild-ViewerFilterControls {
    $flowViewerFilter.SuspendLayout()
    $flowViewerFilter.Controls.Clear()
    $script:viewerFilterInputs = @{}

    if ($script:viewerDataRows.Count -eq 0) {
      $flowViewerFilter.ResumeLayout()
      return
    }

    $colNameSet = New-Object System.Collections.Generic.HashSet[string]
    foreach ($obj in $script:viewerDataRows) {
      foreach ($p in $obj.PSObject.Properties) { [void]$colNameSet.Add($p.Name) }
    }
    $cols = @($colNameSet) | Sort-Object

    foreach ($c in $cols) {
      $holder = New-Object System.Windows.Forms.Panel
      $holder.Width = 280
      $holder.Height = 56

      $lbl = New-Object System.Windows.Forms.Label
      $lbl.Text = [string]$c
      $lbl.Location = New-Object System.Drawing.Point(0, 0)
      $lbl.AutoSize = $true

      $txt = New-Object System.Windows.Forms.TextBox
      $txt.Width = 270
      $txt.Location = New-Object System.Drawing.Point(0, 24)

      $holder.Controls.Add($lbl)
      $holder.Controls.Add($txt)
      [void]$flowViewerFilter.Controls.Add($holder)
      $script:viewerFilterInputs[$c] = $txt
    }

    $flowViewerFilter.ResumeLayout()
  }

  function Apply-ViewerFilters {
    if ($script:viewerDataRows.Count -eq 0) {
      Add-Log (T "ViewerNoData")
      return
    }

    $filtered = @($script:viewerDataRows)
    foreach ($k in $script:viewerFilterInputs.Keys) {
      $txtBox = $script:viewerFilterInputs[$k]
      $needle = ([string]$txtBox.Text).Trim()
      if ([string]::IsNullOrWhiteSpace($needle)) { continue }

      $filtered = @($filtered | Where-Object {
        $v = ""
        try { $v = [string]$_.($k) } catch { $v = "" }
        $v.IndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
      })
    }

    $gridViewer.DataSource = $null
    $gridViewer.DataSource = Get-ViewerRowsAsDataTable $filtered
    Add-Log ("viewer filtered rows={0}" -f $filtered.Count)
  }

  function Clear-ViewerFilters {
    foreach ($k in $script:viewerFilterInputs.Keys) {
      $script:viewerFilterInputs[$k].Text = ""
    }
    $gridViewer.DataSource = $null
    $gridViewer.DataSource = Get-ViewerRowsAsDataTable $script:viewerDataRows
    Add-Log ("viewer rows={0}" -f $script:viewerDataRows.Count)
  }

  function Load-ViewerData {
    $table = Get-ViewerSelectedTableName
    if ([string]::IsNullOrWhiteSpace((Get-BaseUrl))) {
      [System.Windows.Forms.MessageBox]::Show((T "WarnInstance")) | Out-Null
      return
    }
    if ([string]::IsNullOrWhiteSpace($table)) {
      [System.Windows.Forms.MessageBox]::Show((T "WarnTable")) | Out-Null
      return
    }

    try {
      $limit = [int]$numViewerRows.Value
      Add-Log ("viewer load: table={0}, limit={1}" -f $table, $limit)
      $path = "/api/now/table/{0}?sysparm_limit={1}&sysparm_display_value=true&sysparm_exclude_reference_link=true" -f $table, $limit
      $res = Invoke-SnowGet $path

      $results = $null
      if ($res -and ($res.PSObject.Properties.Name -contains "result")) { $results = $res.result }
      if ($null -eq $results) { $results = @() }

      $script:viewerDataRows = @($results)
      Rebuild-ViewerFilterControls
      $gridViewer.DataSource = $null
      $gridViewer.DataSource = Get-ViewerRowsAsDataTable $script:viewerDataRows
      Add-Log ("viewer loaded rows={0}" -f $script:viewerDataRows.Count)
    } catch {
      Add-Log ("{0}: {1}" -f (T "Failed"), $_.Exception.Message)
      [System.Windows.Forms.MessageBox]::Show( ("{0}`r`n{1}" -f (T "Failed"), $_.Exception.Message) ) | Out-Null
    }
  }

  # --- Settings tab layout
  $panelSettings = New-Object System.Windows.Forms.Panel
  $panelSettings.Dock = "Fill"
  $tabSettings.Controls.Add($panelSettings)

  $lblUiLang = New-Object System.Windows.Forms.Label
  $lblUiLang.Location = New-Object System.Drawing.Point(20, 20)
  $lblUiLang.AutoSize = $true

  $cmbLang = New-Object System.Windows.Forms.ComboBox
  $cmbLang.Location = New-Object System.Drawing.Point(220, 16)
  $cmbLang.Size = New-Object System.Drawing.Size(220, 28)
  $cmbLang.DropDownStyle = "DropDownList"
  [void]$cmbLang.Items.Add("ja")
  [void]$cmbLang.Items.Add("en")

  $lblInstance = New-Object System.Windows.Forms.Label
  $lblInstance.Location = New-Object System.Drawing.Point(20, 60)
  $lblInstance.AutoSize = $true

  $txtInstance = New-Object System.Windows.Forms.TextBox
  $txtInstance.Location = New-Object System.Drawing.Point(220, 56)
  $txtInstance.Size = New-Object System.Drawing.Size(500, 28)

  $lblAuthType = New-Object System.Windows.Forms.Label
  $lblAuthType.Location = New-Object System.Drawing.Point(20, 105)
  $lblAuthType.AutoSize = $true

  $rbUserPass = New-Object System.Windows.Forms.RadioButton
  $rbUserPass.Location = New-Object System.Drawing.Point(220, 103)
  $rbUserPass.AutoSize = $true

  $rbApiKey = New-Object System.Windows.Forms.RadioButton
  $rbApiKey.Location = New-Object System.Drawing.Point(420, 103)
  $rbApiKey.AutoSize = $true

  $lblUser = New-Object System.Windows.Forms.Label
  $lblUser.Location = New-Object System.Drawing.Point(20, 150)
  $lblUser.AutoSize = $true

  $txtUser = New-Object System.Windows.Forms.TextBox
  $txtUser.Location = New-Object System.Drawing.Point(220, 146)
  $txtUser.Size = New-Object System.Drawing.Size(260, 28)

  $lblPass = New-Object System.Windows.Forms.Label
  $lblPass.Location = New-Object System.Drawing.Point(20, 190)
  $lblPass.AutoSize = $true

  $txtPass = New-Object System.Windows.Forms.TextBox
  $txtPass.Location = New-Object System.Drawing.Point(220, 186)
  $txtPass.Size = New-Object System.Drawing.Size(360, 28)
  $txtPass.UseSystemPasswordChar = $true

  $btnTogglePass = New-Object System.Windows.Forms.Button
  $btnTogglePass.Location = New-Object System.Drawing.Point(600, 184)
  $btnTogglePass.Size = New-Object System.Drawing.Size(120, 32)

  $lblKey = New-Object System.Windows.Forms.Label
  $lblKey.Location = New-Object System.Drawing.Point(20, 230)
  $lblKey.AutoSize = $true

  $txtKey = New-Object System.Windows.Forms.TextBox
  $txtKey.Location = New-Object System.Drawing.Point(220, 226)
  $txtKey.Size = New-Object System.Drawing.Size(360, 28)
  $txtKey.UseSystemPasswordChar = $true

  $btnToggleKey = New-Object System.Windows.Forms.Button
  $btnToggleKey.Location = New-Object System.Drawing.Point(600, 224)
  $btnToggleKey.Size = New-Object System.Drawing.Size(120, 32)

  $lblSaveHint = New-Object System.Windows.Forms.Label
  $lblSaveHint.Location = New-Object System.Drawing.Point(20, 285)
  $lblSaveHint.AutoSize = $true
  $lblSaveHint.ForeColor = [System.Drawing.Color]::FromArgb(70,70,70)

  $lblTablesHint = New-Object System.Windows.Forms.Label
  $lblTablesHint.Location = New-Object System.Drawing.Point(20, 315)
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
    $lblUiLang, $cmbLang,
    $lblInstance, $txtInstance,
    $lblAuthType, $rbUserPass, $rbApiKey,
    $lblUser, $txtUser,
    $lblPass, $txtPass, $btnTogglePass,
    $lblKey,  $txtKey,  $btnToggleKey,
    $lblSaveHint, $lblTablesHint,
    $lnkCopyright
  ))

  function Apply-Language {
    $form.Text = T "AppTitle"
    $tabExport.Text = T "TabExport"
    $tabViewer.Text = T "TabViewer"
    $tabSettings.Text = T "TabSettings"

    $lblTable.Text = T "TargetTable"
    $btnReloadTables.Text = T "ReloadTables"
    $lblFilter.Text = T "EasyFilter"
    $rbAll.Text = T "FilterAll"
    $rbBetween.Text = T "FilterUpdatedBetween"
    $lblStart.Text = T "Start"
    $lblEnd.Text = T "End"
    $btnLast30Days.Text = T "Last30Days"
    $lblDir.Text = T "ExportDir"
    $btnBrowse.Text = T "Browse"
    $btnExecute.Text = T "Execute"
    $lblOutputFormat.Text = T "OutputFormat"
    $grpLog.Text = T "Log"
    $btnOpenFolder.Text = T "OpenFolder"

    $lblViewerTable.Text = T "ViewerTable"
    $lblViewerRows.Text = T "ViewerRows"
    $btnViewerLoad.Text = T "ViewerLoad"
    $grpViewerFilter.Text = T "ViewerFilter"
    $btnViewerApplyFilter.Text = T "ViewerApplyFilter"
    $btnViewerClearFilter.Text = T "ViewerClearFilter"

    $lblUiLang.Text = T "UiLang"
    $lblInstance.Text = T "Instance"
    $lblAuthType.Text = T "AuthType"
    $rbUserPass.Text = T "AuthUserPass"
    $rbApiKey.Text = T "AuthApiKey"
    $lblUser.Text = T "UserId"
    $lblPass.Text = T "Password"
    $lblKey.Text  = T "ApiKey"
    $btnTogglePass.Text = if ($txtPass.UseSystemPasswordChar) { T "Show" } else { T "Hide" }
    $btnToggleKey.Text  = if ($txtKey.UseSystemPasswordChar)  { T "Show" } else { T "Hide" }

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

      $results = $null
      if ($res -and ($res.PSObject.Properties.Name -contains "result")) { $results = $res.result }
      if ($null -eq $results) { $results = @() }

      $list = @()
      foreach ($r in $results) {
        $name = $r.name
        $label = $r.label
        if (-not [string]::IsNullOrWhiteSpace($name)) {
          if ([string]::IsNullOrWhiteSpace($label)) { $label = $name }
          $list += [pscustomobject]@{ name=$name; label=$label }
        }
      }

      $list = $list | Sort-Object name
      $script:Settings.cachedTables = $list
      $script:Settings.cachedTablesFetchedAt = (Get-Date).ToString("o")
      Save-Settings

      $cmbTable.BeginUpdate()
      $cmbViewerTable.BeginUpdate()
      $cmbTable.Items.Clear()
      $cmbViewerTable.Items.Clear()
      foreach ($t in $list) {
        [void]$cmbTable.Items.Add(("{0} - {1}" -f $t.name, $t.label))
        [void]$cmbViewerTable.Items.Add(("{0} - {1}" -f $t.name, $t.label))
      }
      $cmbTable.EndUpdate()
      $cmbViewerTable.EndUpdate()

      $targetName = ([string]$script:Settings.selectedTableName).Trim()
      if (-not [string]::IsNullOrWhiteSpace($targetName)) {
        $candidate = $null
        foreach ($item in $cmbTable.Items) {
          $itemText = [string]$item
          if ($itemText.StartsWith($targetName + " - ")) {
            $candidate = $item
            break
          }
        }
        if ($candidate) {
          $cmbTable.SelectedItem = $candidate
          $cmbViewerTable.SelectedItem = $candidate
        } else {
          $cmbTable.Text = $targetName
          $cmbViewerTable.Text = $targetName
        }
      }

      Add-Log ("{0}: {1}" -f (T "Done"), $list.Count)
    } catch {
      Add-Log ("{0}: {1}" -f (T "Failed"), $_.Exception.Message)
      Add-Log (T "TableFetchFallback")
      $cmbTable.DroppedDown = $false
      $cmbTable.Select()
    }
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

  function Export-Table {
    $table = Get-SelectedTableName

    if ([string]::IsNullOrWhiteSpace((Get-BaseUrl))) {
      [System.Windows.Forms.MessageBox]::Show((T "WarnInstance")) | Out-Null
      return
    }
    if ([string]::IsNullOrWhiteSpace($table)) {
      [System.Windows.Forms.MessageBox]::Show((T "WarnTable")) | Out-Null
      return
    }

    if ($script:Settings.authType -eq "userpass") {
      $u = [string]$script:Settings.userId
      $p = Unprotect-Secret ([string]$script:Settings.passwordEnc)
      if ([string]::IsNullOrWhiteSpace($u) -or [string]::IsNullOrWhiteSpace($p)) {
        [System.Windows.Forms.MessageBox]::Show((T "WarnAuth")) | Out-Null
        return
      }
    } else {
      $k = Unprotect-Secret ([string]$script:Settings.apiKeyEnc)
      if ([string]::IsNullOrWhiteSpace($k)) {
        [System.Windows.Forms.MessageBox]::Show((T "WarnAuth")) | Out-Null
        return
      }
    }

    $exportDir = Ensure-ExportDir $txtDir.Text
    $script:Settings.exportDirectory = $exportDir
    Save-Settings

    $query = Build-QueryString

    $pageSizeVal = $script:Settings.pageSize
    if ($null -eq $pageSizeVal) { $pageSizeVal = 1000 }
    $pageSize = [int]$pageSizeVal
    if ($pageSize -lt 100) { $pageSize = 100 }
    if ($pageSize -gt 5000) { $pageSize = 5000 }

    Add-Log (T "Exporting")
    Add-Log ("table={0}, pageSize={1}" -f $table, $pageSize)
    Add-Log ("outputFormat={0}" -f [string]$script:Settings.outputFormat)
    if (-not [string]::IsNullOrWhiteSpace($query)) { Add-Log ("query={0}" -f $query) }

    try {
      $all = New-Object System.Collections.Generic.List[object]
      $offset = 0

      $fieldsVal = $script:Settings.exportFields
      if ($null -eq $fieldsVal) { $fieldsVal = "" }
      $fields = ([string]$fieldsVal).Trim()
      $fieldsParam = ""
      if (-not [string]::IsNullOrWhiteSpace($fields)) {
        $fieldsParam = "&sysparm_fields=" + (UrlEncode $fields)
      }

      while ($true) {
        $qs = @{
          sysparm_limit  = $pageSize
          sysparm_offset = $offset
          sysparm_display_value = "false"
          sysparm_exclude_reference_link = "true"
        }

        $queryParts = New-Object System.Collections.Generic.List[string]
        foreach ($k2 in $qs.Keys) {
          [void]$queryParts.Add(("{0}={1}" -f $k2, (UrlEncode ([string]$qs[$k2]))))
        }
        if (-not [string]::IsNullOrWhiteSpace($query)) {
          [void]$queryParts.Add(("sysparm_query={0}" -f (UrlEncode $query)))
        }

        $path = "/api/now/table/" + $table + "?" + ($queryParts -join "&") + $fieldsParam
        $res = Invoke-SnowGet $path

        $batchRes = $null
        if ($res -and ($res.PSObject.Properties.Name -contains "result")) { $batchRes = $res.result }
        if ($null -eq $batchRes) { $batchRes = @() }

        $batch = @($batchRes)
        foreach ($r in $batch) { $all.Add($r) }

        Add-Log ("fetched: offset={0}, count={1}, total={2}" -f $offset, $batch.Count, $all.Count)

        if ($batch.Count -lt $pageSize) { break }
        $offset += $pageSize
        if ($offset -gt 2000000) { break } # safety stop
      }

      if ($all.Count -eq 0) {
        Add-Log "0 records."
        [System.Windows.Forms.MessageBox]::Show("0 records.") | Out-Null
        return
      }

      $colNameSet = New-Object System.Collections.Generic.HashSet[string]
      foreach ($obj in $all) {
        foreach ($p in $obj.PSObject.Properties) { [void]$colNameSet.Add($p.Name) }
      }
      $cols = @($colNameSet) | Sort-Object


      $outRows = foreach ($obj in $all) {
        $h = [ordered]@{}
        foreach ($c in $cols) {
          $val = $null
          try { $val = $obj.$c } catch { $val = $null }
          $h[$c] = $val
        }
        [pscustomobject]$h
      }

      $stamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
      $suffix = if ($rbBetween.Checked) {
        ("_{0}-{1}" -f $dtStart.Value.ToString("yyyyMMddHHmmss"), $dtEnd.Value.ToString("yyyyMMddHHmmss"))
      } else { "" }

      $formatVal = [string]$script:Settings.outputFormat
      if ([string]::IsNullOrWhiteSpace($formatVal)) { $formatVal = "csv" }
      $format = $formatVal.Trim().ToLowerInvariant()
      if ((@("csv","json","xlsx") -notcontains $format)) { $format = "csv" }

      $ext = switch ($format) {
        "json" { "json" }
        "xlsx" { "xlsx" }
        default { "csv" }
      }

      $file = Join-Path $exportDir ("{0}{1}_{2}.{3}" -f $table, $suffix, $stamp, $ext)

      $recordCount = @($outRows).Count

      switch ($format) {
        "json" {
          $outRows | ConvertTo-Json -Depth 10 | Set-Content -Path $file -Encoding UTF8
        }
        "xlsx" {
          $excel = $null
          $workbook = $null
          $worksheet = $null
          try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Add()
            $worksheet = $workbook.Worksheets.Item(1)

            for ($i = 0; $i -lt $cols.Count; $i++) {
              $worksheet.Cells.Item(1, $i + 1) = [string]$cols[$i]
            }

            $rowIndex = 2
            foreach ($row in $outRows) {
              for ($i = 0; $i -lt $cols.Count; $i++) {
                $v = $row.($cols[$i])
                if ($null -eq $v) { $worksheet.Cells.Item($rowIndex, $i + 1) = "" }
                else { $worksheet.Cells.Item($rowIndex, $i + 1) = [string]$v }
              }
              $rowIndex++
            }

            $xlOpenXmlWorkbook = 51
            $workbook.SaveAs($file, $xlOpenXmlWorkbook)
          } finally {
            if ($workbook) { $workbook.Close($false) | Out-Null }
            if ($excel) { $excel.Quit() }
            foreach ($obj in @($worksheet, $workbook, $excel)) {
              if ($obj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
          }
        }
        default {
          $outRows | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
        }
      }
      Add-Log ("{0}: {1}" -f (T "Done"), $file)

      [System.Windows.Forms.MessageBox]::Show(("OK`r`n{0}`r`nRecords: {1}" -f $file, $recordCount)) | Out-Null
    } catch {
      Add-Log ("{0}: {1}" -f (T "Failed"), $_.Exception.Message)
      [System.Windows.Forms.MessageBox]::Show(("{0}`r`n{1}" -f (T "Failed"), $_.Exception.Message)) | Out-Null
    }
  }

  # ----------------------------
  # Initialize from settings
  # ----------------------------
  $cmbLang.SelectedItem = [string]$script:Settings.uiLanguage
  if (-not $cmbLang.SelectedItem) { $cmbLang.SelectedItem = "ja" }

  $txtInstance.Text = [string]$script:Settings.instanceName
  $txtUser.Text = [string]$script:Settings.userId

  if ([string]::IsNullOrWhiteSpace([string]$script:Settings.exportDirectory)) {
    $txtDir.Text = $DefaultExportDir
  } else {
    $txtDir.Text = [string]$script:Settings.exportDirectory
  }

  if ([string]$script:Settings.filterMode -eq "updated_between") { $rbBetween.Checked = $true } else { $rbAll.Checked = $true }

  $initialOutputFormat = ([string]$script:Settings.outputFormat).Trim().ToLowerInvariant()
  if ((@("csv","json","xlsx") -notcontains $initialOutputFormat)) { $initialOutputFormat = "csv" }
  $cmbOutputFormat.SelectedItem = $initialOutputFormat

  try { $dtStart.Value = [datetime]::Parse([string]$script:Settings.startDateTime) } catch { }
  try { $dtEnd.Value   = [datetime]::Parse([string]$script:Settings.endDateTime) } catch { }

  if ([string]$script:Settings.authType -eq "apikey") { $rbApiKey.Checked = $true } else { $rbUserPass.Checked = $true }

  $txtPass.Text = Unprotect-Secret ([string]$script:Settings.passwordEnc)
  $txtKey.Text  = Unprotect-Secret ([string]$script:Settings.apiKeyEnc)

  if ($script:Settings.cachedTables -and $script:Settings.cachedTables.Count -gt 0) {
    $cmbTable.BeginUpdate()
    $cmbViewerTable.BeginUpdate()
    $cmbTable.Items.Clear()
    $cmbViewerTable.Items.Clear()
    foreach ($t in $script:Settings.cachedTables) {
      [void]$cmbTable.Items.Add(("{0} - {1}" -f $t.name, $t.label))
      [void]$cmbViewerTable.Items.Add(("{0} - {1}" -f $t.name, $t.label))
    }
    $cmbTable.EndUpdate()
    $cmbViewerTable.EndUpdate()
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
      $cmbViewerTable.SelectedItem = $candidate
    } else {
      $cmbTable.Text = $initialTableName
      $cmbViewerTable.Text = $initialTableName
    }
  }

  Update-AuthUI
  Update-FilterUI
  Apply-Language

  # ----------------------------
  # Wire events for auto-save
  # ----------------------------
  $cmbLang.add_SelectedIndexChanged({
    $script:Settings.uiLanguage = [string]$cmbLang.SelectedItem
    Save-Settings
    Apply-Language
  })

  $txtInstance.add_TextChanged({
    $script:Settings.instanceName = $txtInstance.Text
    Save-Settings
  })

  $rbUserPass.add_CheckedChanged({
    if ($rbUserPass.Checked) {
      $script:Settings.authType = "userpass"
      Save-Settings
      Update-AuthUI
    }
  })
  $rbApiKey.add_CheckedChanged({
    if ($rbApiKey.Checked) {
      $script:Settings.authType = "apikey"
      Save-Settings
      Update-AuthUI
    }
  })

  $txtUser.add_TextChanged({
    $script:Settings.userId = $txtUser.Text
    Save-Settings
  })

  $txtPass.add_TextChanged({
    $script:Settings.passwordEnc = Protect-Secret $txtPass.Text
    Save-Settings
  })

  $txtKey.add_TextChanged({
    $script:Settings.apiKeyEnc = Protect-Secret $txtKey.Text
    Save-Settings
  })

  $rbAll.add_CheckedChanged({
    if ($rbAll.Checked) {
      $script:Settings.filterMode = "all"
      Save-Settings
      Update-FilterUI
    }
  })
  $rbBetween.add_CheckedChanged({
    if ($rbBetween.Checked) {
      $script:Settings.filterMode = "updated_between"
      Save-Settings
      Update-FilterUI
    }
  })

  $dtStart.add_ValueChanged({
    $script:Settings.startDateTime = $dtStart.Value.ToString("yyyy-MM-dd HH:mm:ss")
    Save-Settings
  })
  $dtEnd.add_ValueChanged({
    $script:Settings.endDateTime = $dtEnd.Value.ToString("yyyy-MM-dd HH:mm:ss")
    Save-Settings
  })

  $cmbTable.add_SelectedIndexChanged({
    if ($script:isSyncingTableSelection) { return }
    $script:isSyncingTableSelection = $true
    try {
      $script:Settings.selectedTableName = Get-SelectedTableName
      $cmbViewerTable.Text = $cmbTable.Text
      Save-Settings
    } finally {
      $script:isSyncingTableSelection = $false
    }
  })

  $cmbTable.add_TextChanged({
    if ($script:isSyncingTableSelection) { return }
    $script:isSyncingTableSelection = $true
    try {
      $script:Settings.selectedTableName = Get-SelectedTableName
      $cmbViewerTable.Text = $cmbTable.Text
      Save-Settings
    } finally {
      $script:isSyncingTableSelection = $false
    }
  })

  $cmbViewerTable.add_SelectedIndexChanged({
    if ($script:isSyncingTableSelection) { return }
    $script:isSyncingTableSelection = $true
    try {
      $script:Settings.selectedTableName = Get-ViewerSelectedTableName
      $cmbTable.Text = $cmbViewerTable.Text
      Save-Settings
    } finally {
      $script:isSyncingTableSelection = $false
    }
  })

  $cmbViewerTable.add_TextChanged({
    if ($script:isSyncingTableSelection) { return }
    $script:isSyncingTableSelection = $true
    try {
      $script:Settings.selectedTableName = Get-ViewerSelectedTableName
      $cmbTable.Text = $cmbViewerTable.Text
      Save-Settings
    } finally {
      $script:isSyncingTableSelection = $false
    }
  })

  $txtDir.add_TextChanged({
    $script:Settings.exportDirectory = $txtDir.Text
    Save-Settings
  })

  $cmbOutputFormat.add_SelectedIndexChanged({
    $script:Settings.outputFormat = [string]$cmbOutputFormat.SelectedItem
    Save-Settings
  })

  $btnTogglePass.add_Click({
    $txtPass.UseSystemPasswordChar = -not $txtPass.UseSystemPasswordChar
    $btnTogglePass.Text = if ($txtPass.UseSystemPasswordChar) { T "Show" } else { T "Hide" }
  })
  $btnToggleKey.add_Click({
    $txtKey.UseSystemPasswordChar = -not $txtKey.UseSystemPasswordChar
    $btnToggleKey.Text = if ($txtKey.UseSystemPasswordChar) { T "Show" } else { T "Hide" }
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

  $btnLast30Days.add_Click({
    $now = Get-Date
    $dtStart.Value = $now.AddDays(-30)
    $dtEnd.Value = $now
    $rbBetween.Checked = $true
  })

  $btnOpenFolder.add_Click({
    $dir = Ensure-ExportDir $txtDir.Text
    Start-Process explorer.exe $dir | Out-Null
  })

  $btnReloadTables.add_Click({ Fetch-Tables })
  $btnExecute.add_Click({ Export-Table })
  $btnViewerLoad.add_Click({ Load-ViewerData })
  $btnViewerApplyFilter.add_Click({ Apply-ViewerFilters })
  $btnViewerClearFilter.add_Click({ Clear-ViewerFilters })

  # First-run export dir
  try { [void](Ensure-ExportDir $txtDir.Text) } catch { }

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
