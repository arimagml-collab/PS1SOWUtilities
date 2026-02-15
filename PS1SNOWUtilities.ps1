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
      TabViewEditor="DataBase View Editor"
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
      ViewName="View内部名"
      ViewLabel="Viewラベル"
      BaseTable="ベーステーブル"
      ReloadColumns="カラム再取得"
      ViewColumns="表示カラム"
      CreateView="View作成"
      JoinDefinitions="JOIN定義"
      AddJoin="JOIN追加"
      RemoveJoin="JOIN削除"
      JoinTable="JOINテーブル"
      JoinBaseColumn="左カラム(ベース)"
      JoinTargetColumn="右カラム(JOIN先)"
      JoinSource="左側 Prefix"
      BasePrefix="ベース Prefix"
      JoinPrefix="Variable Prefix"
      LeftJoin="LEFT JOIN"
      ConditionColumn="カラム"
      ConditionOperator="演算子"
      ConditionValue="値"
      WarnViewName="View内部名を入力してください。"
      WarnViewLabel="Viewラベルを入力してください。"
      WarnBaseTable="ベーステーブルを選択してください。"
      WarnJoinTable="JOINテーブルを選択してください。"
      WarnJoinBaseColumn="JOINの左カラム(ベース)を選択してください。"
      WarnJoinTargetColumn="JOINの右カラム(JOIN先)を選択してください。"
      FetchingColumns="カラム一覧を取得中..."
      ColumnsFetched="カラム一覧を取得"
      ViewCreated="Viewを作成しました"
      ViewCreateFailed="View作成に失敗しました"
      ViewWhereFallback="この環境ではWhere句保存フィールドを特定できませんでした。Viewは作成されましたが、Where句は手動設定してください。"
      ViewJoinFallback="JOIN定義の保存に失敗しました。View本体は作成されましたが、JOINは手動設定してください。"
      CreatedViewListLink="作成したViewリスト"
      CreatedViewDefinitionLink="View定義レコード"
    }
    "en" = @{
      AppTitle="PS1 SNOW Utilities"
      TabExport="Export"
      TabViewEditor="DataBase View Editor"
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
      ViewName="View Name"
      ViewLabel="View Label"
      BaseTable="Base Table"
      ReloadColumns="Reload Columns"
      ViewColumns="Visible Columns"
      CreateView="Create View"
      JoinDefinitions="Join Definitions"
      AddJoin="Add Join"
      RemoveJoin="Remove Join"
      JoinTable="Join Table"
      JoinBaseColumn="Left Column (Base)"
      JoinTargetColumn="Right Column (Join)"
      JoinSource="Left Prefix"
      BasePrefix="Base Prefix"
      JoinPrefix="Variable Prefix"
      LeftJoin="LEFT JOIN"
      ConditionColumn="Column"
      ConditionOperator="Operator"
      ConditionValue="Value"
      WarnViewName="View name is required."
      WarnViewLabel="View label is required."
      WarnBaseTable="Base table must be selected."
      WarnJoinTable="Join table is required."
      WarnJoinBaseColumn="Left join column (base table) is required."
      WarnJoinTargetColumn="Right join column (join table) is required."
      FetchingColumns="Fetching columns..."
      ColumnsFetched="Fetched columns"
      ViewCreated="View created"
      ViewCreateFailed="Failed to create view"
      ViewWhereFallback="Could not detect a writable where-clause field in this instance. View was created, but set the where clause manually."
      ViewJoinFallback="Failed to persist join definitions. View was created, but set joins manually."
      CreatedViewListLink="Created View List"
      CreatedViewDefinitionLink="View Definition Record"
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
      viewEditorViewName = ""
      viewEditorViewLabel = ""
      viewEditorBaseTable = ""
      viewEditorBasePrefix = "t0"
      viewEditorJoinsJson = "[]"
      viewEditorSelectedColumnsJson = "[]"
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
      [System.IO.File]::WriteAllText($SettingsPath, $out, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
      # ignore
    }
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
      "Content-Type" = "application/json; charset=utf-8"
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

  function Invoke-SnowPost([string]$path, [hashtable]$body) {
    $base = Get-BaseUrl
    if ([string]::IsNullOrWhiteSpace($base)) { throw (T "WarnInstance") }

    $uri = $base + $path
    $headers = New-SnowHeaders
    $jsonBody = ($body | ConvertTo-Json -Depth 8)
    $jsonBytes = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)

    if ($script:Settings.authType -eq "userpass") {
      $user = ([string]$script:Settings.userId).Trim()
      $pass = Unprotect-Secret ([string]$script:Settings.passwordEnc)
      if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) { throw (T "WarnAuth") }
      $sec = ConvertTo-SecureString $pass -AsPlainText -Force
      $cred = New-Object System.Management.Automation.PSCredential($user, $sec)
      return Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Credential $cred -TimeoutSec 120 -Body $jsonBytes
    } else {
      return Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -TimeoutSec 120 -Body $jsonBytes
    }
  }

  function Invoke-SnowPatch([string]$path, [hashtable]$body) {
    $base = Get-BaseUrl
    if ([string]::IsNullOrWhiteSpace($base)) { throw (T "WarnInstance") }

    $uri = $base + $path
    $headers = New-SnowHeaders
    $jsonBody = ($body | ConvertTo-Json -Depth 8)
    $jsonBytes = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)

    if ($script:Settings.authType -eq "userpass") {
      $user = ([string]$script:Settings.userId).Trim()
      $pass = Unprotect-Secret ([string]$script:Settings.passwordEnc)
      if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) { throw (T "WarnAuth") }
      $sec = ConvertTo-SecureString $pass -AsPlainText -Force
      $cred = New-Object System.Management.Automation.PSCredential($user, $sec)
      return Invoke-RestMethod -Method Patch -Uri $uri -Headers $headers -Credential $cred -TimeoutSec 120 -Body $jsonBytes
    } else {
      return Invoke-RestMethod -Method Patch -Uri $uri -Headers $headers -TimeoutSec 120 -Body $jsonBytes
    }
  }

  # ----------------------------
  # UI helpers
  # ----------------------------
  function Add-Log([string]$msg) {
    if ($script:txtLog -and $script:txtLog.InvokeRequired) {
      $appendAction = [System.Action[string]]{
        param($m)
        Add-Log $m
      }
      [void]$script:txtLog.BeginInvoke($appendAction, @($msg))
      return
    }
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $script:txtLog.AppendText("[$ts] $msg`r`n")
    $script:txtLog.SelectionStart = $script:txtLog.TextLength
    $script:txtLog.ScrollToCaret()
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

  # ----------------------------
  # Build GUI
  # ----------------------------
  $form = New-Object System.Windows.Forms.Form
  $form.StartPosition = "CenterScreen"
  $form.Size = New-Object System.Drawing.Size(1120, 720)
  $form.MinimumSize = New-Object System.Drawing.Size(1040, 650)

  $tabs = New-Object System.Windows.Forms.TabControl
  $tabs.Dock = "Fill"

  $tabExport = New-Object System.Windows.Forms.TabPage
  $tabViewEditor = New-Object System.Windows.Forms.TabPage
  $tabSettings = New-Object System.Windows.Forms.TabPage

  [void]$tabs.TabPages.Add($tabExport)
  [void]$tabs.TabPages.Add($tabViewEditor)
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
  $btnOpenFolder.Location = New-Object System.Drawing.Point(740, 228)
  $btnOpenFolder.Size = New-Object System.Drawing.Size(180, 42)

  $grpLog = New-Object System.Windows.Forms.GroupBox
  $grpLog.Location = New-Object System.Drawing.Point(20, 275)
  $grpLog.Size = New-Object System.Drawing.Size(900, 360)

  $script:txtLog = New-Object System.Windows.Forms.TextBox
  $script:txtLog.Multiline = $true
  $script:txtLog.ScrollBars = "Both"
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
    $tabViewEditor.Text = T "TabViewEditor"
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
    $lblInstance.Text = T "Instance"
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

  function Refresh-BaseTableItems {
    $cmbBaseTable.BeginUpdate()
    $cmbBaseTable.Items.Clear()
    if ($script:Settings.cachedTables) {
      foreach ($t in @($script:Settings.cachedTables)) {
        [void]$cmbBaseTable.Items.Add(("{0} - {1}" -f $t.name, $t.label))
      }
    }
    $cmbBaseTable.EndUpdate()

    $colJoinTable.Items.Clear()
    if ($script:Settings.cachedTables) {
      foreach ($t in @($script:Settings.cachedTables)) {
        [void]$colJoinTable.Items.Add([string]$t.name)
      }
    }
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
        $joinTable = if ($null -eq $tableCell) { "" } else { ([string]$tableCell).Trim() }
        $baseColumn = if ($null -eq $baseCell) { "" } else { ([string]$baseCell).Trim() }
        $targetColumn = if ($null -eq $targetCell) { "" } else { ([string]$targetCell).Trim() }
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
      $joinTable = if ($null -eq $joinTableCell) { "" } else { ([string]$joinTableCell).Trim() }
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
      $joinTable = if ($null -eq $joinTableCell) { "" } else { ([string]$joinTableCell).Trim() }
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
      $joinTable = if ($null -eq $joinTableCell) { "" } else { ([string]$joinTableCell).Trim() }
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
    $joinTable = if ($null -eq $joinTableCell) { "" } else { ([string]$joinTableCell).Trim() }

    $baseColumns = @()
    $joinColumns = @()
    if (-not [string]::IsNullOrWhiteSpace($baseTable)) { $baseColumns = @(Fetch-ColumnsForTable $baseTable) }
    if (-not [string]::IsNullOrWhiteSpace($joinTable)) { $joinColumns = @(Fetch-ColumnsForTable $joinTable) }

    $baseCell = [System.Windows.Forms.DataGridViewComboBoxCell]$row.Cells[2]
    $targetCell = [System.Windows.Forms.DataGridViewComboBoxCell]$row.Cells[3]

    $selectedBase = if ($null -eq $baseCell.Value) { "" } else { [string]$baseCell.Value }
    $selectedTarget = if ($null -eq $targetCell.Value) { "" } else { [string]$targetCell.Value }

    $baseCell.Items.Clear()
    foreach ($c in $baseColumns) { [void]$baseCell.Items.Add([string]$c.name) }
    if (-not [string]::IsNullOrWhiteSpace($selectedBase) -and $baseCell.Items.Contains($selectedBase)) { $baseCell.Value = $selectedBase }
    else { $baseCell.Value = $null }

    $targetCell.Items.Clear()
    foreach ($c in $joinColumns) { [void]$targetCell.Items.Add([string]$c.name) }
    if (-not [string]::IsNullOrWhiteSpace($selectedTarget) -and $targetCell.Items.Contains($selectedTarget)) { $targetCell.Value = $selectedTarget }
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

    if ([string]::IsNullOrWhiteSpace($viewName)) { [System.Windows.Forms.MessageBox]::Show((T "WarnViewName")) | Out-Null; return }
    if ([string]::IsNullOrWhiteSpace($viewLabel)) { [System.Windows.Forms.MessageBox]::Show((T "WarnViewLabel")) | Out-Null; return }
    if ([string]::IsNullOrWhiteSpace($baseTable)) { [System.Windows.Forms.MessageBox]::Show((T "WarnBaseTable")) | Out-Null; return }

    $joinDefs = @(Get-JoinDefinitions)
    foreach ($j in $joinDefs) {
      if ([string]::IsNullOrWhiteSpace([string]$j.joinTable)) { [System.Windows.Forms.MessageBox]::Show((T "WarnJoinTable")) | Out-Null; return }
      if ([string]::IsNullOrWhiteSpace([string]$j.baseColumn)) { [System.Windows.Forms.MessageBox]::Show((T "WarnJoinBaseColumn")) | Out-Null; return }
      if ([string]::IsNullOrWhiteSpace([string]$j.targetColumn)) { [System.Windows.Forms.MessageBox]::Show((T "WarnJoinTargetColumn")) | Out-Null; return }
    }

    $basePrefix = ([string]$txtBasePrefix.Text).Trim()
    if ([string]::IsNullOrWhiteSpace($basePrefix)) { $basePrefix = "t0" }
    $selectedColumns = @(Get-SelectedViewFieldTokens)

    Add-Log ("Creating DB view: {0}, base={1}, joins={2}" -f $viewName, $baseTable, $joinDefs.Count)
    $ctx = [pscustomobject]@{ viewName=$viewName; viewLabel=$viewLabel; baseTable=$baseTable; basePrefix=$basePrefix; selectedColumns=@($selectedColumns); joinDefs=@($joinDefs) }
    Invoke-Async "Create-DatabaseView" {
      param($state)
      $body = @{ name = $state.viewName; table = $state.baseTable }
      if (@($state.selectedColumns).Count -gt 0) { $body["view_field_list"] = (@($state.selectedColumns) -join ",") }
      $createRes = Invoke-SnowPost "/api/now/table/sys_db_view" $body
      $created = if ($createRes -and ($createRes.PSObject.Properties.Name -contains "result")) { $createRes.result } else { $null }
      $sysId = if ($created) { [string]$created.sys_id } else { "" }

      $joinsSaved = $true
      if (-not [string]::IsNullOrWhiteSpace($sysId)) {
        [void](Invoke-SnowPatch ("/api/now/table/sys_db_view/{0}" -f $sysId) @{ label = $state.viewLabel })

        if (@($state.selectedColumns).Count -gt 0) {
          $fieldCsv = (@($state.selectedColumns) -join ",")
          foreach ($fieldKey in @("view_fields", "field_names", "view_field_list")) {
            try { [void](Invoke-SnowPatch ("/api/now/table/sys_db_view/{0}" -f $sysId) @{ $fieldKey = $fieldCsv }); break } catch {}
          }
        }

        $baseTableRowId = ""
        try {
          $query = "view={0}^table={1}" -f $sysId, $state.baseTable
          $path = "/api/now/table/sys_db_view_table?sysparm_fields=sys_id&sysparm_limit=1&sysparm_query={0}" -f (UrlEncode $query)
          $baseTableRes = Invoke-SnowGet $path
          $baseTableRow = if ($baseTableRes -and ($baseTableRes.PSObject.Properties.Name -contains "result") -and @($baseTableRes.result).Count -gt 0) { $baseTableRes.result[0] } else { $null }
          if ($baseTableRow) { $baseTableRowId = [string]$baseTableRow.sys_id }
        } catch {}

        if ([string]::IsNullOrWhiteSpace($baseTableRowId)) {
          try {
            $baseCreate = Invoke-SnowPost "/api/now/table/sys_db_view_table" @{ view = $sysId; table = $state.baseTable; order = 0; variable_prefix = $state.basePrefix }
            if ($baseCreate -and ($baseCreate.PSObject.Properties.Name -contains "result") -and $baseCreate.result) { $baseTableRowId = [string]$baseCreate.result.sys_id }
          } catch {}
        }

        [void](Save-ViewTableMetadata $baseTableRowId $state.basePrefix "" $false $false)

        if (@($state.joinDefs).Count -gt 0) {
          $joinsSaved = $false
          $joinIndex = 1
          foreach ($joinDef in @($state.joinDefs)) {
            $joinPrefix = ([string]$joinDef.joinPrefix).Trim()
            if ([string]::IsNullOrWhiteSpace($joinPrefix)) { $joinPrefix = ("t{0}" -f $joinIndex) }
            $joinSource = ([string]$joinDef.joinSource).Trim()
            $leftPrefix = if ([string]::IsNullOrWhiteSpace($joinSource) -or $joinSource -eq "__base__") { $state.basePrefix } else { $joinSource }
            $isLeftJoin = $false
            if ($joinDef.PSObject.Properties.Name -contains "leftJoin") { $isLeftJoin = [System.Convert]::ToBoolean($joinDef.leftJoin) }
            $joinWhereClause = Build-JoinWhereClause $leftPrefix ([string]$joinDef.baseColumn) $joinPrefix ([string]$joinDef.targetColumn)
            $joinOrder = $joinIndex * 100
            $joinCreate = Try-CreateViewJoinRow $sysId $joinDef $joinWhereClause $joinPrefix $isLeftJoin $joinOrder
            if (-not [bool]$joinCreate.saved) { $joinsSaved = $false; break }
            if (-not [string]::IsNullOrWhiteSpace([string]$joinCreate.rowId)) { [void](Save-ViewTableMetadata ([string]$joinCreate.rowId) $joinPrefix $joinWhereClause $isLeftJoin $true) }
            $joinIndex++
            $joinsSaved = $true
          }
        }
      }

      return [pscustomobject]@{ viewName=$state.viewName; sysId=$sysId; joinsSaved=$joinsSaved }
    } {
      param($result)
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
    Request-SaveSettings

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

    $fieldsVal = $script:Settings.exportFields
    if ($null -eq $fieldsVal) { $fieldsVal = "" }
    $fields = ([string]$fieldsVal).Trim()

    $formatVal = [string]$script:Settings.outputFormat
    if ([string]::IsNullOrWhiteSpace($formatVal)) { $formatVal = "csv" }
    $format = $formatVal.Trim().ToLowerInvariant()
    if ((@("csv","json","xlsx") -notcontains $format)) { $format = "csv" }

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

    $ctx = [pscustomobject]@{ table=$table; pageSize=$pageSize; query=$query; fields=$fields; format=$format; file=$file }

    Invoke-Async "Export-Table" {
      param($state)
      $offset = 0
      $total = 0
      $isFirstJson = $true
      $jsonWriter = $null
      $csvWriter = $null
      $all = New-Object System.Collections.Generic.List[object]

      try {
        if ($state.format -eq "json") {
          $jsonWriter = New-Object System.IO.StreamWriter($state.file, $false, (New-Object System.Text.UTF8Encoding($false)))
          $jsonWriter.Write("[")
        } elseif ($state.format -eq "csv") {
          $csvWriter = New-Object System.IO.StreamWriter($state.file, $false, (New-Object System.Text.UTF8Encoding($false)))
              $csvWriter.WriteLine(("`"{0}`"" -f $itemJson))
        }

        while ($true) {
          $qs = @{
            sysparm_limit  = $state.pageSize
            sysparm_offset = $offset
            sysparm_display_value = "false"
            sysparm_exclude_reference_link = "true"
          }
          $queryParts = New-Object System.Collections.Generic.List[string]
          foreach ($k2 in $qs.Keys) { [void]$queryParts.Add(("{0}={1}" -f $k2, (UrlEncode ([string]$qs[$k2])))) }
          if (-not [string]::IsNullOrWhiteSpace([string]$state.query)) { [void]$queryParts.Add(("sysparm_query={0}" -f (UrlEncode ([string]$state.query)))) }
          if (-not [string]::IsNullOrWhiteSpace([string]$state.fields)) { [void]$queryParts.Add(("sysparm_fields={0}" -f (UrlEncode ([string]$state.fields)))) }

          $path = "/api/now/table/" + $state.table + "?" + ($queryParts -join "&")
          $res = Invoke-SnowGet $path
          $batchRes = if ($res -and ($res.PSObject.Properties.Name -contains "result")) { $res.result } else { @() }
          $batch = @($batchRes)

          foreach ($r in $batch) {
            if ($state.format -eq "json") {
              $itemJson = ($r | ConvertTo-Json -Depth 10 -Compress)
              if (-not $isFirstJson) { $jsonWriter.Write(",") }
              $jsonWriter.Write($itemJson)
              $isFirstJson = $false
            } elseif ($state.format -eq "csv") {
              $itemJson = ($r | ConvertTo-Json -Depth 10 -Compress).Replace('"','""')
              $csvWriter.WriteLine(("`"{0}`"" -f $itemJson))
            } else {
              $all.Add($r)
            }
          }

          $total += $batch.Count
          if ($batch.Count -lt $state.pageSize) { break }
          $offset += $state.pageSize
          if ($offset -gt 2000000) { break }
        }

        if ($state.format -eq "xlsx") {
          if ($all.Count -gt 0) {
            $colNameSet = New-Object System.Collections.Generic.HashSet[string]
            foreach ($obj in $all) { foreach ($p in $obj.PSObject.Properties) { [void]$colNameSet.Add($p.Name) } }
            $cols = @($colNameSet) | Sort-Object
            $outRows = foreach ($obj in $all) {
              $h = [ordered]@{}
              foreach ($c in $cols) { try { $h[$c] = $obj.$c } catch { $h[$c] = $null } }
              [pscustomobject]$h
            }
            $excel = $null; $workbook = $null; $worksheet = $null
            try {
              $excel = New-Object -ComObject Excel.Application
              $excel.Visible = $false
              $excel.DisplayAlerts = $false
              $workbook = $excel.Workbooks.Add()
              $worksheet = $workbook.Worksheets.Item(1)
              for ($i = 0; $i -lt $cols.Count; $i++) { $worksheet.Cells.Item(1, $i + 1) = [string]$cols[$i] }
              $rowIndex = 2
              foreach ($row in $outRows) {
                for ($i = 0; $i -lt $cols.Count; $i++) {
                  $v = $row.($cols[$i])
                  if ($null -eq $v) { $worksheet.Cells.Item($rowIndex, $i + 1) = "" } else { $worksheet.Cells.Item($rowIndex, $i + 1) = [string]$v }
                }
                $rowIndex++
              }
              $workbook.SaveAs($state.file, 51)
            } finally {
              if ($workbook) { $workbook.Close($false) | Out-Null }
              if ($excel) { $excel.Quit() }
              foreach ($obj in @($worksheet, $workbook, $excel)) { if ($obj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } }
              [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            }
          }
        }
      } finally {
        if ($jsonWriter) { $jsonWriter.Write("]"); $jsonWriter.Dispose() }
        if ($csvWriter) { $csvWriter.Dispose() }
      }

      return [pscustomobject]@{ file=$state.file; total=$total }
    } {
      param($result)
      if ([int]$result.total -eq 0) {
        Add-Log "0 records."
        [System.Windows.Forms.MessageBox]::Show("0 records.") | Out-Null
        return
      }
      Add-Log ("{0}: {1}" -f (T "Done"), [string]$result.file)
      [System.Windows.Forms.MessageBox]::Show(("OK`r`n{0}`r`nRecords: {1}" -f [string]$result.file, [int]$result.total)) | Out-Null
    } $ctx
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

  if (@($script:Settings.cachedTables).Count -gt 0) {
    $cmbTable.BeginUpdate()
    $cmbTable.Items.Clear()
    foreach ($t in @($script:Settings.cachedTables)) {
      [void]$cmbTable.Items.Add(("{0} - {1}" -f $t.name, $t.label))
    }
    $cmbTable.EndUpdate()
    Refresh-BaseTableItems
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
        $gridJoins.Rows[$rowIndex].Cells[0].Value = [string]$j.joinTable
        Populate-JoinColumnsForRow $rowIndex
        if ($j.PSObject.Properties.Name -contains "joinSource") { $gridJoins.Rows[$rowIndex].Cells[1].Value = [string]$j.joinSource }
        else { $gridJoins.Rows[$rowIndex].Cells[1].Value = "__base__" }
        Populate-JoinColumnsForRow $rowIndex
        $gridJoins.Rows[$rowIndex].Cells[2].Value = [string]$j.baseColumn
        $gridJoins.Rows[$rowIndex].Cells[3].Value = [string]$j.targetColumn
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

  $txtDir.add_TextChanged({
    $script:Settings.exportDirectory = $txtDir.Text
    Request-SaveSettings
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
    if ($tabs.SelectedTab -eq $tabViewEditor) {
      Ensure-TablesLoaded
    }
  })

  $cmbOutputFormat.add_SelectedIndexChanged({
    $script:Settings.outputFormat = [string]$cmbOutputFormat.SelectedItem
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

  # First-run export dir
  try { [void](Ensure-ExportDir $txtDir.Text) } catch { }

  $form.add_FormClosing({
    Complete-GridCurrentEdit $gridJoins "Join"
    Save-JoinDefinitionsToSettings
    $script:Settings.viewEditorSelectedColumnsJson = (@(Get-SelectedViewFieldTokens) | ConvertTo-Json -Compress)
    Request-SaveSettings -Immediate
  })

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
