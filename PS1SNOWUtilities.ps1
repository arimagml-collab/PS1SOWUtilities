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
      AddCondition="条件追加"
      RemoveCondition="条件削除"
      WhereClausePreview="Where句プレビュー"
      CreateView="View作成"
      ConditionColumn="カラム"
      ConditionOperator="演算子"
      ConditionValue="値"
      WarnViewName="View内部名を入力してください。"
      WarnViewLabel="Viewラベルを入力してください。"
      WarnBaseTable="ベーステーブルを選択してください。"
      WarnConditionColumn="条件のカラムが未選択です。"
      WarnConditionValue="値が必要な条件があります。"
      FetchingColumns="カラム一覧を取得中..."
      ColumnsFetched="カラム一覧を取得"
      ViewCreated="Viewを作成しました"
      ViewCreateFailed="View作成に失敗しました"
      ViewWhereFallback="この環境ではWhere句保存フィールドを特定できませんでした。Viewは作成されましたが、Where句は手動設定してください。"
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
      AddCondition="Add Condition"
      RemoveCondition="Remove Condition"
      WhereClausePreview="Where Clause Preview"
      CreateView="Create View"
      ConditionColumn="Column"
      ConditionOperator="Operator"
      ConditionValue="Value"
      WarnViewName="View name is required."
      WarnViewLabel="View label is required."
      WarnBaseTable="Base table must be selected."
      WarnConditionColumn="One or more conditions have no column selected."
      WarnConditionValue="One or more conditions require a value."
      FetchingColumns="Fetching columns..."
      ColumnsFetched="Fetched columns"
      ViewCreated="View created"
      ViewCreateFailed="Failed to create view"
      ViewWhereFallback="Could not detect a writable where-clause field in this instance. View was created, but set the where clause manually."
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
      viewEditorWhereClause = ""
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

  function Invoke-SnowPost([string]$path, [hashtable]$body) {
    $base = Get-BaseUrl
    if ([string]::IsNullOrWhiteSpace($base)) { throw (T "WarnInstance") }

    $uri = $base + $path
    $headers = New-SnowHeaders
    $jsonBody = ($body | ConvertTo-Json -Depth 8)

    if ($script:Settings.authType -eq "userpass") {
      $user = ([string]$script:Settings.userId).Trim()
      $pass = Unprotect-Secret ([string]$script:Settings.passwordEnc)
      if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) { throw (T "WarnAuth") }
      $sec = ConvertTo-SecureString $pass -AsPlainText -Force
      $cred = New-Object System.Management.Automation.PSCredential($user, $sec)
      return Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Credential $cred -TimeoutSec 120 -Body $jsonBody
    } else {
      return Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -TimeoutSec 120 -Body $jsonBody
    }
  }

  function Invoke-SnowPatch([string]$path, [hashtable]$body) {
    $base = Get-BaseUrl
    if ([string]::IsNullOrWhiteSpace($base)) { throw (T "WarnInstance") }

    $uri = $base + $path
    $headers = New-SnowHeaders
    $jsonBody = ($body | ConvertTo-Json -Depth 8)

    if ($script:Settings.authType -eq "userpass") {
      $user = ([string]$script:Settings.userId).Trim()
      $pass = Unprotect-Secret ([string]$script:Settings.passwordEnc)
      if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) { throw (T "WarnAuth") }
      $sec = ConvertTo-SecureString $pass -AsPlainText -Force
      $cred = New-Object System.Management.Automation.PSCredential($user, $sec)
      return Invoke-RestMethod -Method Patch -Uri $uri -Headers $headers -Credential $cred -TimeoutSec 120 -Body $jsonBody
    } else {
      return Invoke-RestMethod -Method Patch -Uri $uri -Headers $headers -TimeoutSec 120 -Body $jsonBody
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
  $tabViewEditor = New-Object System.Windows.Forms.TabPage
  $tabSettings = New-Object System.Windows.Forms.TabPage

  [void]$tabs.TabPages.Add($tabExport)
  [void]$tabs.TabPages.Add($tabViewEditor)
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

  # --- DataBase View Editor tab layout
  $panelViewEditor = New-Object System.Windows.Forms.Panel
  $panelViewEditor.Dock = "Fill"
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

  $clbViewColumns = New-Object System.Windows.Forms.CheckedListBox
  $clbViewColumns.Location = New-Object System.Drawing.Point(190, 100)
  $clbViewColumns.Size = New-Object System.Drawing.Size(730, 120)

  $btnAddCondition = New-Object System.Windows.Forms.Button
  $btnAddCondition.Location = New-Object System.Drawing.Point(190, 230)
  $btnAddCondition.Size = New-Object System.Drawing.Size(170, 32)

  $btnRemoveCondition = New-Object System.Windows.Forms.Button
  $btnRemoveCondition.Location = New-Object System.Drawing.Point(370, 230)
  $btnRemoveCondition.Size = New-Object System.Drawing.Size(170, 32)

  $gridConditions = New-Object System.Windows.Forms.DataGridView
  $gridConditions.Location = New-Object System.Drawing.Point(190, 270)
  $gridConditions.Size = New-Object System.Drawing.Size(730, 200)
  $gridConditions.AllowUserToAddRows = $false
  $gridConditions.AllowUserToDeleteRows = $false
  $gridConditions.RowHeadersVisible = $false
  $gridConditions.SelectionMode = "FullRowSelect"
  $gridConditions.MultiSelect = $false
  $gridConditions.AutoSizeColumnsMode = "Fill"

  $colCondColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
  $colCondColumn.Name = "ConditionColumn"
  $colCondColumn.FlatStyle = "Popup"
  $colCondColumn.DisplayStyle = "DropDownButton"
  $colCondColumn.FillWeight = 45

  $colCondOperator = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
  $colCondOperator.Name = "ConditionOperator"
  $colCondOperator.FlatStyle = "Popup"
  $colCondOperator.DisplayStyle = "DropDownButton"
  $colCondOperator.FillWeight = 20
  [void]$colCondOperator.Items.Add("=")
  [void]$colCondOperator.Items.Add("!=")
  [void]$colCondOperator.Items.Add("LIKE")
  [void]$colCondOperator.Items.Add("STARTSWITH")
  [void]$colCondOperator.Items.Add("ENDSWITH")
  [void]$colCondOperator.Items.Add("IN")
  [void]$colCondOperator.Items.Add("ISEMPTY")
  [void]$colCondOperator.Items.Add("ISNOTEMPTY")

  $colCondValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colCondValue.Name = "ConditionValue"
  $colCondValue.FillWeight = 35

  [void]$gridConditions.Columns.Add($colCondColumn)
  [void]$gridConditions.Columns.Add($colCondOperator)
  [void]$gridConditions.Columns.Add($colCondValue)

  $lblWherePreview = New-Object System.Windows.Forms.Label
  $lblWherePreview.Location = New-Object System.Drawing.Point(20, 485)
  $lblWherePreview.AutoSize = $true

  $txtWherePreview = New-Object System.Windows.Forms.TextBox
  $txtWherePreview.Location = New-Object System.Drawing.Point(190, 482)
  $txtWherePreview.Size = New-Object System.Drawing.Size(730, 60)
  $txtWherePreview.Multiline = $true
  $txtWherePreview.ReadOnly = $true
  $txtWherePreview.ScrollBars = "Vertical"

  $btnCreateView = New-Object System.Windows.Forms.Button
  $btnCreateView.Location = New-Object System.Drawing.Point(740, 550)
  $btnCreateView.Size = New-Object System.Drawing.Size(180, 42)

  $panelViewEditor.Controls.AddRange(@(
    $lblViewName, $txtViewName,
    $lblViewLabel, $txtViewLabel,
    $lblBaseTable, $cmbBaseTable, $btnReloadColumns,
    $lblViewColumns, $clbViewColumns,
    $btnAddCondition, $btnRemoveCondition,
    $gridConditions,
    $lblWherePreview, $txtWherePreview,
    $btnCreateView
  ))

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
    $btnAddCondition.Text = T "AddCondition"
    $btnRemoveCondition.Text = T "RemoveCondition"
    $lblWherePreview.Text = T "WhereClausePreview"
    $btnCreateView.Text = T "CreateView"
    $colCondColumn.HeaderText = T "ConditionColumn"
    $colCondOperator.HeaderText = T "ConditionOperator"
    $colCondValue.HeaderText = T "ConditionValue"

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
      $cmbTable.Items.Clear()
      foreach ($t in $list) {
        [void]$cmbTable.Items.Add(("{0} - {1}" -f $t.name, $t.label))
      }
      $cmbTable.EndUpdate()
      Refresh-BaseTableItems

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
        } else {
          $cmbTable.Text = $targetName
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
      foreach ($t in $script:Settings.cachedTables) {
        [void]$cmbBaseTable.Items.Add(("{0} - {1}" -f $t.name, $t.label))
      }
    }
    $cmbBaseTable.EndUpdate()
  }

  function Build-ViewWhereClause {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($row in $gridConditions.Rows) {
      if ($row.IsNewRow) { continue }
      $columnCell = $row.Cells[0].Value
      $opCell = $row.Cells[1].Value
      $valueCell = $row.Cells[2].Value

      $column = if ($null -eq $columnCell) { "" } else { [string]$columnCell }
      $op = if ($null -eq $opCell -or [string]::IsNullOrWhiteSpace([string]$opCell)) { "=" } else { [string]$opCell }
      $value = if ($null -eq $valueCell) { "" } else { [string]$valueCell }

      if ([string]::IsNullOrWhiteSpace($column)) { continue }
      if ((@("ISEMPTY","ISNOTEMPTY") -contains $op)) {
        [void]$parts.Add(("{0}{1}" -f $column, $op))
      } else {
        [void]$parts.Add(("{0}{1}{2}" -f $column, $op, $value))
      }
    }
    return ($parts -join "^")
  }

  function Update-WherePreview {
    $txtWherePreview.Text = Build-ViewWhereClause
    $script:Settings.viewEditorWhereClause = $txtWherePreview.Text
    Save-Settings
  }

  function Fetch-ColumnsForBaseTable {
    $table = Get-SelectedBaseTableName
    if ([string]::IsNullOrWhiteSpace($table)) {
      [System.Windows.Forms.MessageBox]::Show((T "WarnBaseTable")) | Out-Null
      return
    }

    Add-Log ("{0} [{1}]" -f (T "FetchingColumns"), $table)
    try {
      $q = "name=$table^elementISNOTEMPTY"
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
      $list = $list | Sort-Object name -Unique

      $clbViewColumns.BeginUpdate()
      $clbViewColumns.Items.Clear()
      foreach ($c in $list) {
        [void]$clbViewColumns.Items.Add(("{0} - {1}" -f $c.name, $c.label), $true)
      }
      $clbViewColumns.EndUpdate()

      $colCondColumn.Items.Clear()
      foreach ($c in $list) {
        [void]$colCondColumn.Items.Add(("{0} - {1}" -f $c.name, $c.label))
      }

      Add-Log ("{0}: {1}" -f (T "ColumnsFetched"), $list.Count)
    } catch {
      Add-Log ("{0}: {1}" -f (T "Failed"), $_.Exception.Message)
    }
  }

  function Create-DatabaseView {
    $viewName = ([string]$txtViewName.Text).Trim()
    $viewLabel = ([string]$txtViewLabel.Text).Trim()
    $baseTable = Get-SelectedBaseTableName

    if ([string]::IsNullOrWhiteSpace($viewName)) { [System.Windows.Forms.MessageBox]::Show((T "WarnViewName")) | Out-Null; return }
    if ([string]::IsNullOrWhiteSpace($viewLabel)) { [System.Windows.Forms.MessageBox]::Show((T "WarnViewLabel")) | Out-Null; return }
    if ([string]::IsNullOrWhiteSpace($baseTable)) { [System.Windows.Forms.MessageBox]::Show((T "WarnBaseTable")) | Out-Null; return }

    foreach ($row in $gridConditions.Rows) {
      if ($row.IsNewRow) { continue }
      $cVal = $row.Cells[0].Value
      $oVal = $row.Cells[1].Value
      $vVal = $row.Cells[2].Value
      $columnText = if ($null -eq $cVal) { "" } else { [string]$cVal }
      $opText = if ($null -eq $oVal -or [string]::IsNullOrWhiteSpace([string]$oVal)) { "=" } else { [string]$oVal }
      $valueText = if ($null -eq $vVal) { "" } else { [string]$vVal }
      if ([string]::IsNullOrWhiteSpace($columnText)) { [System.Windows.Forms.MessageBox]::Show((T "WarnConditionColumn")) | Out-Null; return }
      if ((@("ISEMPTY","ISNOTEMPTY") -notcontains $opText) -and [string]::IsNullOrWhiteSpace($valueText)) { [System.Windows.Forms.MessageBox]::Show((T "WarnConditionValue")) | Out-Null; return }
    }

    $whereClause = Build-ViewWhereClause
    $selectedColumns = New-Object System.Collections.Generic.List[string]
    foreach ($item in $clbViewColumns.CheckedItems) {
      $itemText = [string]$item
      $idx = $itemText.IndexOf(" - ")
      if ($idx -gt 0) { [void]$selectedColumns.Add($itemText.Substring(0, $idx).Trim()) }
      elseif (-not [string]::IsNullOrWhiteSpace($itemText)) { [void]$selectedColumns.Add($itemText.Trim()) }
    }

    Add-Log ("Creating DB view: {0}, base={1}" -f $viewName, $baseTable)
    try {
      $body = @{ name = $viewName; label = $viewLabel; table = $baseTable }
      if ($selectedColumns.Count -gt 0) { $body["view_fields"] = ($selectedColumns -join ",") }
      $createRes = Invoke-SnowPost "/api/now/table/sys_db_view" $body
      $created = if ($createRes -and ($createRes.PSObject.Properties.Name -contains "result")) { $createRes.result } else { $null }
      $sysId = if ($created) { [string]$created.sys_id } else { "" }

      $whereSaved = $false
      if (-not [string]::IsNullOrWhiteSpace($sysId) -and -not [string]::IsNullOrWhiteSpace($whereClause)) {
        foreach ($whereField in @("where_clause", "where", "condition")) {
          try {
            [void](Invoke-SnowPatch ("/api/now/table/sys_db_view/{0}" -f $sysId) @{ $whereField = $whereClause })
            $whereSaved = $true
            break
          } catch {
          }
        }
      } else {
        $whereSaved = $true
      }

      if (-not $whereSaved) {
        Add-Log (T "ViewWhereFallback")
        [System.Windows.Forms.MessageBox]::Show((T "ViewWhereFallback")) | Out-Null
      }

      Add-Log ("{0}: {1}" -f (T "ViewCreated"), $viewName)
      [System.Windows.Forms.MessageBox]::Show(("{0}`r`n{1}" -f (T "ViewCreated"), $viewName)) | Out-Null
    } catch {
      Add-Log ("{0}: {1}" -f (T "ViewCreateFailed"), $_.Exception.Message)
      [System.Windows.Forms.MessageBox]::Show(("{0}`r`n{1}" -f (T "ViewCreateFailed"), $_.Exception.Message)) | Out-Null
    }
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
    $cmbTable.Items.Clear()
    foreach ($t in $script:Settings.cachedTables) {
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

  $txtWherePreview.Text = [string]$script:Settings.viewEditorWhereClause

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
    $script:Settings.selectedTableName = Get-SelectedTableName
    Save-Settings
  })

  $cmbTable.add_TextChanged({
    $script:Settings.selectedTableName = Get-SelectedTableName
    Save-Settings
  })

  $txtDir.add_TextChanged({
    $script:Settings.exportDirectory = $txtDir.Text
    Save-Settings
  })

  $txtViewName.add_TextChanged({
    $script:Settings.viewEditorViewName = $txtViewName.Text
    Save-Settings
  })

  $txtViewLabel.add_TextChanged({
    $script:Settings.viewEditorViewLabel = $txtViewLabel.Text
    Save-Settings
  })

  $cmbBaseTable.add_SelectedIndexChanged({
    $script:Settings.viewEditorBaseTable = Get-SelectedBaseTableName
    Save-Settings
  })

  $cmbBaseTable.add_TextChanged({
    $script:Settings.viewEditorBaseTable = Get-SelectedBaseTableName
    Save-Settings
  })

  $btnReloadColumns.add_Click({ Fetch-ColumnsForBaseTable })

  $btnAddCondition.add_Click({
    $rowIndex = $gridConditions.Rows.Add()
    if ($rowIndex -ge 0) {
      $gridConditions.Rows[$rowIndex].Cells[1].Value = "="
      Update-WherePreview
    }
  })

  $btnRemoveCondition.add_Click({
    if ($gridConditions.SelectedRows.Count -gt 0) {
      $gridConditions.Rows.Remove($gridConditions.SelectedRows[0])
      Update-WherePreview
    }
  })

  $gridConditions.add_CellValueChanged({ Update-WherePreview })
  $gridConditions.add_RowsRemoved({ Update-WherePreview })
  $gridConditions.add_CurrentCellDirtyStateChanged({
    if ($gridConditions.IsCurrentCellDirty) {
      [void]$gridConditions.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
  })

  $btnCreateView.add_Click({ Create-DatabaseView })

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
