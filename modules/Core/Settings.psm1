Set-StrictMode -Version Latest

function Protect-Secret {
  param([string]$Plain)
  if ([string]::IsNullOrWhiteSpace($Plain)) { return "" }
  $sec = ConvertTo-SecureString $Plain -AsPlainText -Force
  return (ConvertFrom-SecureString $sec)
}

function Unprotect-Secret {
  param([string]$Encrypted)
  if ([string]::IsNullOrWhiteSpace($Encrypted)) { return "" }
  try {
    $sec = ConvertTo-SecureString $Encrypted
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
  } catch {
    return ""
  }
}

function New-DefaultSettings {
  return [pscustomobject]@{
    settingsVersion = 2
    uiLanguage = "ja"
    instanceName = ""
    authType = "userpass"
    userId = ""
    passwordEnc = ""
    apiKeyEnc = ""
    exportDirectory = ""
    filterMode = "all"
    startDateTime = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss")
    endDateTime   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    cachedTables = @()
    cachedTablesFetchedAt = ""
    selectedTableName = ""
    exportFields = ""
    pageSize = 1000
    outputFormat = "csv"
    viewEditorViewName = ""
    viewEditorViewLabel = ""
    viewEditorBaseTable = ""
    viewEditorBasePrefix = "t0"
    viewEditorJoinsJson = "[]"
    viewEditorSelectedColumnsJson = "[]"
    deleteTargetTable = ""
    deleteMaxRetries = 99
  }
}

function Get-SettingsVersion {
  param($Settings)

  if (-not $Settings) { return 1 }
  if ($Settings.PSObject.Properties.Name -contains 'settingsVersion') {
    $ver = 0
    if ([int]::TryParse([string]$Settings.settingsVersion, [ref]$ver) -and $ver -ge 1) {
      return $ver
    }
  }
  return 1
}

function Migrate-SettingsV1ToV2 {
  param([Parameter(Mandatory=$true)]$Settings)

  if (-not ($Settings.PSObject.Properties.Name -contains 'outputFormat')) {
    $Settings | Add-Member -NotePropertyName outputFormat -NotePropertyValue 'csv'
  }
  if (-not ($Settings.PSObject.Properties.Name -contains 'viewEditorViewName')) {
    $Settings | Add-Member -NotePropertyName viewEditorViewName -NotePropertyValue ''
  }
  if (-not ($Settings.PSObject.Properties.Name -contains 'viewEditorViewLabel')) {
    $Settings | Add-Member -NotePropertyName viewEditorViewLabel -NotePropertyValue ''
  }
  if (-not ($Settings.PSObject.Properties.Name -contains 'viewEditorBaseTable')) {
    $Settings | Add-Member -NotePropertyName viewEditorBaseTable -NotePropertyValue ''
  }
  if (-not ($Settings.PSObject.Properties.Name -contains 'viewEditorBasePrefix')) {
    $Settings | Add-Member -NotePropertyName viewEditorBasePrefix -NotePropertyValue 't0'
  }
  if (-not ($Settings.PSObject.Properties.Name -contains 'viewEditorJoinsJson')) {
    $Settings | Add-Member -NotePropertyName viewEditorJoinsJson -NotePropertyValue '[]'
  }
  if (-not ($Settings.PSObject.Properties.Name -contains 'viewEditorSelectedColumnsJson')) {
    $Settings | Add-Member -NotePropertyName viewEditorSelectedColumnsJson -NotePropertyValue '[]'
  }
  if (-not ($Settings.PSObject.Properties.Name -contains 'deleteTargetTable')) {
    $Settings | Add-Member -NotePropertyName deleteTargetTable -NotePropertyValue ''
  }
  if (-not ($Settings.PSObject.Properties.Name -contains 'deleteMaxRetries')) {
    $Settings | Add-Member -NotePropertyName deleteMaxRetries -NotePropertyValue 99
  }

  if ($Settings.PSObject.Properties.Name -contains 'settingsVersion') {
    $Settings.settingsVersion = 2
  } else {
    $Settings | Add-Member -NotePropertyName settingsVersion -NotePropertyValue 2
  }

  return $Settings
}

function Migrate-Settings {
  param([Parameter(Mandatory=$true)]$Settings)

  $originalVersion = Get-SettingsVersion -Settings $Settings
  $currentVersion = $originalVersion
  $migrated = $Settings

  if ($currentVersion -lt 2) {
    $migrated = Migrate-SettingsV1ToV2 -Settings $migrated
    $currentVersion = 2
  }

  return [pscustomobject]@{
    Settings = $migrated
    Migrated = ($originalVersion -ne $currentVersion)
  }
}

function Load-Settings {
  param(
    [Parameter(Mandatory=$true)][string]$SettingsPath
  )

  $defaults = New-DefaultSettings
  $settings = $defaults
  $isMigrated = $false

  if (Test-Path $SettingsPath) {
    try {
      $json = Get-Content $SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
      if ($json) {
        $migration = Migrate-Settings -Settings $json
        $settings = $migration.Settings
        $isMigrated = [bool]$migration.Migrated

        foreach ($p in $defaults.PSObject.Properties.Name) {
          if ($settings -and ($settings.PSObject.Properties.Name -contains $p) -and $null -ne $settings.$p) {
            $defaults.$p = $settings.$p
          }
        }
      }
    } catch {
      # keep defaults
    }
  }

  if ($isMigrated) {
    Save-Settings -Settings $defaults -SettingsPath $SettingsPath
  }

  return $defaults
}

function Save-Settings {
  param(
    [Parameter(Mandatory=$true)]$Settings,
    [Parameter(Mandatory=$true)][string]$SettingsPath
  )

  try {
    $out = ($Settings | ConvertTo-Json -Depth 8)
    [System.IO.File]::WriteAllText($SettingsPath, $out, (New-Object System.Text.UTF8Encoding($false)))
  } catch {
    # ignore write failure
  }
}

Export-ModuleMember -Function Protect-Secret, Unprotect-Secret, New-DefaultSettings, Get-SettingsVersion, Migrate-Settings, Load-Settings, Save-Settings
