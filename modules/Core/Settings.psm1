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

function Load-Settings {
  param(
    [Parameter(Mandatory=$true)][string]$SettingsPath
  )

  $defaults = New-DefaultSettings
  if (Test-Path $SettingsPath) {
    try {
      $json = Get-Content $SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
      foreach ($p in $defaults.PSObject.Properties.Name) {
        if ($json -and ($json.PSObject.Properties.Name -contains $p) -and $null -ne $json.$p) {
          $defaults.$p = $json.$p
        }
      }
    } catch {
      # keep defaults
    }
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

Export-ModuleMember -Function Protect-Secret, Unprotect-Secret, New-DefaultSettings, Load-Settings, Save-Settings
