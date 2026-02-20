Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

Import-Module (Join-Path $repoRoot 'modules/Core/Settings.psm1') -Force -Prefix Core
Import-Module (Join-Path $repoRoot 'modules/Core/I18n.psm1') -Force -Prefix Core

$tmpDir = Join-Path $env:TEMP ('ps1snow_phase3_' + [Guid]::NewGuid().ToString('N'))
New-Item -Path $tmpDir -ItemType Directory | Out-Null

try {
  $legacyPath = Join-Path $tmpDir 'settings.json'
  @'
{
  "uiLanguage": "en",
  "instanceName": "demo",
  "authType": "userpass",
  "userId": "admin",
  "passwordEnc": "",
  "apiKeyEnc": "",
  "exportDirectory": "",
  "filterMode": "all",
  "startDateTime": "",
  "endDateTime": "",
  "cachedTables": [],
  "cachedTablesFetchedAt": "",
  "selectedTableName": "incident",
  "exportFields": "",
  "pageSize": 1000
}
'@ | Set-Content -Path $legacyPath -Encoding UTF8

  $loaded = CoreLoad-Settings -SettingsPath $legacyPath
  if ($loaded.settingsVersion -ne 2) { throw 'settingsVersion migration failed.' }
  if (-not ($loaded.PSObject.Properties.Name -contains 'outputFormat')) { throw 'outputFormat was not added by migration.' }
  if (-not ($loaded.PSObject.Properties.Name -contains 'outputEncoding')) { throw 'outputEncoding was not added by migration.' }
  if (-not ($loaded.PSObject.Properties.Name -contains 'outputBom')) { throw 'outputBom was not added by migration.' }

  $persisted = Get-Content -Path $legacyPath -Raw -Encoding UTF8 | ConvertFrom-Json
  if ($persisted.settingsVersion -ne 2) { throw 'Migrated settings were not persisted.' }

  $i18n = CoreLoad-I18nResources -LocalesDirectory (Join-Path $repoRoot 'locales') -DefaultLanguage 'ja'
  $ja = CoreResolve-I18nText -I18nResources $i18n -Language 'ja' -Key 'AppTitle' -DefaultLanguage 'ja'
  $fallback = CoreResolve-I18nText -I18nResources $i18n -Language 'en' -Key 'NoSuchKey' -DefaultLanguage 'ja'
  if ($ja -ne 'PS1 SNOW Utilities') { throw 'ja locale load failed.' }
  if ($fallback -ne 'NoSuchKey') { throw 'fallback behavior failed.' }

  Write-Host 'phase3 checks passed'
}
finally {
  if (Test-Path $tmpDir) {
    Remove-Item -Path $tmpDir -Recurse -Force
  }
}
