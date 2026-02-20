Set-StrictMode -Version Latest

function ConvertTo-HashtableRecursive {
  param([Parameter(Mandatory=$true)]$InputObject)

  if ($InputObject -is [hashtable]) {
    $copied = @{}
    foreach ($key in $InputObject.Keys) {
      $copied[$key] = ConvertTo-HashtableRecursive -InputObject $InputObject[$key]
    }
    return $copied
  }

  if ($InputObject -is [System.Collections.IDictionary]) {
    $dict = @{}
    foreach ($entry in $InputObject.GetEnumerator()) {
      $dict[[string]$entry.Key] = ConvertTo-HashtableRecursive -InputObject $entry.Value
    }
    return $dict
  }

  if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
    $arr = @()
    foreach ($item in $InputObject) {
      $arr += ,(ConvertTo-HashtableRecursive -InputObject $item)
    }
    return $arr
  }

  if ($InputObject -is [pscustomobject]) {
    $obj = @{}
    foreach ($property in $InputObject.PSObject.Properties) {
      $obj[$property.Name] = ConvertTo-HashtableRecursive -InputObject $property.Value
    }
    return $obj
  }

  return $InputObject
}

function Load-I18nResources {
  param(
    [Parameter(Mandatory=$true)][string]$LocalesDirectory,
    [string]$DefaultLanguage = 'ja'
  )

  $resources = @{}
  if (-not (Test-Path -LiteralPath $LocalesDirectory)) {
    throw "Locales directory not found: $LocalesDirectory"
  }

  $localeFiles = Get-ChildItem -LiteralPath $LocalesDirectory -Filter '*.json' -File
  foreach ($localeFile in $localeFiles) {
    try {
      $parsed = Get-Content -LiteralPath $localeFile.FullName -Encoding UTF8 -Raw | ConvertFrom-Json
      $resources[$localeFile.BaseName] = ConvertTo-HashtableRecursive -InputObject $parsed
    } catch {
      throw "Failed to load locale file: $($localeFile.FullName). $($_.Exception.Message)"
    }
  }

  if (-not $resources.ContainsKey($DefaultLanguage)) {
    throw "Default locale is missing: $DefaultLanguage"
  }

  return $resources
}

function Resolve-I18nText {
  param(
    [Parameter(Mandatory=$true)][hashtable]$I18nResources,
    [Parameter(Mandatory=$true)][string]$Language,
    [Parameter(Mandatory=$true)][string]$Key,
    [string]$DefaultLanguage = 'ja'
  )

  if ($I18nResources.ContainsKey($Language) -and $I18nResources[$Language].ContainsKey($Key)) {
    return [string]$I18nResources[$Language][$Key]
  }

  if ($I18nResources.ContainsKey($DefaultLanguage) -and $I18nResources[$DefaultLanguage].ContainsKey($Key)) {
    return [string]$I18nResources[$DefaultLanguage][$Key]
  }

  return $Key
}

Export-ModuleMember -Function Load-I18nResources, Resolve-I18nText
