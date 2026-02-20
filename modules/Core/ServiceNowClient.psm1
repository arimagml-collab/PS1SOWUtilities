Set-StrictMode -Version Latest

function UrlEncode {
  param([string]$Value)
  return [System.Uri]::EscapeDataString($Value)
}

function Get-BaseUrl {
  param([Parameter(Mandatory=$true)]$Settings)

  $instVal = $Settings.instanceName
  if ($null -eq $instVal) { $instVal = "" }
  $inst = ([string]$instVal).Trim()
  if ([string]::IsNullOrWhiteSpace($inst)) { return "" }

  if ($inst -match '^https?://') { return $inst.TrimEnd('/') }
  if ($inst -match '\.service-now\.com$') { return ("https://{0}" -f $inst).TrimEnd('/') }
  return ("https://{0}.service-now.com" -f $inst).TrimEnd('/')
}

function New-SnowHeaders {
  param(
    [Parameter(Mandatory=$true)]$Settings,
    [Parameter(Mandatory=$true)][scriptblock]$UnprotectSecret
  )

  $headers = @{
    "Accept" = "application/json"
    "Content-Type" = "application/json; charset=utf-8"
  }

  if ($Settings.authType -eq "apikey") {
    $key = & $UnprotectSecret ([string]$Settings.apiKeyEnc)
    if (-not [string]::IsNullOrWhiteSpace($key)) {
      $headers["Authorization"] = "Bearer $key"
    }
  }

  return $headers
}

function Invoke-SnowRequest {
  param(
    [Parameter(Mandatory=$true)][ValidateSet('Get','Post','Patch','Delete')][string]$Method,
    [Parameter(Mandatory=$true)][string]$Path,
    [AllowNull()]$Body,
    [Parameter(Mandatory=$true)]$Settings,
    [Parameter(Mandatory=$true)][scriptblock]$UnprotectSecret,
    [Parameter(Mandatory=$true)][scriptblock]$GetText,
    [int]$TimeoutSec = 120
  )

  $base = Get-BaseUrl -Settings $Settings
  if ([string]::IsNullOrWhiteSpace($base)) { throw (& $GetText "WarnInstance") }

  $uri = $base + $Path
  $headers = New-SnowHeaders -Settings $Settings -UnprotectSecret $UnprotectSecret

  $requestParams = @{
    Method = $Method
    Uri = $uri
    Headers = $headers
    TimeoutSec = $TimeoutSec
  }

  if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
    $jsonBody = ($Body | ConvertTo-Json -Depth 8)
    $requestParams.Body = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)
  }

  if ($Settings.authType -eq "userpass") {
    $user = ([string]$Settings.userId).Trim()
    $pass = & $UnprotectSecret ([string]$Settings.passwordEnc)
    if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) {
      throw (& $GetText "WarnAuth")
    }

    $sec = ConvertTo-SecureString $pass -AsPlainText -Force
    $requestParams.Credential = New-Object System.Management.Automation.PSCredential($user, $sec)
  }

  return Invoke-RestMethod @requestParams
}

function Invoke-SnowBatchDelete {
  param(
    [Parameter(Mandatory=$true)][string]$Table,
    [string[]]$SysIds,
    [Parameter(Mandatory=$true)][scriptblock]$InvokePost
  )

  if ([string]::IsNullOrWhiteSpace($Table)) {
    return @{ deletedCount = 0; failedIds = @() }
  }
  $ids = @($SysIds | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
  if ($ids.Count -eq 0) { return @{ deletedCount = 0; failedIds = @() } }

  $requests = New-Object System.Collections.Generic.List[hashtable]
  $index = 0
  foreach ($id in $ids) {
    $index++
    $requests.Add(@{
      id = [string]$index
      method = "DELETE"
      url = ("/api/now/table/{0}/{1}" -f $Table, [string]$id)
      headers = @{ "X-no-response-body" = "true" }
    }) | Out-Null
  }

  $batchBody = @{
    batch_request_id = [Guid]::NewGuid().ToString("N")
    rest_requests = @($requests)
  }

  $res = & $InvokePost "/api/now/v1/batch" $batchBody
  $responses = @()
  if ($res -and ($res.PSObject.Properties.Name -contains "result")) {
    $responses = @($res.result)
  }

  $ok = 0
  $failedIds = New-Object System.Collections.Generic.List[string]
  foreach ($item in $responses) {
    $status = 0
    try { $status = [int]$item.status_code } catch { $status = 0 }
    $responseId = ""
    try { $responseId = [string]$item.id } catch { $responseId = "" }
    $reqIndex = 0
    if (-not [int]::TryParse($responseId, [ref]$reqIndex)) { $reqIndex = 0 }
    $targetId = ""
    if ($reqIndex -ge 1 -and $reqIndex -le $ids.Count) {
      $targetId = [string]$ids[$reqIndex - 1]
    }
    if ($status -ge 200 -and $status -lt 300) {
      $ok++
    } elseif (-not [string]::IsNullOrWhiteSpace($targetId)) {
      $failedIds.Add($targetId) | Out-Null
    }
  }

  if ($responses.Count -eq 0) {
    return @{ deletedCount = $ids.Count; failedIds = @() }
  }

  return @{ deletedCount = $ok; failedIds = @($failedIds) }
}

Export-ModuleMember -Function UrlEncode, Get-BaseUrl, New-SnowHeaders, Invoke-SnowRequest, Invoke-SnowBatchDelete
