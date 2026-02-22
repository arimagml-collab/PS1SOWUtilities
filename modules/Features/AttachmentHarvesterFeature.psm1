Set-StrictMode -Version Latest

function Validate-AttachmentHarvesterInput {
  param(
    [Parameter(Mandatory=$true)][string]$BaseUrl,
    [Parameter(Mandatory=$true)][string]$Table,
    [Parameter(Mandatory=$true)][string]$DownloadDirectory,
    [Parameter(Mandatory=$true)]$Settings,
    [Parameter(Mandatory=$true)][scriptblock]$UnprotectSecret,
    [Parameter(Mandatory=$true)][scriptblock]$GetText
  )

  if ([string]::IsNullOrWhiteSpace($BaseUrl)) {
    return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnInstance")) }
  }
  if ([string]::IsNullOrWhiteSpace($Table)) {
    return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnTable")) }
  }
  if ([string]::IsNullOrWhiteSpace($DownloadDirectory)) {
    return [pscustomobject]@{ IsValid = $false; Errors = @("Download directory is required.") }
  }

  $authType = ([string]$Settings.authType).Trim().ToLowerInvariant()
  if ($authType -eq "userpass") {
    $user = [string]$Settings.userId
    $pass = & $UnprotectSecret ([string]$Settings.passwordEnc)
    if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) {
      return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnAuth")) }
    }
  } elseif ($authType -eq "apikey") {
    $key = & $UnprotectSecret ([string]$Settings.apiKeyEnc)
    if ([string]::IsNullOrWhiteSpace($key)) {
      return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnAuth")) }
    }
  } else {
    return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnAuth")) }
  }

  return [pscustomobject]@{ IsValid = $true; Errors = @() }
}

function Invoke-AttachmentHarvesterUseCase {
  param(
    [Parameter(Mandatory=$true)]$Context,
    [Parameter(Mandatory=$true)][scriptblock]$InvokeSnowGet,
    [Parameter(Mandatory=$true)][scriptblock]$UrlEncode,
    [Parameter(Mandatory=$true)][scriptblock]$DownloadAttachmentBytes,
    [Parameter(Mandatory=$true)][scriptblock]$WriteLog
  )

  function Convert-ToSafeName {
    param([AllowNull()][string]$Text)

    $value = if ($null -eq $Text) { '' } else { [string]$Text }
    $value = $value -replace '[\\/:\*\?"<>\|]', '_'
    $value = $value -replace '[\r\n\t]', '_'
    $value = $value -replace '\s+', '_'
    $value = $value -replace '_+', '_'
    $value = $value.Trim('_', ' ')
    if ([string]::IsNullOrWhiteSpace($value)) { return 'unnamed' }
    return $value
  }

  function Get-Sha256Hex {
    param([byte[]]$Bytes)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
      $hash = $sha.ComputeHash($Bytes)
      return ([System.BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
    } finally {
      $sha.Dispose()
    }
  }

  function Resolve-UniquePath {
    param(
      [Parameter(Mandatory=$true)][string]$Directory,
      [Parameter(Mandatory=$true)][string]$FileName,
      [Parameter(Mandatory=$true)][byte[]]$Content
    )

    $full = Join-Path $Directory $FileName
    $newHash = Get-Sha256Hex -Bytes $Content

    if (-not (Test-Path $full)) {
      return [pscustomobject]@{ Action = 'Save'; Path = $full; Hash = $newHash }
    }

    $existingHash = Get-Sha256Hex -Bytes ([System.IO.File]::ReadAllBytes($full))
    if ($existingHash -eq $newHash) {
      return [pscustomobject]@{ Action = 'SkipSame'; Path = $full; Hash = $newHash }
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $ext = [System.IO.Path]::GetExtension($FileName)

    for ($i = 1; $i -le 9999; $i++) {
      $candidateName = '{0}_new_{1:00}{2}' -f $baseName, $i, $ext
      $candidatePath = Join-Path $Directory $candidateName
      if (-not (Test-Path $candidatePath)) {
        return [pscustomobject]@{ Action = 'Save'; Path = $candidatePath; Hash = $newHash }
      }

      $candidateHash = Get-Sha256Hex -Bytes ([System.IO.File]::ReadAllBytes($candidatePath))
      if ($candidateHash -eq $newHash) {
        return [pscustomobject]@{ Action = 'SkipSame'; Path = $candidatePath; Hash = $newHash }
      }
    }

    throw 'Could not resolve unique file name after 9999 attempts.'
  }

  $saved = 0
  $skipped = 0
  $failed = 0
  $recordCount = 0
  $attachmentCount = 0

  $table = [string]$Context.table
  $dateField = [string]$Context.dateField
  $start = [datetime]$Context.startDateTime
  $end = [datetime]$Context.endDateTime
  $downloadDirectory = [string]$Context.downloadDirectory
  $createSubfolderPerTable = [bool]$Context.createSubfolderPerTable

  if (-not (Test-Path $downloadDirectory)) { [void](New-Item -Path $downloadDirectory -ItemType Directory -Force) }

  $query = "{0}BETWEENjavascript:gs.dateGenerate('{1}','{2}')@javascript:gs.dateGenerate('{3}','{4}')" -f `
    $dateField,
    $start.ToString('yyyy-MM-dd'), $start.ToString('HH:mm:ss'),
    $end.ToString('yyyy-MM-dd'), $end.ToString('HH:mm:ss')

  & $WriteLog ("Record filter query: {0}" -f $query)

  $recordMap = @{}
  $pageSize = 500
  $offset = 0

  while ($true) {
    $qs = @{
      sysparm_display_value = 'false'
      sysparm_exclude_reference_link = 'true'
      sysparm_fields = "sys_id,number,short_description,$dateField"
      sysparm_limit = $pageSize
      sysparm_offset = $offset
      sysparm_query = $query
    }

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($k in $qs.Keys) {
      [void]$parts.Add(("{0}={1}" -f $k, (& $UrlEncode ([string]$qs[$k]))))
    }

    $res = & $InvokeSnowGet ("/api/now/table/{0}?{1}" -f $table, ($parts -join '&'))
    $batch = if ($res -and ($res.PSObject.Properties.Name -contains 'result')) { @($res.result) } else { @() }

    foreach ($r in $batch) {
      $sysId = [string]$r.sys_id
      if ([string]::IsNullOrWhiteSpace($sysId)) { continue }
      $recordMap[$sysId] = [pscustomobject]@{
        sys_id = $sysId
        number = [string]$r.number
        short_description = [string]$r.short_description
      }
    }

    $recordCount += $batch.Count
    if ($batch.Count -lt $pageSize) { break }
    $offset += $pageSize
  }

  if ($recordMap.Count -eq 0) {
    return [pscustomobject]@{ Saved = 0; Skipped = 0; Failed = 0; Records = 0; Attachments = 0; Success = $true }
  }

  $recordIds = @($recordMap.Keys)
  $nameCounterByRecord = @{}

  for ($chunkStart = 0; $chunkStart -lt $recordIds.Count; $chunkStart += 100) {
    $chunkEnd = [Math]::Min($chunkStart + 99, $recordIds.Count - 1)
    $chunkIds = @($recordIds[$chunkStart..$chunkEnd])
    $inQuery = [string]::Join(',', $chunkIds)
    $attachQuery = "table_name={0}^table_sys_idIN{1}" -f $table, $inQuery

    $attachOffset = 0
    while ($true) {
      $qsAttach = @{
        sysparm_display_value = 'false'
        sysparm_exclude_reference_link = 'true'
        sysparm_fields = 'sys_id,table_name,table_sys_id,file_name'
        sysparm_limit = $pageSize
        sysparm_offset = $attachOffset
        sysparm_query = $attachQuery
      }
      $partsAttach = New-Object System.Collections.Generic.List[string]
      foreach ($k in $qsAttach.Keys) {
        [void]$partsAttach.Add(("{0}={1}" -f $k, (& $UrlEncode ([string]$qsAttach[$k]))))
      }

      $resAttach = & $InvokeSnowGet ("/api/now/table/sys_attachment?{0}" -f ($partsAttach -join '&'))
      $attachBatch = if ($resAttach -and ($resAttach.PSObject.Properties.Name -contains 'result')) { @($resAttach.result) } else { @() }

      foreach ($att in $attachBatch) {
        $attachmentCount++
        try {
          $attId = [string]$att.sys_id
          $tableName = [string]$att.table_name
          $tableSysId = [string]$att.table_sys_id
          $originalName = [string]$att.file_name

          if ([string]::IsNullOrWhiteSpace($attId) -or [string]::IsNullOrWhiteSpace($tableSysId)) {
            $failed++
            & $WriteLog "Attachment skipped due to missing sys_id/table_sys_id."
            continue
          }

          $record = $recordMap[$tableSysId]
          if ($null -eq $record) {
            $failed++
            & $WriteLog ("Attachment record not found in map: {0}" -f $tableSysId)
            continue
          }

          $recordKey = if (-not [string]::IsNullOrWhiteSpace([string]$record.number)) {
            [string]$record.number
          } elseif (-not [string]::IsNullOrWhiteSpace([string]$record.short_description)) {
            [string]$record.short_description
          } else {
            [string]$record.sys_id
          }

          $safeTable = Convert-ToSafeName $tableName
          $safeRecordKey = Convert-ToSafeName $recordKey
          $safeOriginal = Convert-ToSafeName $originalName

          $baseName = "{0}_{1}_{2}" -f $safeTable, $safeRecordKey, $safeOriginal
          $ext = [System.IO.Path]::GetExtension($baseName)
          $nameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($baseName)

          $counterKey = "{0}|{1}|{2}" -f $tableSysId, $nameNoExt, $ext
          if (-not $nameCounterByRecord.ContainsKey($counterKey)) {
            $nameCounterByRecord[$counterKey] = 0
          }
          $nameCounterByRecord[$counterKey] = [int]$nameCounterByRecord[$counterKey] + 1
          $dupNo = [int]$nameCounterByRecord[$counterKey]

          $fileName = if ($dupNo -gt 1) {
            "{0}_{1:00}{2}" -f $nameNoExt, $dupNo, $ext
          } else {
            "{0}{1}" -f $nameNoExt, $ext
          }

          $targetDir = $downloadDirectory
          if ($createSubfolderPerTable) {
            $targetDir = Join-Path $downloadDirectory $safeTable
          }
          if (-not (Test-Path $targetDir)) { [void](New-Item -Path $targetDir -ItemType Directory -Force) }

          $bytes = & $DownloadAttachmentBytes $attId
          if ($null -eq $bytes -or $bytes.Length -eq 0) {
            $failed++
            & $WriteLog ("Attachment download failed/empty: {0}" -f $attId)
            continue
          }

          $pathDecision = Resolve-UniquePath -Directory $targetDir -FileName $fileName -Content $bytes
          if ([string]$pathDecision.Action -eq 'SkipSame') {
            $skipped++
            & $WriteLog ("同一ファイルのためスキップ: {0}" -f [string]$pathDecision.Path)
            continue
          }

          [System.IO.File]::WriteAllBytes([string]$pathDecision.Path, $bytes)
          $saved++
          & $WriteLog ("Saved: {0}" -f [string]$pathDecision.Path)
        } catch {
          $failed++
          & $WriteLog ("Attachment process failed: {0}" -f $_.Exception.Message)
        }
      }

      if ($attachBatch.Count -lt $pageSize) { break }
      $attachOffset += $pageSize
    }
  }

  return [pscustomobject]@{
    Saved = $saved
    Skipped = $skipped
    Failed = $failed
    Records = $recordCount
    Attachments = $attachmentCount
    Success = ($failed -eq 0)
  }
}

Export-ModuleMember -Function Validate-AttachmentHarvesterInput, Invoke-AttachmentHarvesterUseCase
