Set-StrictMode -Version Latest

function Validate-TruncateInput {
  param(
    [Parameter(Mandatory=$true)][string]$Table,
    [Parameter(Mandatory=$true)][int]$MaxRetries,
    [Parameter(Mandatory=$true)][string]$ExpectedCode,
    [Parameter(Mandatory=$true)][string]$InputCode,
    [Parameter(Mandatory=$true)][scriptblock]$GetText
  )

  if ([string]::IsNullOrWhiteSpace($Table)) {
    return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnTable")) }
  }
  if ($MaxRetries -lt 1 -or $MaxRetries -gt 999) {
    return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "DeleteMaxRetriesInvalid")) }
  }
  if ($ExpectedCode -ne $InputCode) {
    return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "DeleteCodeMismatch")) }
  }

  return [pscustomobject]@{ IsValid = $true; Errors = @() }
}

function Invoke-TruncateUseCase {
  param(
    [Parameter(Mandatory=$true)]$Context,
    [Parameter(Mandatory=$true)][scriptblock]$InvokeSnowGet,
    [Parameter(Mandatory=$true)][scriptblock]$InvokeSnowDelete,
    [Parameter(Mandatory=$true)][scriptblock]$InvokeSnowBatchDelete,
    [Parameter(Mandatory=$true)][scriptblock]$WriteLog,
    [Parameter(Mandatory=$true)][scriptblock]$SetProgress,
    [Parameter(Mandatory=$true)][scriptblock]$GetText
  )

  & $WriteLog (& $GetText "DeleteFetchCount")
  & $SetProgress 0 (& $GetText "DeleteFetchCount")

  $countPath = "/api/now/stats/{0}?sysparm_count=true" -f $Context.table
  $countRes = & $InvokeSnowGet $countPath
  $initialTotal = 0
  try {
    if ($countRes -and $countRes.result -and $countRes.result.stats -and $countRes.result.stats.count) {
      $initialTotal = [int]$countRes.result.stats.count
    }
  } catch {
    $initialTotal = 0
  }

  if ($initialTotal -le 0) {
    & $WriteLog (& $GetText "DeleteNoRecord")
    & $SetProgress 100 "100% (0/0)"
    return [pscustomobject]@{ Status = "NoRecord"; Deleted = 0; InitialTotal = 0 }
  }

  $deleted = 0
  $attempt = 0
  while ($attempt -lt $Context.maxRetries) {
    $attempt++

    $listPath = "/api/now/table/{0}?sysparm_fields=sys_id&sysparm_limit=10000&sysparm_display_value=false&sysparm_exclude_reference_link=true" -f $Context.table
    $listRes = & $InvokeSnowGet $listPath
    $rows = if ($listRes -and ($listRes.PSObject.Properties.Name -contains "result")) { @($listRes.result) } else { @() }

    if ($rows.Count -eq 0) {
      & $WriteLog (("{0}: {1}" -f (& $GetText "DeleteDone"), $Context.table))
      & $SetProgress 100 ("100% ({0}/{1})" -f [Math]::Min($deleted, $initialTotal), $initialTotal)
      return [pscustomobject]@{ Status = "Done"; Deleted = $deleted; InitialTotal = $initialTotal }
    }

    $ids = @()
    foreach ($r in $rows) {
      $id = [string]$r.sys_id
      if ([string]::IsNullOrWhiteSpace($id)) { continue }
      $ids += $id
    }

    $batchSize = 50
    for ($i = 0; $i -lt $ids.Count; $i += $batchSize) {
      $take = [Math]::Min($batchSize, $ids.Count - $i)
      $chunk = $ids[$i..($i + $take - 1)]
      try {
        $batchResult = & $InvokeSnowBatchDelete $Context.table $chunk
        $ok = [int]$batchResult.deletedCount
        $deleted += $ok
        $failedIds = if ($batchResult -and ($batchResult.PSObject.Properties.Name -contains "failedIds")) { @($batchResult.failedIds) } else { @() }
        if ($failedIds.Count -gt 0) {
          foreach ($failedId in $failedIds) {
            try {
              [void](& $InvokeSnowDelete ("/api/now/table/{0}/{1}" -f $Context.table, $failedId))
              $deleted++
            } catch {
              & $WriteLog (("{0}: {1}" -f (& $GetText "Failed"), $_.Exception.Message))
            }
          }
        }
      } catch {
        foreach ($id in $chunk) {
          try {
            [void](& $InvokeSnowDelete ("/api/now/table/{0}/{1}" -f $Context.table, $id))
            $deleted++
          } catch {
            & $WriteLog (("{0}: {1}" -f (& $GetText "Failed"), $_.Exception.Message))
          }
        }
      }

      $pct = [int]([Math]::Floor(([Math]::Min($deleted, $initialTotal) * 100.0) / $initialTotal))
      & $SetProgress $pct ("{0}% ({1}/{2})" -f $pct, [Math]::Min($deleted, $initialTotal), $initialTotal)
    }

    & $WriteLog (("{0} {1}/{2}" -f (& $GetText "DeleteRetry"), $attempt, $Context.maxRetries))
  }

  & $WriteLog (& $GetText "DeleteStopped")
  return [pscustomobject]@{ Status = "Stopped"; Deleted = $deleted; InitialTotal = $initialTotal }
}

Export-ModuleMember -Function Validate-TruncateInput, Invoke-TruncateUseCase
