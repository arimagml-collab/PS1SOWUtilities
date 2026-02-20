Set-StrictMode -Version Latest

function Validate-ViewInput {
  param(
    [Parameter(Mandatory=$true)][string]$ViewName,
    [Parameter(Mandatory=$true)][string]$ViewLabel,
    [Parameter(Mandatory=$true)][string]$BaseTable,
    [Parameter(Mandatory=$true)][object[]]$JoinDefinitions,
    [Parameter(Mandatory=$true)][scriptblock]$GetText
  )

  if ([string]::IsNullOrWhiteSpace($ViewName)) {
    return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnViewName")) }
  }
  if ([string]::IsNullOrWhiteSpace($ViewLabel)) {
    return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnViewLabel")) }
  }
  if ([string]::IsNullOrWhiteSpace($BaseTable)) {
    return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnBaseTable")) }
  }

  foreach ($j in $JoinDefinitions) {
    if ([string]::IsNullOrWhiteSpace([string]$j.joinTable)) {
      return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnJoinTable")) }
    }
    if ([string]::IsNullOrWhiteSpace([string]$j.baseColumn)) {
      return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnJoinBaseColumn")) }
    }
    if ([string]::IsNullOrWhiteSpace([string]$j.targetColumn)) {
      return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnJoinTargetColumn")) }
    }
  }

  return [pscustomobject]@{ IsValid = $true; Errors = @() }
}

function Invoke-CreateViewUseCase {
  param(
    [Parameter(Mandatory=$true)]$Context,
    [Parameter(Mandatory=$true)][scriptblock]$InvokeSnowPost,
    [Parameter(Mandatory=$true)][scriptblock]$InvokeSnowPatch,
    [Parameter(Mandatory=$true)][scriptblock]$InvokeSnowGet,
    [Parameter(Mandatory=$true)][scriptblock]$UrlEncode,
    [Parameter(Mandatory=$true)][scriptblock]$SaveViewTableMetadata,
    [Parameter(Mandatory=$true)][scriptblock]$BuildJoinWhereClause,
    [Parameter(Mandatory=$true)][scriptblock]$TryCreateViewJoinRow
  )

  $body = @{ name = $Context.viewName; table = $Context.baseTable }
  if (@($Context.selectedColumns).Count -gt 0) { $body["view_field_list"] = (@($Context.selectedColumns) -join ",") }
  $createRes = & $InvokeSnowPost "/api/now/table/sys_db_view" $body
  $created = if ($createRes -and ($createRes.PSObject.Properties.Name -contains "result")) { $createRes.result } else { $null }
  $sysId = if ($created) { [string]$created.sys_id } else { "" }

  $joinsSaved = $true
  if (-not [string]::IsNullOrWhiteSpace($sysId)) {
    [void](& $InvokeSnowPatch ("/api/now/table/sys_db_view/{0}" -f $sysId) @{ label = $Context.viewLabel })

    if (@($Context.selectedColumns).Count -gt 0) {
      $fieldCsv = (@($Context.selectedColumns) -join ",")
      foreach ($fieldKey in @("view_fields", "field_names", "view_field_list")) {
        try { [void](& $InvokeSnowPatch ("/api/now/table/sys_db_view/{0}" -f $sysId) @{ $fieldKey = $fieldCsv }); break } catch {}
      }
    }

    $baseTableRowId = ""
    try {
      $query = "view={0}^table={1}" -f $sysId, $Context.baseTable
      $path = "/api/now/table/sys_db_view_table?sysparm_fields=sys_id&sysparm_limit=1&sysparm_query={0}" -f (& $UrlEncode $query)
      $baseTableRes = & $InvokeSnowGet $path
      $baseTableRow = if ($baseTableRes -and ($baseTableRes.PSObject.Properties.Name -contains "result") -and @($baseTableRes.result).Count -gt 0) { $baseTableRes.result[0] } else { $null }
      if ($baseTableRow) { $baseTableRowId = [string]$baseTableRow.sys_id }
    } catch {}

    if ([string]::IsNullOrWhiteSpace($baseTableRowId)) {
      try {
        $baseCreate = & $InvokeSnowPost "/api/now/table/sys_db_view_table" @{ view = $sysId; table = $Context.baseTable; order = 0; variable_prefix = $Context.basePrefix }
        if ($baseCreate -and ($baseCreate.PSObject.Properties.Name -contains "result") -and $baseCreate.result) { $baseTableRowId = [string]$baseCreate.result.sys_id }
      } catch {}
    }

    [void](& $SaveViewTableMetadata $baseTableRowId $Context.basePrefix "" $false $false)

    if (@($Context.joinDefs).Count -gt 0) {
      $joinsSaved = $false
      $joinIndex = 1
      foreach ($joinDef in @($Context.joinDefs)) {
        $joinPrefix = ([string]$joinDef.joinPrefix).Trim()
        if ([string]::IsNullOrWhiteSpace($joinPrefix)) { $joinPrefix = ("t{0}" -f $joinIndex) }
        $joinSource = ([string]$joinDef.joinSource).Trim()
        $leftPrefix = if ([string]::IsNullOrWhiteSpace($joinSource) -or $joinSource -eq "__base__") { $Context.basePrefix } else { $joinSource }
        $isLeftJoin = $false
        if ($joinDef.PSObject.Properties.Name -contains "leftJoin") { $isLeftJoin = [System.Convert]::ToBoolean($joinDef.leftJoin) }
        $joinWhereClause = & $BuildJoinWhereClause $leftPrefix ([string]$joinDef.baseColumn) $joinPrefix ([string]$joinDef.targetColumn)
        $joinOrder = $joinIndex * 100
        $joinCreate = & $TryCreateViewJoinRow $sysId $joinDef $joinWhereClause $joinPrefix $isLeftJoin $joinOrder
        if (-not [bool]$joinCreate.saved) { $joinsSaved = $false; break }
        if (-not [string]::IsNullOrWhiteSpace([string]$joinCreate.rowId)) { [void](& $SaveViewTableMetadata ([string]$joinCreate.rowId) $joinPrefix $joinWhereClause $isLeftJoin $true) }
        $joinIndex++
        $joinsSaved = $true
      }
    }
  }

  return [pscustomobject]@{ viewName = $Context.viewName; sysId = $sysId; joinsSaved = $joinsSaved }
}

Export-ModuleMember -Function Validate-ViewInput, Invoke-CreateViewUseCase
