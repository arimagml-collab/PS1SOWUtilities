Set-StrictMode -Version Latest

function Validate-ExportInput {
  param(
    [Parameter(Mandatory=$true)][string]$BaseUrl,
    [Parameter(Mandatory=$true)][string]$Table,
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

  if ($Settings.authType -eq "userpass") {
    $user = [string]$Settings.userId
    $pass = & $UnprotectSecret ([string]$Settings.passwordEnc)
    if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) {
      return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnAuth")) }
    }
  } else {
    $key = & $UnprotectSecret ([string]$Settings.apiKeyEnc)
    if ([string]::IsNullOrWhiteSpace($key)) {
      return [pscustomobject]@{ IsValid = $false; Errors = @((& $GetText "WarnAuth")) }
    }
  }

  return [pscustomobject]@{ IsValid = $true; Errors = @() }
}

function Invoke-ExportUseCase {
  param(
    [Parameter(Mandatory=$true)]$Context,
    [Parameter(Mandatory=$true)][scriptblock]$InvokeSnowGet,
    [Parameter(Mandatory=$true)][scriptblock]$UrlEncode
  )

  $offset = 0
  $total = 0
  $isFirstJson = $true
  $jsonWriter = $null
  $csvWriter = $null
  $all = New-Object System.Collections.Generic.List[object]

  try {
    if ($Context.format -eq "json") {
      $jsonWriter = New-Object System.IO.StreamWriter($Context.file, $false, (New-Object System.Text.UTF8Encoding($false)))
      $jsonWriter.Write("[")
    } elseif ($Context.format -eq "csv") {
      $csvWriter = New-Object System.IO.StreamWriter($Context.file, $false, (New-Object System.Text.UTF8Encoding($false)))
    }

    while ($true) {
      $qs = @{
        sysparm_limit  = $Context.pageSize
        sysparm_offset = $offset
        sysparm_display_value = "false"
        sysparm_exclude_reference_link = "true"
      }
      $queryParts = New-Object System.Collections.Generic.List[string]
      foreach ($k2 in $qs.Keys) { [void]$queryParts.Add(("{0}={1}" -f $k2, (& $UrlEncode ([string]$qs[$k2])))) }
      if (-not [string]::IsNullOrWhiteSpace([string]$Context.query)) { [void]$queryParts.Add(("sysparm_query={0}" -f (& $UrlEncode ([string]$Context.query)))) }
      if (-not [string]::IsNullOrWhiteSpace([string]$Context.fields)) { [void]$queryParts.Add(("sysparm_fields={0}" -f (& $UrlEncode ([string]$Context.fields)))) }

      $path = "/api/now/table/" + $Context.table + "?" + ($queryParts -join "&")
      $res = & $InvokeSnowGet $path
      $batchRes = if ($res -and ($res.PSObject.Properties.Name -contains "result")) { $res.result } else { @() }
      $batch = @($batchRes)

      foreach ($r in $batch) {
        if ($Context.format -eq "json") {
          $itemJson = ($r | ConvertTo-Json -Depth 10 -Compress)
          if (-not $isFirstJson) { $jsonWriter.Write(",") }
          $jsonWriter.Write($itemJson)
          $isFirstJson = $false
        } elseif ($Context.format -eq "csv") {
          $itemJson = ($r | ConvertTo-Json -Depth 10 -Compress).Replace('"','""')
          $csvWriter.WriteLine(("`"{0}`"" -f $itemJson))
        } else {
          $all.Add($r)
        }
      }

      $total += $batch.Count
      if ($batch.Count -lt $Context.pageSize) { break }
      $offset += $Context.pageSize
      if ($offset -gt 2000000) { break }
    }

    if ($Context.format -eq "xlsx") {
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
          $workbook.SaveAs($Context.file, 51)
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

  return [pscustomobject]@{ file=$Context.file; total=$total }
}

Export-ModuleMember -Function Validate-ExportInput, Invoke-ExportUseCase
