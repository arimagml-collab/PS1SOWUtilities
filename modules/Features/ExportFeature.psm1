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

function Build-ExportQuery {
  param(
    [AllowEmptyString()][string]$BaseQuery,
    [AllowNull()][string]$LastCreatedOn,
    [AllowNull()][string]$LastSysId
  )

  $queryWithSort = if ([string]::IsNullOrWhiteSpace($BaseQuery)) {
    "ORDERBYsys_created_on^ORDERBYsys_id"
  } else {
    "{0}^ORDERBYsys_created_on^ORDERBYsys_id" -f $BaseQuery
  }

  if ([string]::IsNullOrWhiteSpace($LastCreatedOn) -or [string]::IsNullOrWhiteSpace($LastSysId)) {
    return $queryWithSort
  }

  $disjunct1 = "sys_created_on>{0}^ORDERBYsys_created_on^ORDERBYsys_id" -f $LastCreatedOn
  $disjunct2 = "sys_created_on={0}^sys_id>{1}^ORDERBYsys_created_on^ORDERBYsys_id" -f $LastCreatedOn, $LastSysId

  if ([string]::IsNullOrWhiteSpace($BaseQuery)) {
    return "{0}^NQ{1}" -f $disjunct1, $disjunct2
  }

  return "{0}^{1}^NQ{0}^{2}" -f $BaseQuery, $disjunct1, $disjunct2
}

function Invoke-ExportUseCase {
  param(
    [Parameter(Mandatory=$true)]$Context,
    [Parameter(Mandatory=$true)][scriptblock]$InvokeSnowGet,
    [Parameter(Mandatory=$true)][scriptblock]$UrlEncode
  )

  $total = 0
  $isFirstJson = $true
  $jsonWriter = $null
  $csvWriter = $null
  $csvFiles = New-Object System.Collections.Generic.List[string]
  $csvPartNo = 0
  $csvRowsInPart = 0
  $all = New-Object System.Collections.Generic.List[object]
  $lastCreatedOn = $null
  $lastSysId = $null

  function New-CsvPartWriter {
    param([int]$PartNo)

    $dir = [System.IO.Path]::GetDirectoryName([string]$Context.file)
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = "." }
    $name = [System.IO.Path]::GetFileNameWithoutExtension([string]$Context.file)
    $ext = [System.IO.Path]::GetExtension([string]$Context.file)
    $partFile = Join-Path $dir ("{0}-{1:000}{2}" -f $name, $PartNo, $ext)
    return [pscustomobject]@{ File = $partFile; Writer = (New-Object System.IO.StreamWriter($partFile, $false, (New-Object System.Text.UTF8Encoding($false)))) }
  }

  try {
    if ($Context.format -eq "json") {
      $jsonWriter = New-Object System.IO.StreamWriter($Context.file, $false, (New-Object System.Text.UTF8Encoding($false)))
      $jsonWriter.Write("[")
    } elseif ($Context.format -eq "csv") {
      $csvPartNo = 1
      $firstPart = New-CsvPartWriter -PartNo $csvPartNo
      $csvWriter = $firstPart.Writer
      [void]$csvFiles.Add([string]$firstPart.File)
    }

    while ($true) {
      $limit = [int]$Context.pageSize
      if ($Context.format -ne "csv") {
        $remaining = [int]$Context.maxRows - $total
        if ($remaining -le 0) { break }
        $limit = [Math]::Min([int]$Context.pageSize, $remaining)
      }
      $requestQuery = Build-ExportQuery -BaseQuery ([string]$Context.query) -LastCreatedOn $lastCreatedOn -LastSysId $lastSysId

      $qs = @{
        sysparm_limit  = $limit
        sysparm_display_value = "false"
        sysparm_exclude_reference_link = "true"
        sysparm_query = $requestQuery
      }
      if (-not [string]::IsNullOrWhiteSpace([string]$Context.fields)) { $qs.sysparm_fields = [string]$Context.fields }

      $queryParts = New-Object System.Collections.Generic.List[string]
      foreach ($k2 in $qs.Keys) { [void]$queryParts.Add(("{0}={1}" -f $k2, (& $UrlEncode ([string]$qs[$k2])))) }

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
          if ($csvRowsInPart -ge [int]$Context.maxRows) {
            if ($csvWriter) { $csvWriter.Dispose() }
            $csvPartNo++
            $nextPart = New-CsvPartWriter -PartNo $csvPartNo
            $csvWriter = $nextPart.Writer
            [void]$csvFiles.Add([string]$nextPart.File)
            $csvRowsInPart = 0
          }
          $itemJson = ($r | ConvertTo-Json -Depth 10 -Compress).Replace('"','""')
          $csvWriter.WriteLine(("`"{0}`"" -f $itemJson))
          $csvRowsInPart++
        } else {
          $all.Add($r)
        }

        $lastCreatedOn = [string]$r.sys_created_on
        $lastSysId = [string]$r.sys_id
      }

      $total += $batch.Count
      if ($batch.Count -lt $limit) { break }
      if ([string]::IsNullOrWhiteSpace($lastCreatedOn) -or [string]::IsNullOrWhiteSpace($lastSysId)) { break }
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

  if ($Context.format -eq "csv" -and $csvFiles.Count -eq 1) {
    $singleFile = [string]$csvFiles[0]
    if ($singleFile -ne [string]$Context.file) {
      if (Test-Path $Context.file) { Remove-Item -LiteralPath $Context.file -Force }
      Move-Item -LiteralPath $singleFile -Destination $Context.file
      $csvFiles.Clear()
      [void]$csvFiles.Add([string]$Context.file)
    }
  }

  return [pscustomobject]@{ file=$Context.file; files=@($csvFiles); total=$total }
}

Export-ModuleMember -Function Validate-ExportInput, Invoke-ExportUseCase
