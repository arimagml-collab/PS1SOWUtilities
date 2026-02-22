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
  $csvHeaderWritten = $false
  $csvColumns = @()
  $all = New-Object System.Collections.Generic.List[object]
  $lastCreatedOn = $null
  $lastSysId = $null

  function Resolve-OutputEncoding {
    param([Parameter(Mandatory=$true)]$State)

    $format = ([string]$State.format).Trim().ToLowerInvariant()
    $rawName = ([string]$State.outputEncoding).Trim()
    $normalized = $rawName.ToLowerInvariant()

    $useBom = $false
    if ($State.PSObject.Properties.Name -contains 'outputBom') {
      $useBom = [bool]$State.outputBom
    } elseif ($format -eq 'csv') {
      # Excel / ServiceNow import workflows often mis-detect UTF-8 without BOM.
      $useBom = $true
    }

    if ([string]::IsNullOrWhiteSpace($normalized)) {
      $normalized = 'utf-8'
    }

    switch ($normalized) {
      { $_ -in @('utf8', 'utf-8', 'utf8bom', 'utf-8-bom') } {
        return [pscustomobject]@{
          Name = 'utf-8'
          Encoding = (New-Object System.Text.UTF8Encoding($useBom))
          Bom = $useBom
        }
      }
      { $_ -in @('sjis', 'shift-jis', 'shift_jis', 'cp932', 'ms932', 'windows-31j') } {
        return [pscustomobject]@{
          Name = 'shift_jis'
          Encoding = [System.Text.Encoding]::GetEncoding(932)
          Bom = $false
        }
      }
      default {
        return [pscustomobject]@{
          Name = 'utf-8'
          Encoding = (New-Object System.Text.UTF8Encoding($useBom))
          Bom = $useBom
        }
      }
    }
  }

  $encodingMeta = Resolve-OutputEncoding -State $Context

  function New-CsvPartWriter {
    param([int]$PartNo)

    $dir = [System.IO.Path]::GetDirectoryName([string]$Context.file)
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = "." }
    $name = [System.IO.Path]::GetFileNameWithoutExtension([string]$Context.file)
    $ext = [System.IO.Path]::GetExtension([string]$Context.file)
    $partFile = Join-Path $dir ("{0}-{1:000}{2}" -f $name, $PartNo, $ext)
    return [pscustomobject]@{ File = $partFile; Writer = (New-Object System.IO.StreamWriter($partFile, $false, $encodingMeta.Encoding)) }
  }

  function Convert-ToCsvText {
    param(
      [Parameter(Mandatory=$true)][string[]]$Columns,
      [Parameter(Mandatory=$true)]$Source
    )

    function Convert-ServiceNowFieldValue {
      param([AllowNull()]$Value)

      if ($null -eq $Value) { return '' }
      if ($Value -is [string] -or $Value -is [ValueType]) { return [string]$Value }

      $valueProp = $null
      try { $valueProp = $Value.PSObject.Properties['value'] } catch { $valueProp = $null }
      if ($valueProp) {
        return (Convert-ServiceNowFieldValue -Value $valueProp.Value)
      }

      if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = New-Object System.Collections.Generic.List[string]
        foreach ($entry in $Value) {
          $resolved = Convert-ServiceNowFieldValue -Value $entry
          if (-not [string]::IsNullOrWhiteSpace($resolved)) {
            [void]$items.Add($resolved)
          }
        }
        if ($items.Count -gt 0) {
          return ($items -join ";")
        }
      }

      return ($Value | ConvertTo-Json -Depth 10 -Compress)
    }

    function Convert-CsvField {
      param([AllowNull()][string]$Text)

      if ($null -eq $Text) { return '' }
      $needsQuote = ($Text.Contains('"') -or $Text.Contains(',') -or $Text.Contains("`r") -or $Text.Contains("`n"))
      if ($needsQuote) {
        return ('"{0}"' -f ($Text -replace '"', '""'))
      }
      return $Text
    }

    $line = foreach ($col in $Columns) {
      $v = $null
      try { $v = $Source.$col } catch { $v = $null }
      Convert-ServiceNowFieldValue -Value $v
    }

    return ((@($line) | ForEach-Object { Convert-CsvField -Text ([string]$_) }) -join ',')
  }

  function Convert-ToCsvHeaderText {
    param([Parameter(Mandatory=$true)][string[]]$Columns)

    return ((@($Columns) | ForEach-Object {
      if ($_.Contains('"') -or $_.Contains(',') -or $_.Contains("`r") -or $_.Contains("`n")) {
        ('"{0}"' -f ($_ -replace '"', '""'))
      } else {
        $_
      }
    }) -join ',')
  }

  try {
    if ($Context.format -eq "json") {
      $jsonWriter = New-Object System.IO.StreamWriter($Context.file, $false, $encodingMeta.Encoding)
      $jsonWriter.Write("[")
    } elseif ($Context.format -eq "csv") {
      $csvPartNo = 1
      $firstPart = New-CsvPartWriter -PartNo $csvPartNo
      $csvWriter = $firstPart.Writer
      [void]$csvFiles.Add([string]$firstPart.File)
      $csvHeaderWritten = $false
      $csvColumns = @()

      $rawFields = [string]$Context.fields
      if (-not [string]::IsNullOrWhiteSpace($rawFields)) {
        $csvColumns = @($rawFields.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
      }

      if ($csvColumns.Count -gt 0) {
        $csvWriter.WriteLine((Convert-ToCsvHeaderText -Columns $csvColumns))
        $csvHeaderWritten = $true
      }
    }

    while ($true) {
      $limit = [int]$Context.pageSize
      if ($Context.format -eq "csv") {
        $remaining = [int]$Context.maxRows - $csvRowsInPart
        if ($remaining -le 0) { $remaining = [int]$Context.maxRows }
        $limit = [Math]::Min([int]$Context.pageSize, $remaining)
      }
      $requestQuery = Build-ExportQuery -BaseQuery ([string]$Context.query) -LastCreatedOn $lastCreatedOn -LastSysId $lastSysId

      $qs = @{
        sysparm_display_value = "false"
        sysparm_exclude_reference_link = "true"
        sysparm_query = $requestQuery
      }
      if ($Context.format -eq "csv") {
        $qs.sysparm_limit = $limit
      }
      if (-not [string]::IsNullOrWhiteSpace([string]$Context.fields)) { $qs.sysparm_fields = [string]$Context.fields }

      $queryParts = New-Object System.Collections.Generic.List[string]
      foreach ($k2 in $qs.Keys) { [void]$queryParts.Add(("{0}={1}" -f $k2, (& $UrlEncode ([string]$qs[$k2])))) }

      $path = "/api/now/table/" + $Context.table + "?" + ($queryParts -join "&")
      $res = & $InvokeSnowGet $path
      $batchRes = if ($res -and ($res.PSObject.Properties.Name -contains "result")) { $res.result } else { @() }
      $batch = New-Object System.Collections.Generic.List[object]
      foreach ($item in @($batchRes)) {
        if ($item -is [System.Array]) {
          foreach ($nestedItem in $item) { [void]$batch.Add($nestedItem) }
        } else {
          [void]$batch.Add($item)
        }
      }

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

            if ($csvColumns.Count -gt 0) {
              $csvWriter.WriteLine((Convert-ToCsvHeaderText -Columns $csvColumns))
              $csvHeaderWritten = $true
            } else {
              $csvHeaderWritten = $false
            }
          }

          if (-not $csvHeaderWritten) {
            $csvColumns = @($r.PSObject.Properties.Name)
            if ($csvColumns.Count -gt 0) {
              $csvWriter.WriteLine((Convert-ToCsvHeaderText -Columns $csvColumns))
              $csvHeaderWritten = $true
            }
          }

          if ($csvColumns.Count -gt 0) {
            $csvLine = Convert-ToCsvText -Columns $csvColumns -Source $r
            $csvWriter.WriteLine($csvLine)
          }
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

  return [pscustomobject]@{ file=$Context.file; files=@($csvFiles); total=$total; outputEncoding=$encodingMeta.Name; outputBom=$encodingMeta.Bom }
}

Export-ModuleMember -Function Validate-ExportInput, Invoke-ExportUseCase
