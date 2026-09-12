param(
  [string]$ReviewPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'source\combat_atlas_review.xlsm'),
  [string]$MainPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'source\combat_atlas_main.xlsm'),
  [string]$ResultPath = '',
  [string]$ReviewUpdatePath = '',
  [switch]$SyncDocsOnly
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Write-Result([string]$Text) {
  if ($ResultPath) { [System.IO.File]::WriteAllText($ResultPath, $Text, [System.Text.Encoding]::Unicode) }
}
function Clean-Text($Value) { if ($null -eq $Value) { return '' }; return ([string]$Value).Trim() }
function Excel-Multiline-Text($Value) { return (Clean-Text $Value).Replace("`r`n", "`n").Replace("`r", "`n") }
function Clean-Tsv($Value) { return (Clean-Text $Value).Replace("`t", ' ').Replace("`r", '').Replace("`n", '\n') }
function Normalize-Url([string]$Value) {
  $text = (Clean-Text $Value).TrimEnd('/')
  if (-not $text) { return '' }
  try { $uri = [Uri]$text; return (($uri.Scheme + '://' + $uri.Host + $uri.AbsolutePath.TrimEnd('/') + $uri.Query).ToLowerInvariant()) } catch { return $text.ToLowerInvariant() }
}
function New-SourceId { return 'src-' + (Get-Date -Format 'yyyyMMddHHmmss') + '-' + ([Guid]::NewGuid().ToString('N').Substring(0,8)) }

if ($SyncDocsOnly) {
  try {
    $env:COMBAT_ATLAS_REVIEW = $ReviewPath
    & node (Join-Path $PSScriptRoot 'sync-curation-docs.mjs') | Out-Null
    if ($LASTEXITCODE -ne 0) { throw '文档同步脚本返回失败。' }
    $message = '本地主表与 review 已更新；策展文档同步完成。尚未发布网页。'
    Write-Result $message; Write-Output $message; exit 0
  } catch {
    $message = '文章已入库且 review 状态已保存，但文档同步失败，可再次点击批准入库重试：' + $_.Exception.Message
    Write-Result $message; [Console]::Error.WriteLine($message); exit 2
  }
}

$runtimeDir = [System.IO.Path]::GetTempPath()
$acceptedPath = Join-Path $runtimeDir ("CombatAtlas-accepted-{0}.json" -f [Guid]::NewGuid().ToString('N'))
$validationReportPath = Join-Path $runtimeDir ("CombatAtlas-validation-{0}.json" -f [Guid]::NewGuid().ToString('N'))
$env:COMBAT_ATLAS_REVIEW = $ReviewPath
$excel = $null; $book = $null
try {
  $previous = $ErrorActionPreference
  try {
    $ErrorActionPreference = 'Continue'
    & node (Join-Path $PSScriptRoot 'review-workbook.mjs') --accepted-json $acceptedPath --report-json $validationReportPath 2>$null | Out-Null
    $validationExitCode = $LASTEXITCODE
  } finally { $ErrorActionPreference = $previous }
  if ($validationExitCode -ne 0) {
    if (Test-Path -LiteralPath $validationReportPath) {
      $report = Get-Content -Raw -Encoding UTF8 -LiteralPath $validationReportPath | ConvertFrom-Json
      $message = [string]$report.title + [Environment]::NewLine + [string]::Join([Environment]::NewLine, [string[]]@($report.errors))
    } else { $message = '审核技术校验失败，但未生成诊断报告。请运行 npm run review:validate。' }
    Write-Result $message; [Console]::Error.WriteLine($message); exit 1
  }
  $payload = Get-Content -Raw -Encoding UTF8 -LiteralPath $acceptedPath | ConvertFrom-Json
  $items = @($payload.accepted)
  if ($items.Count -eq 0) {
    $message = '没有状态为“已接受”的条目。已入库条目无需重复处理。'
    Write-Result $message; Write-Output $message; return
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false; $excel.DisplayAlerts = $false
  $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $MainPath).Path)
  $table = $book.Worksheets.Item('知识库文章').ListObjects.Item('CombatAtlasSources')
  $headers = @{}; for ($i = 1; $i -le $table.ListColumns.Count; $i++) { $headers[$table.ListColumns.Item($i).Name] = $i }
  $requiredHeaders = @('发布状态','中文标题','原标题','作者','来源','原文 URL','发布年份','语言','媒介','来源层级','主题','简介','核心结论','阅读时间（分钟）','策展价值（1-5）','精选状态','收录日期','sourceId')
  $missingHeaders = @($requiredHeaders | Where-Object { -not $headers.ContainsKey($_) })
  if ($missingHeaders.Count) { throw '主表结构损坏，缺少列：' + ($missingHeaders -join '、') }

  $existingById = @{}; $existingByUrl = @{}
  if ($table.DataBodyRange) {
    for ($r = 1; $r -le $table.DataBodyRange.Rows.Count; $r++) {
      $id = Clean-Text $table.DataBodyRange.Cells($r,$headers['sourceId']).Value2
      $url = Normalize-Url (Clean-Text $table.DataBodyRange.Cells($r,$headers['原文 URL']).Value2)
      if ($id) { $existingById[$id] = $r }
      if ($url) { if ($existingByUrl.ContainsKey($url)) { throw "主表存在重复规范 URL：$url" }; $existingByUrl[$url] = $id }
    }
  }

  $prepared = @(); $technicalErrors = @()
  foreach ($item in $items) {
    $row = $item.row
    $candidateId = Clean-Text $item.candidateId
    $sourceId = Clean-Text $item.sourceId
    $canonical = Normalize-Url (Clean-Text $row.'规范 URL')
    if (-not $sourceId -and $canonical -and $existingByUrl.ContainsKey($canonical)) { $sourceId = [string]$existingByUrl[$canonical] }
    if (-not $sourceId) { $sourceId = New-SourceId }
    $existingRow = if ($existingById.ContainsKey($sourceId)) { [int]$existingById[$sourceId] } else { 0 }
    if ($canonical -and $existingByUrl.ContainsKey($canonical) -and [string]$existingByUrl[$canonical] -ne $sourceId) {
      $technicalErrors += "「$($row.'中文标题')」的规范 URL 已属于 sourceId $($existingByUrl[$canonical])，但 review 指定为 $sourceId。"
      continue
    }
    $values = @{
      '发布状态'='发布'; '中文标题'=Clean-Text $row.'中文标题'; '原标题'=Clean-Text $row.'原标题'; '作者'=Clean-Text $row.'作者'; '来源'=Clean-Text $row.'来源'; '原文 URL'=Clean-Text $row.'规范 URL';
      '发布年份'=Clean-Text $row.'发布年份'; '语言'=Clean-Text $row.'语言'; '媒介'=Clean-Text $row.'媒介'; '来源层级'=Clean-Text $row.'建议来源层级'; '主题'=Clean-Text $row.'建议主题'; '简介'=Clean-Text $row.'摘要';
      '核心结论'=Excel-Multiline-Text $row.'核心方法'; '阅读时间（分钟）'=Clean-Text $row.'阅读时间（分钟）'; '精选状态'=Clean-Text $row.'精选状态'; '收录日期'=Clean-Text $row.'入库日期'; 'sourceId'=$sourceId
    }
    $values['策展价值（1-5）'] = Clean-Text $row.'展示价值（1-5）'
    if ($existingRow -gt 0) {
      $reviewValues = @{}
      foreach ($key in $values.Keys) {
        if ($key -in @('发布状态','sourceId')) { continue }
        if (Clean-Text $values[$key]) { $reviewValues[$key] = $values[$key] }
      }
      $prepared += [pscustomobject]@{ CandidateId=$candidateId; SourceId=$sourceId; Existing=$true; ExistingRow=$existingRow; Item=$item; Values=$reviewValues }
      if ($canonical) { $existingByUrl[$canonical] = $sourceId }
      continue
    }
    $mandatory = @('中文标题','作者','来源','原文 URL','发布年份','语言','媒介','来源层级','主题','简介','核心结论','阅读时间（分钟）','策展价值（1-5）','精选状态','收录日期','sourceId')
    $missing = @($mandatory | Where-Object { -not (Clean-Text $values[$_]) })
    if ($missing.Count) { $technicalErrors += "「$($row.'中文标题')」新增主表行缺少网页必需字段：$($missing -join '、')"; continue }
    try { $uri = [Uri]$values['原文 URL']; if ($uri.Scheme -notin @('http','https')) { throw 'bad' } } catch { $technicalErrors += "「$($row.'中文标题')」的原文 URL 非法：$($values['原文 URL'])"; continue }
    $prepared += [pscustomobject]@{ CandidateId=$candidateId; SourceId=$sourceId; Existing=$false; ExistingRow=0; Item=$item; Values=$values }
    $existingById[$sourceId] = -1; if ($canonical) { $existingByUrl[$canonical] = $sourceId }
  }
  if ($technicalErrors.Count) {
    $message = '入库技术校验未通过：' + [Environment]::NewLine + [string]::Join([Environment]::NewLine, $technicalErrors)
    Write-Result $message; [Console]::Error.WriteLine($message); exit 1
  }

  $added = 0; $existingCount = 0; $updatedFields = 0; $updates = New-Object System.Collections.Generic.List[string]
  $updates.Add("#columns`t$($payload.columns.'候选 ID')`t$($payload.columns.sourceId)`t$($payload.columns.'审核状态')`t$($payload.columns.'审核备注')")
  foreach ($entry in $prepared) {
    if ($entry.Existing) {
      $existingCount++
      $targetRow = $table.DataBodyRange.Rows.Item($entry.ExistingRow)
      foreach ($key in $entry.Values.Keys) {
        $oldValue = Clean-Text $targetRow.Cells(1,$headers[$key]).Value2
        $newValue = Clean-Text $entry.Values[$key]
        if ($newValue -and $oldValue -ne $newValue) {
          $targetRow.Cells(1,$headers[$key]).Value2 = $entry.Values[$key]
          $updatedFields++
        }
      }
    }
    else {
      $newRow = $table.ListRows.Add().Range
      foreach ($key in $entry.Values.Keys) { $newRow.Cells(1,$headers[$key]).Value2 = $entry.Values[$key] }
      $added++
    }
    $updates.Add((Clean-Tsv $entry.CandidateId) + "`t" + (Clean-Tsv $entry.SourceId) + "`t已入库`t" + (Clean-Tsv $entry.Item.finalReviewNote))
  }
  if ($added -gt 0 -or $updatedFields -gt 0) { $book.Save() }
  $book.Close($false); $book = $null; $excel.Quit(); $excel = $null
  if ($ReviewUpdatePath) { [System.IO.File]::WriteAllLines($ReviewUpdatePath, $updates, [System.Text.Encoding]::Unicode) }
  $message = "主表新增 $added 条，匹配已有 $existingCount 条并采用 Review 非空内容更新 $updatedFields 个字段；请由 Excel 完成 review 状态写回。尚未发布网页。"
  Write-Result $message; Write-Output $message
} catch {
  $message = $_.Exception.Message
  Write-Result $message; [Console]::Error.WriteLine($message); exit 1
} finally {
  Remove-Item -LiteralPath $acceptedPath -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $validationReportPath -ErrorAction SilentlyContinue
  if ($book) { try { $book.Close($false) } catch {} }
  if ($excel) { try { $excel.Quit() } catch {} }
}
