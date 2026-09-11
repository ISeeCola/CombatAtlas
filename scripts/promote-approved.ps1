param(
  [string]$ReviewPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'source\combat_atlas_review.xlsm'),
  [string]$MainPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'source\combat_atlas_main.xlsm')
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$pendingDir = Join-Path $root 'automation\pending'
New-Item -ItemType Directory -Force -Path $pendingDir | Out-Null
$acceptedPath = Join-Path $pendingDir ("accepted-{0}.json" -f [Guid]::NewGuid().ToString('N'))
$env:COMBAT_ATLAS_REVIEW = $ReviewPath
try {
  & node (Join-Path $PSScriptRoot 'review-workbook.mjs') --accepted-json $acceptedPath
  if ($LASTEXITCODE -ne 0) { throw '审核表校验失败' }
  $rows = Get-Content -Raw -LiteralPath $acceptedPath | ConvertFrom-Json
  if (@($rows).Count -eq 0) { Write-Output '没有状态为“已接受”的条目。'; exit 0 }
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false; $excel.DisplayAlerts = $false
  $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $MainPath).Path)
  $table = $book.Worksheets.Item('知识库文章').ListObjects.Item('CombatAtlasSources')
  $headers = @{}; for ($i = 1; $i -le $table.ListColumns.Count; $i++) { $headers[$table.ListColumns.Item($i).Name] = $i }
  $existing = @{}; if ($table.DataBodyRange) { for ($r = 1; $r -le $table.DataBodyRange.Rows.Count; $r++) { $existing[[string]$table.DataBodyRange.Cells($r,$headers['sourceId']).Value2] = $true } }
  $added = 0
  foreach ($row in @($rows)) {
    $sourceId = [string]$row.sourceId
    if (-not $sourceId) { $sourceId = 'src-' + (Get-Date -Format 'yyyyMMddHHmmss') + '-' + ([Guid]::NewGuid().ToString('N').Substring(0,8)) }
    if ($existing.ContainsKey($sourceId)) { continue }
    $newRow = $table.ListRows.Add().Range
    $values = @{
      '发布状态'='发布'; '中文标题'=$row.'中文标题'; '原标题'=$row.'原标题'; '作者'=$row.'作者'; '来源'=$row.'来源'; '原文 URL'=$row.'规范 URL';
      '发布年份'=$row.'发布年份'; '语言'=$row.'语言'; '媒介'=$row.'媒介'; '来源层级'=$row.'建议来源层级'; '主题'=$row.'建议主题'; '简介'=$row.'摘要';
      '核心结论'=$row.'核心方法'; '阅读时间（分钟）'=$row.'阅读时间（分钟）'; '策展价值（1-5）'=$row.'展示价值（1-5）'; '精选状态'=$row.'精选状态'; '收录日期'=$row.'入库日期'; 'sourceId'=$sourceId
    }
    foreach ($key in $values.Keys) { if ($headers.ContainsKey($key)) { $newRow.Cells(1,$headers[$key]).Value2 = $values[$key] } }
    $existing[$sourceId] = $true; $added++
  }
  $book.Save(); $book.Close($true); $excel.Quit()
  & node (Join-Path $PSScriptRoot 'sync-curation-docs.mjs')
  Write-Output "已迁移 $added 条到本地主表；尚未发布网页。"
} finally {
  Remove-Item -LiteralPath $acceptedPath -ErrorAction SilentlyContinue
  if ($book) { try { $book.Close($false) } catch {} }
  if ($excel) { try { $excel.Quit() } catch {} }
}
