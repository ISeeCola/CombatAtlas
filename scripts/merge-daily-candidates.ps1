param(
  [Parameter(Mandatory = $true)][string]$CandidateJson,
  [string]$ReviewPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'source\combat_atlas_review.xlsm')
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$pendingDir = Join-Path $root 'automation\pending'; New-Item -ItemType Directory -Force -Path $pendingDir | Out-Null
$lockPath = "$ReviewPath.lock"
$lock = $null; $excel = $null; $book = $null
try {
  try { $lock = [System.IO.File]::Open($lockPath, 'CreateNew', 'Write', 'None') }
  catch { Copy-Item -LiteralPath $CandidateJson -Destination (Join-Path $pendingDir ("daily-{0}.json" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))); throw '审核工作簿正被占用；候选已保存到 pending。' }
  $candidates = @(Get-Content -Raw -LiteralPath $CandidateJson | ConvertFrom-Json)
  if ($candidates.Count -gt 3) { throw '单次候选不得超过 3 条' }
  $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false
  $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $ReviewPath).Path)
  $table = $book.Worksheets.Item('待审批文章').ListObjects.Item('CombatAtlasInbox')
  $headers = @{}; for ($i=1; $i -le $table.ListColumns.Count; $i++) { $headers[$table.ListColumns.Item($i).Name]=$i }
  $seen = @{}; if ($table.DataBodyRange) { for ($r=1; $r -le $table.DataBodyRange.Rows.Count; $r++) { foreach($key in @('候选 ID','规范 URL','平台 ID')) { $value=[string]$table.DataBodyRange.Cells($r,$headers[$key]).Value2; if($value){$seen[$value.ToLowerInvariant()]=$true} } } }
  $added = 0
  foreach ($candidate in $candidates) {
    $canonical = [string]$candidate.canonicalUrl; $platform = [string]$candidate.platformId
    if (($canonical -and $seen.ContainsKey($canonical.ToLowerInvariant())) -or ($platform -and $seen.ContainsKey($platform.ToLowerInvariant()))) { continue }
    $row = $table.ListRows.Add().Range
    $candidateId = if ($candidate.candidateId) { [string]$candidate.candidateId } else { 'cand-' + (Get-Date -Format 'yyyyMMdd') + '-' + ([Guid]::NewGuid().ToString('N').Substring(0,8)) }
    $values = @{ '候选 ID'=$candidateId; '审核状态'='待人工复核'; '发现日期'=(Get-Date); '中文标题'=$candidate.title; '原标题'=$candidate.originalTitle; '作者'=$candidate.author; '来源'=$candidate.source; '原文 URL'=$candidate.url; '规范 URL'=$canonical; '平台 ID'=$platform; '发布年份'=$candidate.year; '语言'=$candidate.language; '媒介'=$candidate.medium; '建议来源层级'=$candidate.tier; '建议主题'=(@($candidate.topics) -join '；'); '证据状态'=$candidate.evidenceStatus; '短读证据'=$candidate.shortEvidence; '去重结果'='已与主表、review、catalog 和历史拒收记录去重'; '审核备注'=$candidate.reason; '评论'=$candidate.comment }
    foreach ($key in $values.Keys) { if ($headers.ContainsKey($key)) { $row.Cells(1,$headers[$key]).Value2=$values[$key] } }
    $seen[$canonical.ToLowerInvariant()]=$true; $added++
  }
  $book.Save(); Write-Output "已合并 $added 条候选到 review；未修改主表或发布网页。"
} finally {
  if ($book) { try { $book.Close($true) } catch {} }; if ($excel) { try { $excel.Quit() } catch {} }
  if ($lock) { $lock.Dispose(); Remove-Item -LiteralPath $lockPath -ErrorAction SilentlyContinue }
}
