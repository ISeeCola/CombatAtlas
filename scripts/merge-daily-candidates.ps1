param(
  [Parameter(Mandatory = $true)][string]$CandidateJson,
  [string]$ReviewPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'source\combat_atlas_review.xlsm')
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$pendingDir = Join-Path $root 'automation\pending'
New-Item -ItemType Directory -Force -Path $pendingDir | Out-Null
$lockPath = "$ReviewPath.lock"
$lock = $null; $excel = $null; $book = $null

function Clean($Value) { if ($null -eq $Value) { return '' }; return ([string]$Value).Trim() }
function Comment-Hash([string]$Value) {
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($Value)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  try { return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant() } finally { $sha.Dispose() }
}
function Set-Cell($range, $headers, [string]$name, $value) {
  if (-not $headers.ContainsKey($name)) { return }
  if ($value -is [datetime]) { $value = $value.ToString('yyyy-MM-dd') }
  $range.Cells(1,$headers[$name]).Value2 = [string]$value
}

try {
  try { $lock = [System.IO.File]::Open($lockPath, 'CreateNew', 'Write', 'None') }
  catch {
    $pending = Join-Path $pendingDir ("daily-{0}.json" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    Copy-Item -LiteralPath $CandidateJson -Destination $pending
    throw "审核工作簿正被占用；输入已保存到 $pending。"
  }
  $raw = Get-Content -Raw -Encoding UTF8 -LiteralPath $CandidateJson | ConvertFrom-Json
  if ($raw -is [System.Array]) { $reworkResults = @(); $newCandidates = @($raw) }
  elseif ($null -ne $raw.newCandidates -or $null -ne $raw.reworkResults) { $reworkResults = @($raw.reworkResults); $newCandidates = @($raw.newCandidates) }
  else { $reworkResults = @(); $newCandidates = @($raw) }
  if ($reworkResults.Count -gt 3) { throw '单次待修改重做不得超过 3 条' }
  if ($newCandidates.Count -gt 3) { throw '单次新增候选不得超过 3 条' }

  $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false
  $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $ReviewPath).Path)
  $table = $book.Worksheets.Item('待审批文章').ListObjects.Item('CombatAtlasInbox')
  $headers = @{}; for ($i=1; $i -le $table.ListColumns.Count; $i++) { $headers[$table.ListColumns.Item($i).Name]=$i }
  foreach ($required in @('候选 ID','审核状态','评论','审核备注','规范 URL','平台 ID')) { if (-not $headers.ContainsKey($required)) { throw "review 表缺少列：$required" } }

  $rowsByCandidate = @{}; $seen = @{}
  if ($table.DataBodyRange) {
    for ($r=1; $r -le $table.DataBodyRange.Rows.Count; $r++) {
      $candidateId = Clean $table.DataBodyRange.Cells($r,$headers['候选 ID']).Value2
      if ($candidateId) { $rowsByCandidate[$candidateId] = $r; $seen[$candidateId.ToLowerInvariant()] = $true }
      foreach($key in @('规范 URL','平台 ID')) { $value=Clean $table.DataBodyRange.Cells($r,$headers[$key]).Value2; if($value){$seen[$value.ToLowerInvariant()]=$true} }
    }
  }

  $reworked = 0; $skippedRework = 0
  $allowedUpdateFields = @('中文标题','原标题','作者','来源','原文 URL','规范 URL','平台 ID','发布年份','语言','媒介','建议来源层级','建议主题','证据状态','短读证据','深读证据','去重结果','置信度','摘要','核心方法','证据与案例','术语对照','短摘录','适用边界','库内连接','实验落点','可迁移假设','最小验证动作','观察指标','失败信号','归档路径','阅读时间（分钟）','展示价值（1-5）','精选状态','入库日期')
  foreach ($result in $reworkResults) {
    $candidateId = Clean $result.candidateId
    if (-not $rowsByCandidate.ContainsKey($candidateId)) { throw "待修改记录不存在：$candidateId" }
    $rowIndex = $rowsByCandidate[$candidateId]
    $range = $table.DataBodyRange.Rows.Item($rowIndex)
    $status = Clean $range.Cells(1,$headers['审核状态']).Value2
    if ($status -ne '待修改') { throw "$candidateId 当前不是待修改状态，而是：$status" }
    $comment = Clean $range.Cells(1,$headers['评论']).Value2
    if (-not $comment) { throw "$candidateId 缺少人工评论，无法执行重做" }
    $actualHash = Comment-Hash $comment
    $submittedHash = Clean $result.commentSha256
    if ($submittedHash -and $submittedHash -ne $actualHash) { throw "$candidateId 的评论已变化，请重新读取后处理" }
    $note = Clean $range.Cells(1,$headers['审核备注']).Value2
    if ($note -match [regex]::Escape("commentSha256=$actualHash")) { $skippedRework++; continue }
    $outcome = (Clean $result.outcome).ToLowerInvariant()
    if ($outcome -notin @('resolved','unresolved')) { throw "$candidateId 的 outcome 必须为 resolved 或 unresolved" }
    if ($result.updates) {
      foreach ($property in $result.updates.PSObject.Properties) {
        if ($property.Name -notin $allowedUpdateFields) { throw "$candidateId 试图修改不允许的字段：$($property.Name)" }
        Set-Cell $range $headers $property.Name $property.Value
      }
    }
    $resultText = Clean $result.result
    $marker = "【自动重做】commentSha256=$actualHash; time=$((Get-Date).ToString('s')); result=$outcome"
    if ($resultText) { $marker += "; note=$($resultText.Replace("`r",' ').Replace("`n",' '))" }
    $combinedNote = (@($note,$marker) | Where-Object { $_ }) -join "`n"
    Set-Cell $range $headers '审核备注' $combinedNote
    if ($outcome -eq 'resolved') { Set-Cell $range $headers '审核状态' '待复核' }
    $reworked++
  }

  $added = 0
  foreach ($candidate in $newCandidates) {
    $canonical = Clean $candidate.canonicalUrl; if (-not $canonical) { $canonical = Clean $candidate.url }
    $platform = Clean $candidate.platformId
    if (($canonical -and $seen.ContainsKey($canonical.ToLowerInvariant())) -or ($platform -and $seen.ContainsKey($platform.ToLowerInvariant()))) { continue }
    $candidateId = if ($candidate.candidateId) { Clean $candidate.candidateId } else { 'cand-' + (Get-Date -Format 'yyyyMMdd') + '-' + ([Guid]::NewGuid().ToString('N').Substring(0,8)) }
    if ($seen.ContainsKey($candidateId.ToLowerInvariant())) { continue }
    $row = $table.ListRows.Add().Range
    $values = @{ '候选 ID'=$candidateId; '审核状态'='待复核'; '发现日期'=(Get-Date); '中文标题'=$candidate.title; '原标题'=$candidate.originalTitle; '作者'=$candidate.author; '来源'=$candidate.source; '原文 URL'=$candidate.url; '规范 URL'=$canonical; '平台 ID'=$platform; '发布年份'=$candidate.year; '语言'=$candidate.language; '媒介'=$candidate.medium; '建议来源层级'=$candidate.tier; '建议主题'=(@($candidate.topics) -join '；'); '证据状态'=$candidate.evidenceStatus; '短读证据'=$candidate.shortEvidence; '去重结果'='已与主表、review、catalog 和终态记录去重'; '审核备注'=$candidate.reason; '评论'=$candidate.comment }
    foreach ($key in $values.Keys) { Set-Cell $row $headers $key $values[$key] }
    $seen[$candidateId.ToLowerInvariant()]=$true; if($canonical){$seen[$canonical.ToLowerInvariant()]=$true}; if($platform){$seen[$platform.ToLowerInvariant()]=$true}; $added++
  }
  $book.Save()
  Write-Output "已处理待修改 $reworked 条（跳过已处理评论 $skippedRework 条），新增待复核 $added 条；未修改主表或发布网页。"
} finally {
  if ($book) { try { $book.Close($true) } catch {} }; if ($excel) { try { $excel.Quit() } catch {} }
  if ($lock) { $lock.Dispose(); Remove-Item -LiteralPath $lockPath -ErrorAction SilentlyContinue }
}
