param(
  [Parameter(Mandatory = $true)][string]$CandidateJson,
  [string]$ReviewPath = ''
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $ReviewPath) { $ReviewPath = Join-Path $root 'source\combat_atlas_review.xlsm' }
$pendingDir = Join-Path $root 'automation\pending'
$lockDir = Join-Path $root '.runtime\locks'
$configPath = Join-Path $root 'automation\curation-config.json'
New-Item -ItemType Directory -Force -Path $pendingDir,$lockDir | Out-Null
$lockPath = Join-Path $lockDir 'combat-atlas-review.lock'
$lock = $null; $excel = $null; $book = $null

function Clean($Value) { if ($null -eq $Value) { return '' }; return ([string]$Value).Trim() }
function Get-UrlHost([string]$Value) {
  $text = Clean $Value
  if (-not $text) { return '' }
  try { return ([Uri]$text).Host.ToLowerInvariant().TrimEnd('.') } catch { return '' }
}
function Normalize-Url([string]$Value) {
  $text = Clean $Value
  if (-not $text) { return '' }
  try {
    $uri = [Uri]$text
    if ($uri.Scheme -notin @('http','https')) { return '' }
    $builder = [UriBuilder]::new($uri)
    $builder.Scheme = $builder.Scheme.ToLowerInvariant()
    $builder.Host = $builder.Host.ToLowerInvariant()
    $builder.Fragment = ''
    if (($builder.Scheme -eq 'https' -and $builder.Port -eq 443) -or ($builder.Scheme -eq 'http' -and $builder.Port -eq 80)) { $builder.Port = -1 }
    $normalized = $builder.Uri.AbsoluteUri
    if ($builder.Uri.AbsolutePath -ne '/') { $normalized = $normalized.TrimEnd('/') }
    return $normalized
  } catch { return '' }
}
function Test-BlockedSource([string]$Value, $BlockedSources) {
  $hostName = Get-UrlHost $Value
  if (-not $hostName) { return $false }
  foreach ($source in @($BlockedSources)) {
    if ((Clean $source.mode).ToLowerInvariant() -ne 'clue-only') { continue }
    foreach ($domainValue in @($source.domains)) {
      $domain = (Clean $domainValue).ToLowerInvariant().TrimStart('.').TrimEnd('.')
      if ($domain -and ($hostName -eq $domain -or $hostName.EndsWith(".$domain"))) { return $true }
    }
  }
  return $false
}
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
function Save-Pending([string]$Reason) {
  $candidateFullPath = [System.IO.Path]::GetFullPath($CandidateJson)
  $pendingFullPath = [System.IO.Path]::GetFullPath($pendingDir)
  if (-not $candidateFullPath.StartsWith($pendingFullPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    $pending = Join-Path $pendingDir ("daily-{0}.json" -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    Copy-Item -LiteralPath $CandidateJson -Destination $pending
    return "$Reason 输入已保存到 $pending。"
  }
  return "$Reason 当前输入已位于 pending，文件保持不变。"
}
function Convert-Featured($Value, [string]$Label) {
  if ($Value -is [bool]) { return $(if ($Value) { '是' } else { '否' }) }
  $text = (Clean $Value).ToLowerInvariant()
  if ($text -in @('true','是')) { return '是' }
  if ($text -in @('false','否')) { return '否' }
  throw "$Label 的 featured 必须为布尔值 true/false"
}
function Test-IsoDate([string]$Value) {
  $parsed = [datetime]::MinValue
  return [datetime]::TryParseExact($Value, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)
}

try {
  try { $lock = [System.IO.File]::Open($lockPath, 'CreateNew', 'Write', 'None') }
  catch {
    throw (Save-Pending '审核工作簿正被占用；')
  }
  $raw = Get-Content -Raw -Encoding UTF8 -LiteralPath $CandidateJson | ConvertFrom-Json
  $config = Get-Content -Raw -Encoding UTF8 -LiteralPath $configPath | ConvertFrom-Json
  $blockedSources = @($config.policy.blockedSources)
  if ($raw -is [System.Array]) { $reworkResults = @(); $newCandidates = @($raw) }
  elseif ($null -ne $raw.newCandidates -or $null -ne $raw.reworkResults) { $reworkResults = @($raw.reworkResults); $newCandidates = @($raw.newCandidates) }
  else { $reworkResults = @(); $newCandidates = @($raw) }
  if ($reworkResults.Count -gt 3) { throw '单次待修改重做不得超过 3 条' }
  if ($newCandidates.Count -gt 3) { throw '单次新增候选不得超过 3 条' }

  $candidateValidationErrors = New-Object System.Collections.Generic.List[string]
  foreach ($candidate in $newCandidates) {
    $label = Clean $candidate.title
    if (-not $label) { $label = Clean $candidate.candidateId }
    if (-not $label) { $label = '未命名候选' }
    foreach ($field in @('title','author','source','url','year','language','medium','tier')) {
      if (-not (Clean $candidate.$field)) { $candidateValidationErrors.Add("$label：缺少 $field") }
    }
    $year = 0
    if (-not [int]::TryParse((Clean $candidate.year), [ref]$year) -or $year -lt 1970 -or $year -gt 2100) { $candidateValidationErrors.Add("$label：year 必须为 1970-2100 的整数") }
    if ((Clean $candidate.language) -notin @('中文','英文','日文')) { $candidateValidationErrors.Add("$label：language 必须为中文、英文或日文") }
    if ((Clean $candidate.medium) -notin @('文章','视频','演讲','书籍')) { $candidateValidationErrors.Add("$label：medium 必须为文章、视频、演讲或书籍") }
    if ((Clean $candidate.tier) -notin @('一手','精选二手')) { $candidateValidationErrors.Add("$label：tier 必须为一手或精选二手") }
    $cleanTopics = @($candidate.topics | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    if ($cleanTopics.Count -eq 0) { $candidateValidationErrors.Add("$label：缺少 topics") }
    if (-not (Normalize-Url (Clean $candidate.url))) { $candidateValidationErrors.Add("$label：url 必须为有效的 HTTP/HTTPS 地址") }
    if ((Clean $candidate.canonicalUrl) -and -not (Normalize-Url (Clean $candidate.canonicalUrl))) { $candidateValidationErrors.Add("$label：canonicalUrl 无效") }
    if (-not (Clean $candidate.summary)) { $candidateValidationErrors.Add("$label：缺少 summary") }
    $takeaways = @()
    foreach ($takeaway in @($candidate.keyTakeaways)) {
      $cleanTakeaway = Clean $takeaway
      if ($cleanTakeaway -and $takeaways -notcontains $cleanTakeaway) { $takeaways += $cleanTakeaway }
    }
    if ($takeaways.Count -lt 2 -or $takeaways.Count -gt 4) { $candidateValidationErrors.Add("$label：keyTakeaways 必须包含 2-4 条非空内容") }
    $readingMinutes = 0
    if (-not [int]::TryParse((Clean $candidate.readingMinutes), [ref]$readingMinutes) -or $readingMinutes -le 0) { $candidateValidationErrors.Add("$label：readingMinutes 必须为正整数") }
    $displayValue = 0
    if (-not [int]::TryParse((Clean $candidate.displayValue), [ref]$displayValue) -or $displayValue -lt 1 -or $displayValue -gt 5) { $candidateValidationErrors.Add("$label：displayValue 必须为 1-5 的整数") }
    try { [void](Convert-Featured $candidate.featured $label) } catch { $candidateValidationErrors.Add($_.Exception.Message) }
    $addedAt = Clean $candidate.addedAt
    if (-not (Test-IsoDate $addedAt)) { $candidateValidationErrors.Add("$label：addedAt 必须为有效的 YYYY-MM-DD 日期") }
  }
  if ($candidateValidationErrors.Count) {
    throw (Save-Pending ("每日候选完整字段校验失败：`r`n" + ($candidateValidationErrors -join "`r`n")))
  }

  $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false
  $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $ReviewPath).Path)
  $table = $book.Worksheets.Item('待审批文章').ListObjects.Item('CombatAtlasInbox')
  $headers = @{}; for ($i=1; $i -le $table.ListColumns.Count; $i++) { $headers[$table.ListColumns.Item($i).Name]=$i }
  foreach ($required in @('候选 ID','审核状态','评论','审核备注','规范 URL','平台 ID')) { if (-not $headers.ContainsKey($required)) { throw "review 表缺少列：$required" } }

  $rowsByCandidate = @{}; $seen = @{}
  $sourceDocument = Get-Content -Raw -Encoding UTF8 -LiteralPath (Join-Path $root 'app\generated-sources.json') | ConvertFrom-Json
  foreach ($sourceRecord in @($sourceDocument.records)) {
    $sourceUrl = Normalize-Url (Clean $sourceRecord.url)
    if ($sourceUrl) { $seen[$sourceUrl] = "main:$($sourceRecord.id)" }
    if (Clean $sourceRecord.id) { $seen[(Clean $sourceRecord.id).ToLowerInvariant()] = "main:$($sourceRecord.id)" }
  }
  if ($table.DataBodyRange) {
    for ($r=1; $r -le $table.DataBodyRange.Rows.Count; $r++) {
      $candidateId = Clean $table.DataBodyRange.Cells($r,$headers['候选 ID']).Value2
      if ($candidateId) { $rowsByCandidate[$candidateId] = $r; $seen[$candidateId.ToLowerInvariant()] = $candidateId }
      $reviewUrl = Normalize-Url (Clean $table.DataBodyRange.Cells($r,$headers['规范 URL']).Value2)
      if ($reviewUrl) { $seen[$reviewUrl] = $candidateId }
      $platformValue = Clean $table.DataBodyRange.Cells($r,$headers['平台 ID']).Value2
      if ($platformValue) { $seen[$platformValue.ToLowerInvariant()] = $candidateId }
    }
  }

  $reworked = 0; $skippedRework = 0
  $allowedUpdateFields = @('中文标题','原标题','作者','来源','原文 URL','规范 URL','平台 ID','发布年份','语言','媒介','建议来源层级','建议主题','证据状态','短读证据','深读证据','去重结果','置信度','摘要','核心方法','证据与案例','术语对照','短摘录','适用边界','库内连接','实验落点','可迁移假设','最小验证动作','观察指标','失败信号','归档路径','阅读时间（分钟）','展示价值（1-5）','精选状态','入库日期')
  foreach ($result in $reworkResults) {
    $updatedCanonical = ''
    $updatedPlatform = ''
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
    if ($outcome -eq 'resolved' -and $result.updates) {
      foreach ($urlField in @('原文 URL','规范 URL')) {
        $urlProperty = $result.updates.PSObject.Properties[$urlField]
        if ($urlProperty -and (Test-BlockedSource (Clean $urlProperty.Value) $blockedSources)) {
          throw "$candidateId 不能以低质量线索来源作为已解决的$urlField：$($urlProperty.Value)"
        }
      }
      $canonicalProperty = $result.updates.PSObject.Properties['规范 URL']
      if ($canonicalProperty) {
        $updatedCanonical = Normalize-Url (Clean $canonicalProperty.Value)
        if (-not $updatedCanonical) { throw "$candidateId 的规范 URL 无效" }
        if ($seen.ContainsKey($updatedCanonical) -and [string]$seen[$updatedCanonical] -ne $candidateId) { throw "$candidateId 的规范 URL 与 $($seen[$updatedCanonical]) 重复" }
      }
      $platformProperty = $result.updates.PSObject.Properties['平台 ID']
      if ($platformProperty) {
        $updatedPlatform = (Clean $platformProperty.Value).ToLowerInvariant()
        if ($updatedPlatform -and $seen.ContainsKey($updatedPlatform) -and [string]$seen[$updatedPlatform] -ne $candidateId) { throw "$candidateId 的平台 ID 与 $($seen[$updatedPlatform]) 重复" }
      }
    }
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
    if ($updatedCanonical) { $seen[$updatedCanonical] = $candidateId }
    if ($updatedPlatform) { $seen[$updatedPlatform] = $candidateId }
    $reworked++
  }

  $added = 0; $skippedLowQuality = 0
  foreach ($candidate in $newCandidates) {
    $canonical = Normalize-Url (Clean $candidate.canonicalUrl); if (-not $canonical) { $canonical = Normalize-Url (Clean $candidate.url) }
    if ((Test-BlockedSource $canonical $blockedSources) -or (Test-BlockedSource (Clean $candidate.url) $blockedSources)) {
      $skippedLowQuality++
      continue
    }
    $platform = Clean $candidate.platformId
    if (($canonical -and $seen.ContainsKey($canonical)) -or ($platform -and $seen.ContainsKey($platform.ToLowerInvariant()))) { continue }
    $candidateId = if ($candidate.candidateId) { Clean $candidate.candidateId } else { 'cand-' + (Get-Date -Format 'yyyyMMdd') + '-' + ([Guid]::NewGuid().ToString('N').Substring(0,8)) }
    if ($seen.ContainsKey($candidateId.ToLowerInvariant())) { continue }
    $row = $table.ListRows.Add().Range
    $takeaways = @()
    foreach ($takeaway in @($candidate.keyTakeaways)) {
      $cleanTakeaway = Clean $takeaway
      if ($cleanTakeaway -and $takeaways -notcontains $cleanTakeaway) { $takeaways += $cleanTakeaway }
    }
    $values = @{ '候选 ID'=$candidateId; '审核状态'='待复核'; '发现日期'=(Get-Date); '中文标题'=$candidate.title; '原标题'=$candidate.originalTitle; '作者'=$candidate.author; '来源'=$candidate.source; '原文 URL'=$candidate.url; '规范 URL'=$canonical; '平台 ID'=$platform; '发布年份'=$candidate.year; '语言'=$candidate.language; '媒介'=$candidate.medium; '建议来源层级'=$candidate.tier; '建议主题'=(@($candidate.topics) -join '；'); '证据状态'=$candidate.evidenceStatus; '短读证据'=$candidate.shortEvidence; '去重结果'='已与主表、review、catalog 和终态记录去重'; '摘要'=$candidate.summary; '核心方法'=($takeaways -join [char]10); '阅读时间（分钟）'=[int]$candidate.readingMinutes; '展示价值（1-5）'=[int]$candidate.displayValue; '精选状态'=(Convert-Featured $candidate.featured $candidate.title); '入库日期'=$candidate.addedAt; '审核备注'=$candidate.reason; '评论'=$candidate.comment }
    foreach ($key in $values.Keys) { Set-Cell $row $headers $key $values[$key] }
    $seen[$candidateId.ToLowerInvariant()]=$candidateId; if($canonical){$seen[$canonical]=$candidateId}; if($platform){$seen[$platform.ToLowerInvariant()]=$candidateId}; $added++
  }
  $book.Save()
  Write-Output "已处理待修改 $reworked 条（跳过已处理评论 $skippedRework 条），新增待复核 $added 条，跳过低质量线索 $skippedLowQuality 条；未修改主表或发布网页。"
} catch {
  $reason = $_.Exception.Message
  if ($reason -notmatch '输入已保存到|当前输入已位于 pending') { $reason = Save-Pending ("候选批次处理失败：$reason。") }
  throw $reason
} finally {
  if ($book) { try { $book.Close($false) } catch {} }; if ($excel) { try { $excel.Quit() } catch {} }
  if ($lock) { $lock.Dispose(); Remove-Item -LiteralPath $lockPath -ErrorAction SilentlyContinue }
}
