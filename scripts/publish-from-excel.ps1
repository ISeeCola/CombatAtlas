param(
  [string]$WorkbookPath = '',
  [string]$ProgressPath = '',
  [string]$ResultPath = '',
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$root = Split-Path $PSScriptRoot -Parent
if (-not $WorkbookPath) { $WorkbookPath = Join-Path $root 'source\combat_atlas_main.xlsm' }
Set-Location $root
$runtimeRoot = Join-Path $root '.runtime\publish'
$logRoot = Join-Path $runtimeRoot 'logs'
$tempRoot = Join-Path $runtimeRoot 'temp'
New-Item -ItemType Directory -Force -Path $logRoot,$tempRoot | Out-Null
$env:TEMP = $tempRoot
$env:TMP = $tempRoot
$env:npm_config_cache = Join-Path $runtimeRoot 'npm-cache'
$logPath = Join-Path $logRoot ("publish-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$script:deploymentComplete = $false

function Invoke-Git([string[]]$Arguments) {
  $result = & git @Arguments 2>&1
  if ($LASTEXITCODE -ne 0) { throw ($result -join [Environment]::NewLine) }
  return ($result -join [Environment]::NewLine).Trim()
}

function Set-PublishProgress([int]$Percent, [string]$Message) {
  if ([string]::IsNullOrWhiteSpace($ProgressPath)) { return }
  $progressDirectory = Split-Path -Parent $ProgressPath
  if ($progressDirectory) { New-Item -ItemType Directory -Force -Path $progressDirectory | Out-Null }
  [System.IO.File]::WriteAllText(
    $ProgressPath,
    ("发布进度 {0}%：{1}" -f $Percent, $Message),
    [System.Text.Encoding]::Unicode
  )
}

function Set-PublishResult([string]$Message) {
  if ([string]::IsNullOrWhiteSpace($ResultPath)) { return }
  $resultDirectory = Split-Path -Parent $ResultPath
  if ($resultDirectory) { New-Item -ItemType Directory -Force -Path $resultDirectory | Out-Null }
  [System.IO.File]::WriteAllText($ResultPath, $Message, [System.Text.Encoding]::Unicode)
}

function Get-SharedFileSha256([string]$Path) {
  $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
  $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
  try {
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
      return (($sha256.ComputeHash($stream) | ForEach-Object { $_.ToString('x2') }) -join '')
    }
    finally { $sha256.Dispose() }
  }
  finally { $stream.Dispose() }
}

function Copy-SharedFile([string]$Source, [string]$Destination) {
  $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
  $input = [System.IO.File]::Open($Source, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
  try {
    $output = [System.IO.File]::Open($Destination, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try { $input.CopyTo($output) } finally { $output.Dispose() }
  }
  finally { $input.Dispose() }
}

function Get-GitHubHeaders {
  $headers = @{ Accept = 'application/vnd.github+json'; 'User-Agent' = 'CombatAtlas-Publisher' }
  $roots = @((Join-Path $env:LOCALAPPDATA 'GitHubDesktop'), 'D:\_Applications\Driver\GitHubDesktop')
  foreach ($desktopRoot in $roots) {
    if (-not (Test-Path -LiteralPath $desktopRoot)) { continue }
    $versions = Get-ChildItem -LiteralPath $desktopRoot -Directory -Filter 'app-*' | Sort-Object Name -Descending
    foreach ($version in $versions) {
      $manager = Join-Path $version.FullName 'resources\app\git\mingw64\bin\git-credential-manager.exe'
      if (-not (Test-Path -LiteralPath $manager)) { continue }
      $credential = "protocol=https`nhost=github.com`n`n" | & $manager get
      $token = ($credential | Where-Object { $_ -like 'password=*' }) -replace '^password=', ''
      if ($token) { $headers.Authorization = "Bearer $token" }
      return $headers
    }
  }
  return $headers
}

function Invoke-Publish {

Set-PublishProgress 3 '初始化并检查工作簿'
if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "找不到工作簿：$WorkbookPath" }
if ((Invoke-Git @('branch', '--show-current')) -ne 'main') { throw '只能从 main 分支发布' }

Set-PublishProgress 8 '检查工作区状态'
$allowed = @('app/generated-sources.json', 'app/source-manifest.json')
$dirty = @(& git status --porcelain=v1 | ForEach-Object { $_.Substring(3).Replace('\', '/') })
$unrelated = @($dirty | Where-Object { $_ -notin $allowed })
if ($unrelated.Count -gt 0) { throw "存在与发布无关的工作区修改，请先处理：$($unrelated -join '、')" }

Set-PublishProgress 14 '核对 GitHub 版本'
$githubHeaders = Get-GitHubHeaders
$remote = Invoke-RestMethod -Uri 'https://api.github.com/repos/ISeeCola/CombatAtlas/commits/main' -Headers $githubHeaders
$localSha = Invoke-Git @('rev-parse', 'HEAD')
if ($remote.sha -ne $localSha) {
  & git merge-base --is-ancestor $remote.sha $localSha
  if ($LASTEXITCODE -ne 0) { throw "本地 main 与 GitHub 已分叉或落后，请先同步。local=$($localSha.Substring(0,7)) remote=$($remote.sha.Substring(0,7))" }
}

Set-PublishProgress 20 '备份本地主表'
$backupDir = Join-Path $root 'source\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$backupPath = Join-Path $backupDir ("CombatAtlas-{0}.xlsm" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Copy-SharedFile -Source $WorkbookPath -Destination $backupPath

Set-PublishProgress 28 '生成并校验网页数据'
$env:COMBAT_ATLAS_WORKBOOK = $WorkbookPath
$reportJson = (& node scripts/sources-workbook.mjs --sync | Select-Object -Last 1)
if ($LASTEXITCODE -ne 0) { throw '工作簿数据生成失败' }
$report = $reportJson | ConvertFrom-Json

if ($report.highRisk -and -not $DryRun) {
  Add-Type -AssemblyName System.Windows.Forms
  $message = "检测到高风险变更：新增 $($report.added)，修改 $($report.modified)，下架 $($report.downlisted)。是否继续发布？"
  $choice = [System.Windows.Forms.MessageBox]::Show($message, 'CombatAtlas 发布确认', 'YesNo', 'Warning')
  if ($choice -ne 'Yes') { throw '用户取消发布' }
}

Set-PublishProgress 38 '校验 Review 并同步本地文档'
& node scripts/review-workbook.mjs --validate
if ($LASTEXITCODE -ne 0) { throw '审核工作簿校验失败' }
& node scripts/sync-curation-docs.mjs
if ($LASTEXITCODE -ne 0) { throw '本地策展文档同步失败' }

Set-PublishProgress 48 '运行代码规范检查'
& npm run lint
if ($LASTEXITCODE -ne 0) { throw 'Lint 未通过' }
Set-PublishProgress 56 '运行 TypeScript 类型检查'
& npm run typecheck
if ($LASTEXITCODE -ne 0) { throw '类型检查未通过' }
Set-PublishProgress 61 '运行自动化测试'
& npm test
if ($LASTEXITCODE -ne 0) { throw '自动化测试未通过' }
Set-PublishProgress 65 '检查高危依赖'
& npm run audit:high
if ($LASTEXITCODE -ne 0) { throw '依赖安全检查未通过' }
Set-PublishProgress 69 '构建静态网页；此阶段可能需要数分钟'
$buildLog = Join-Path $tempRoot ("CombatAtlas-build-{0}.log" -f [Guid]::NewGuid().ToString('N'))
$buildCommand = "npm run build > `"$buildLog`" 2>&1"
$buildProcess = Start-Process -FilePath 'cmd.exe' -ArgumentList @('/d', '/s', '/c', $buildCommand) -Wait -PassThru -WindowStyle Hidden
$buildExitCode = $buildProcess.ExitCode
$buildText = [System.IO.File]::ReadAllText($buildLog)
if ($buildExitCode -ne 0) {
  throw "正式构建失败，完整日志：$buildLog`n$buildText"
}
if ($DryRun) {
  $resultMessage = "演练完成。`r`n共 $($report.total) 条，发布 $($report.published) 条。`r`n未提交或推送。"
  Set-PublishProgress 100 '演练完成，未提交或推送'
  Set-PublishResult $resultMessage
  Write-Output $resultMessage
  return
}

Set-PublishProgress 75 '准备 Git 提交'
& git add -- app/generated-sources.json app/source-manifest.json
if ($LASTEXITCODE -ne 0) { throw '无法暂存发布文件' }
& git diff --cached --quiet
$hasDataChanges = $LASTEXITCODE -ne 0
$message = "Update CombatAtlas sources $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
if ($hasDataChanges) {
  Invoke-Git @('commit', '-m', $message) | Out-Null
}
$publishedSha = Invoke-Git @('rev-parse', 'HEAD')
$aheadCount = [int](Invoke-Git @('rev-list', '--count', "$($remote.sha)..HEAD"))
if ($aheadCount -eq 0) {
  $resultMessage = '没有需要发布的变化。'
  Set-PublishProgress 100 '没有需要发布的变化'
  Set-PublishResult $resultMessage
  Write-Output $resultMessage
  return
}
if (-not $hasDataChanges) { $message = "Publish $aheadCount pending CombatAtlas commit(s)" }

Set-PublishProgress 82 '推送到 GitHub'
$previousErrorActionPreference = $ErrorActionPreference
try {
  $ErrorActionPreference = 'Continue'
  $pushOutput = & git push origin main 2>&1
  $pushExitCode = $LASTEXITCODE
  if ($pushExitCode -ne 0) {
    $fallback = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/publish-via-github-api.ps1 -RemoteSha $remote.sha -Message $message 2>&1
    $fallbackExitCode = $LASTEXITCODE
  }
}
finally { $ErrorActionPreference = $previousErrorActionPreference }
if ($pushExitCode -ne 0) {
  if ($fallbackExitCode -ne 0) { throw "Git 推送和 API 备用通道均失败：$($pushOutput -join ' ')；$($fallback -join ' ')" }
  $publishedSha = ($fallback | Select-Object -Last 1).Trim()
}

Set-PublishProgress 90 '等待 GitHub Pages 部署'
$deadline = (Get-Date).AddMinutes(8)
$run = $null
do {
  Start-Sleep -Seconds 10
  $runs = Invoke-RestMethod -Uri "https://api.github.com/repos/ISeeCola/CombatAtlas/actions/runs?head_sha=$publishedSha&per_page=3" -Headers $githubHeaders
  $run = $runs.workflow_runs | Select-Object -First 1
} while ((-not $run -or $run.status -ne 'completed') -and (Get-Date) -lt $deadline)
if (-not $run -or $run.status -ne 'completed') { throw "代码已推送，但等待部署超时：$($publishedSha.Substring(0,7))" }
if ($run.conclusion -ne 'success') { throw "代码已推送，但 GitHub Pages 部署失败：$($run.html_url)" }
Set-PublishProgress 97 '验证线上页面'
$page = Invoke-WebRequest -Uri 'https://iseecola.github.io/CombatAtlas/' -UseBasicParsing
if ($page.StatusCode -ne 200) { throw 'GitHub Pages 未返回成功状态' }
$script:deploymentComplete = $true
$auditDir = Join-Path $root 'automation\private-audit'
New-Item -ItemType Directory -Force -Path $auditDir | Out-Null
$workbookHash = Get-SharedFileSha256 -Path $WorkbookPath
$audit = [ordered]@{ generatedAt = (Get-Date).ToString('o'); workbookSha256 = $workbookHash; dataSha256 = $report.dataSha256; commit = $publishedSha; deployment = $run.html_url }
$audit | ConvertTo-Json | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $auditDir ("publish-{0}.json" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
Set-PublishProgress 100 '发布完成'
$resultMessage = "发布完成。`r`n提交：$($publishedSha.Substring(0,7))`r`n网址：https://iseecola.github.io/CombatAtlas/"
Set-PublishResult $resultMessage
Write-Output $resultMessage
}

try {
  Start-Transcript -LiteralPath $logPath -Force | Out-Null
  Invoke-Publish
}
catch {
  $prefix = if ($script:deploymentComplete) { '网页已经发布，但本地审计未完成。' } else { '发布尚未完成。' }
  $resultMessage = "$prefix`r`n$($_.Exception.Message)`r`n日志：$logPath"
  Set-PublishResult $resultMessage
  [Console]::Error.WriteLine($resultMessage)
  exit 1
}
finally {
  try { Stop-Transcript | Out-Null } catch { }
}


