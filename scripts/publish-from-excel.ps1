param(
  [string]$WorkbookPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'source\combat_atlas_main.xlsm'),
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$root = Split-Path $PSScriptRoot -Parent
Set-Location $root

function Invoke-Git([string[]]$Arguments) {
  $result = & git @Arguments 2>&1
  if ($LASTEXITCODE -ne 0) { throw ($result -join [Environment]::NewLine) }
  return ($result -join [Environment]::NewLine).Trim()
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

if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "找不到工作簿：$WorkbookPath" }
if ((Invoke-Git @('branch', '--show-current')) -ne 'main') { throw '只能从 main 分支发布' }

$allowed = @('app/generated-sources.json', 'app/source-manifest.json')
$dirty = @(& git status --porcelain=v1 | ForEach-Object { $_.Substring(3).Replace('\', '/') })
$unrelated = @($dirty | Where-Object { $_ -notin $allowed })
if ($unrelated.Count -gt 0) { throw "存在与发布无关的工作区修改，请先处理：$($unrelated -join '、')" }

$githubHeaders = Get-GitHubHeaders
$remote = Invoke-RestMethod -Uri 'https://api.github.com/repos/ISeeCola/CombatAtlas/commits/main' -Headers $githubHeaders
$localSha = Invoke-Git @('rev-parse', 'HEAD')
if ($remote.sha -ne $localSha) { throw "本地 main 与 GitHub 不一致，请先同步。local=$($localSha.Substring(0,7)) remote=$($remote.sha.Substring(0,7))" }

$backupDir = Join-Path $env:LOCALAPPDATA 'CombatAtlas\Backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$backupPath = Join-Path $backupDir ("CombatAtlas-{0}.xlsm" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath

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

& node scripts/review-workbook.mjs --validate
if ($LASTEXITCODE -ne 0) { throw '审核工作簿校验失败' }
& node scripts/sync-curation-docs.mjs
if ($LASTEXITCODE -ne 0) { throw '本地策展文档同步失败' }

& npm run lint
if ($LASTEXITCODE -ne 0) { throw 'Lint 未通过' }
& npx tsc --noEmit --incremental false
if ($LASTEXITCODE -ne 0) { throw '类型检查未通过' }
$buildStarted = Get-Date
$buildLog = Join-Path $env:TEMP ("CombatAtlas-build-{0}.log" -f [Guid]::NewGuid().ToString('N'))
try {
  $buildCommand = "npm run build > `"$buildLog`" 2>&1"
  $buildProcess = Start-Process -FilePath 'cmd.exe' -ArgumentList @('/d', '/s', '/c', $buildCommand) -Wait -PassThru -WindowStyle Hidden
  $buildExitCode = $buildProcess.ExitCode
  $buildText = [System.IO.File]::ReadAllText($buildLog)
}
finally { Remove-Item -LiteralPath $buildLog -ErrorAction SilentlyContinue }
if ($buildExitCode -ne 0) {
  $artifact = Get-Item 'dist/client/index.html' -ErrorAction SilentlyContinue
  $knownWindowsExit = $buildText.Contains('Build complete.') -and $buildText.Contains('UV_HANDLE_CLOSING') -and $artifact -and $artifact.LastWriteTime -ge $buildStarted
  if (-not $knownWindowsExit) { throw "正式构建失败：$buildText" }
}
if ($DryRun) { Write-Output "演练通过：共 $($report.total) 条，发布 $($report.published) 条；未提交或推送。"; exit 0 }

& git add -- app/generated-sources.json app/source-manifest.json
if ($LASTEXITCODE -ne 0) { throw '无法暂存发布文件' }
& git diff --cached --quiet
if ($LASTEXITCODE -eq 0) { Write-Output '没有需要发布的变化。'; exit 0 }
$message = "Update CombatAtlas sources $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
Invoke-Git @('commit', '-m', $message) | Out-Null
$publishedSha = Invoke-Git @('rev-parse', 'HEAD')

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

$deadline = (Get-Date).AddMinutes(8)
$run = $null
do {
  Start-Sleep -Seconds 10
  $runs = Invoke-RestMethod -Uri "https://api.github.com/repos/ISeeCola/CombatAtlas/actions/runs?head_sha=$publishedSha&per_page=3" -Headers $githubHeaders
  $run = $runs.workflow_runs | Select-Object -First 1
} while ((-not $run -or $run.status -ne 'completed') -and (Get-Date) -lt $deadline)
if (-not $run -or $run.status -ne 'completed') { throw "代码已推送，但等待部署超时：$($publishedSha.Substring(0,7))" }
if ($run.conclusion -ne 'success') { throw "代码已推送，但 GitHub Pages 部署失败：$($run.html_url)" }
$page = Invoke-WebRequest -Uri 'https://iseecola.github.io/CombatAtlas/' -UseBasicParsing
if ($page.StatusCode -ne 200) { throw 'GitHub Pages 未返回成功状态' }
$auditDir = Join-Path $root 'automation\private-audit'
New-Item -ItemType Directory -Force -Path $auditDir | Out-Null
$workbookHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $WorkbookPath).Hash.ToLowerInvariant()
$audit = [ordered]@{ generatedAt = (Get-Date).ToString('o'); workbookSha256 = $workbookHash; dataSha256 = $report.dataSha256; commit = $publishedSha; deployment = $run.html_url }
$audit | ConvertTo-Json | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $auditDir ("publish-{0}.json" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
Write-Output "发布成功：$($publishedSha.Substring(0,7)) https://iseecola.github.io/CombatAtlas/"


