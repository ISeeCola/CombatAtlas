param(
  [Parameter(Mandatory = $true)][string]$RemoteSha,
  [Parameter(Mandatory = $true)][string]$Message,
  [switch]$PurgedRoot
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Set-Location $root

function Find-CredentialManager {
  $roots = @((Join-Path $env:LOCALAPPDATA 'GitHubDesktop'), 'D:\_Applications\Driver\GitHubDesktop')
  foreach ($desktopRoot in $roots) {
    if (-not (Test-Path -LiteralPath $desktopRoot)) { continue }
    $versions = Get-ChildItem -LiteralPath $desktopRoot -Directory -Filter 'app-*' | Sort-Object Name -Descending
    foreach ($version in $versions) {
      $candidate = Join-Path $version.FullName 'resources\app\git\mingw64\bin\git-credential-manager.exe'
      if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
  }
  throw '找不到 GitHub Desktop Credential Manager'
}

$manager = Find-CredentialManager
$credential = "protocol=https`nhost=github.com`n`n" | & $manager get
$token = ($credential | Where-Object { $_ -like 'password=*' }) -replace '^password=', ''
if (-not $token) { throw 'GitHub Desktop 凭据不可用' }
$headers = @{ Authorization = "Bearer $token"; Accept = 'application/vnd.github+json'; 'X-GitHub-Api-Version' = '2022-11-28'; 'User-Agent' = 'CombatAtlas-Publisher' }

function Invoke-GitHubApi([string]$Method, [string]$Endpoint, $Body = $null) {
  $parameters = @{ Method = $Method; Uri = "https://api.github.com/repos/ISeeCola/CombatAtlas$Endpoint"; Headers = $headers }
  if ($null -ne $Body) {
    $parameters.ContentType = 'application/json'
    $parameters.Body = $Body | ConvertTo-Json -Depth 8 -Compress
  }
  return Invoke-RestMethod @parameters
}

function Invoke-GitChecked([string[]]$Arguments) {
  $result = & git @Arguments 2>&1
  if ($LASTEXITCODE -ne 0) { throw ($result -join [Environment]::NewLine) }
  return ($result -join [Environment]::NewLine).Trim()
}

function Read-GitBlobBytes([string]$Sha) {
  $startInfo = New-Object System.Diagnostics.ProcessStartInfo
  $startInfo.FileName = 'git'
  $startInfo.Arguments = "cat-file blob $Sha"
  $startInfo.UseShellExecute = $false
  $startInfo.RedirectStandardOutput = $true
  $startInfo.RedirectStandardError = $true
  $process = New-Object System.Diagnostics.Process
  $process.StartInfo = $startInfo
  if (-not $process.Start()) { throw "无法读取 Git blob：$Sha" }
  $memory = New-Object System.IO.MemoryStream
  $process.StandardOutput.BaseStream.CopyTo($memory)
  $errorText = $process.StandardError.ReadToEnd()
  $process.WaitForExit()
  if ($process.ExitCode -ne 0) { throw $errorText }
  return $memory.ToArray()
}

function Get-GitObjectSha([string]$Type, [byte[]]$Content) {
  $encoding = New-Object System.Text.UTF8Encoding($false)
  $header = $encoding.GetBytes("$Type $($Content.Length)`0")
  $framed = New-Object byte[] ($header.Length + $Content.Length)
  [Array]::Copy($header, 0, $framed, 0, $header.Length)
  [Array]::Copy($Content, 0, $framed, $header.Length, $Content.Length)
  $sha1 = [System.Security.Cryptography.SHA1]::Create()
  try { return (($sha1.ComputeHash($framed) | ForEach-Object { $_.ToString('x2') }) -join '') }
  finally { $sha1.Dispose() }
}

function Write-GitHubCommitObject($Commit) {
  $encoding = New-Object System.Text.UTF8Encoding($false)
  $authorEpoch = [DateTimeOffset]::Parse($Commit.author.date).ToUnixTimeSeconds()
  $committerEpoch = [DateTimeOffset]::Parse($Commit.committer.date).ToUnixTimeSeconds()
  $parentLines = @($Commit.parents | ForEach-Object { "parent $($_.sha)" })
  $offsets = @('+0000', '+0800')
  foreach ($sign in @('+', '-')) {
    foreach ($hour in 0..14) {
      foreach ($minute in @(0, 15, 30, 45)) {
        $candidate = '{0}{1:00}{2:00}' -f $sign, $hour, $minute
        if ($candidate -notin $offsets) { $offsets += $candidate }
      }
    }
  }
  foreach ($offset in $offsets) {
    foreach ($tail in @('', "`n")) {
      $lines = @("tree $($Commit.tree.sha)") + $parentLines + @(
        "author $($Commit.author.name) <$($Commit.author.email)> $authorEpoch $offset",
        "committer $($Commit.committer.name) <$($Commit.committer.email)> $committerEpoch $offset",
        '',
        $Commit.message
      )
      $content = $encoding.GetBytes(($lines -join "`n") + $tail)
      if ((Get-GitObjectSha 'commit' $content) -ne $Commit.sha) { continue }
      $tempPath = Join-Path $env:TEMP ("CombatAtlas-{0}.commit" -f [Guid]::NewGuid().ToString('N'))
      try {
        [System.IO.File]::WriteAllBytes($tempPath, $content)
        $written = Invoke-GitChecked @('hash-object', '-t', 'commit', '-w', $tempPath)
        if ($written -ne $Commit.sha) { throw "本地提交对象校验失败：$written" }
        return
      }
      finally { Remove-Item -LiteralPath $tempPath -ErrorAction SilentlyContinue }
    }
  }
  throw "无法在本地复现 GitHub 提交对象：$($Commit.sha)"
}

$remoteCommit = Invoke-GitHubApi 'GET' "/git/commits/$RemoteSha"
$localHead = Invoke-GitChecked @('rev-parse', 'HEAD')
if ($localHead -eq $RemoteSha) { Write-Output $RemoteSha; exit 0 }
$treeEntries = @()
$changes = if ($PurgedRoot) { @(& git ls-files | ForEach-Object { "A`t$_" }) } else { @(& git diff --name-status "$RemoteSha..HEAD") }
foreach ($change in $changes) {
  if (-not $change) { continue }
  $parts = $change -split "`t"
  $status = $parts[0]
  $file = $parts[-1].Replace('\', '/')
  if ($status -eq 'D') {
    $treeEntries += @{ path = $file; mode = '100644'; type = 'blob'; sha = $null }
    continue
  }
  $blobSha = Invoke-GitChecked @('rev-parse', "HEAD:$file")
  $bytes = Read-GitBlobBytes $blobSha
  $blob = Invoke-GitHubApi 'POST' '/git/blobs' @{ content = [Convert]::ToBase64String($bytes); encoding = 'base64' }
  if ($blob.sha -ne $blobSha) { throw "GitHub blob 校验失败：$file" }
  $mode = ((Invoke-GitChecked @('ls-tree', 'HEAD', '--', $file)) -split '\s+')[0]
  if (-not $mode) { $mode = '100644' }
  $treeEntries += @{ path = $file; mode = $mode; type = 'blob'; sha = $blob.sha }
}

$treeBody = if ($PurgedRoot) { @{ tree = $treeEntries } } else { @{ base_tree = $remoteCommit.tree.sha; tree = $treeEntries } }
$tree = Invoke-GitHubApi 'POST' '/git/trees' $treeBody
$localTree = Invoke-GitChecked @('rev-parse', 'HEAD^{tree}')
if ($tree.sha -ne $localTree) { throw "GitHub tree 与本地提交不一致：remote=$($tree.sha) local=$localTree" }
$commitBody = if ($PurgedRoot) { @{ message = $Message; tree = $tree.sha } } else { @{ message = $Message; tree = $tree.sha; parents = @($RemoteSha) } }
$commit = Invoke-GitHubApi 'POST' '/git/commits' $commitBody
Write-GitHubCommitObject $commit
Invoke-GitHubApi 'PATCH' '/git/refs/heads/main' @{ sha = $commit.sha; force = [bool]$PurgedRoot } | Out-Null
Invoke-GitChecked @('update-ref', 'refs/heads/main', $commit.sha, $localHead) | Out-Null
Invoke-GitChecked @('update-ref', 'refs/remotes/origin/main', $commit.sha) | Out-Null
Write-Output $commit.sha

