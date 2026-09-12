$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent; Set-Location $root
$runtimeRoot = Join-Path $root '.runtime\dry-run'
New-Item -ItemType Directory -Force -Path $runtimeRoot | Out-Null
$env:TEMP = $runtimeRoot
$env:TMP = $runtimeRoot
$env:npm_config_cache = Join-Path $runtimeRoot 'npm-cache'
& node scripts/sources-workbook.mjs --verify-generated; if ($LASTEXITCODE) { throw '公开数据校验失败' }
& node scripts/review-workbook.mjs --validate; if ($LASTEXITCODE) { throw 'review 校验失败' }
& node scripts/sync-curation-docs.mjs --check; if ($LASTEXITCODE) { throw '文档一致性检查失败' }
& npm run lint; if ($LASTEXITCODE) { throw 'lint 失败' }
& npm run typecheck; if ($LASTEXITCODE) { throw '类型检查失败' }
& npm test; if ($LASTEXITCODE) { throw '测试失败' }
& npm run audit:high; if ($LASTEXITCODE) { throw '依赖安全检查失败' }
& npm run build; if ($LASTEXITCODE) { throw '静态构建失败' }
Write-Output 'CombatAtlas 完整 Dry Run 通过；未提交或发布。'
