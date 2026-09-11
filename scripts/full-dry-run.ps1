$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent; Set-Location $root
& node scripts/sources-workbook.mjs --verify-generated; if ($LASTEXITCODE) { throw '公开数据校验失败' }
& node scripts/review-workbook.mjs --validate; if ($LASTEXITCODE) { throw 'review 校验失败' }
& node scripts/sync-curation-docs.mjs; if ($LASTEXITCODE) { throw '文档同步失败' }
& npm run lint; if ($LASTEXITCODE) { throw 'lint 失败' }
& npx tsc --noEmit --incremental false; if ($LASTEXITCODE) { throw '类型检查失败' }
& npm run build; if ($LASTEXITCODE) { throw '静态构建失败' }
Write-Output 'CombatAtlas 完整 Dry Run 通过；未提交或发布。'
