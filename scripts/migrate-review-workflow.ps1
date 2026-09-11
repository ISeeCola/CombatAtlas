param(
  [string]$ReviewPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'source\combat_atlas_review.xlsm')
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$backupDir = Join-Path $root 'source\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backupPath = Join-Path $backupDir ("combat_atlas_review-before-status-v2-$stamp.xlsm")
Copy-Item -LiteralPath $ReviewPath -Destination $backupPath

$securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $securityPath)) { New-Item -Path $securityPath -Force | Out-Null }
$old = Get-ItemProperty -Path $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue
$excel = $null; $book = $null
try {
  Set-ItemProperty -Path $securityPath -Name AccessVBOM -Type DWord -Value 1
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false; $excel.DisplayAlerts = $false
  $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $ReviewPath).Path)
  $sheet = $book.Worksheets.Item('待审批文章')
  $table = $sheet.ListObjects.Item('CombatAtlasInbox')
  $headers = @{}; for ($i=1; $i -le $table.ListColumns.Count; $i++) { $headers[$table.ListColumns.Item($i).Name]=$i }
  foreach ($required in @('审核状态','候选 ID','审核备注','评论')) { if (-not $headers.ContainsKey($required)) { throw "review 表缺少列：$required" } }

  $statusCol = $headers['审核状态']
  $alreadyV2 = $false
  for ($r=1; $r -le $table.DataBodyRange.Rows.Count; $r++) {
    if (([string]$table.DataBodyRange.Cells($r,$statusCol).Value2).Trim() -in @('待复核','待修改')) { $alreadyV2 = $true; break }
  }
  for ($r=1; $r -le $table.DataBodyRange.Rows.Count; $r++) {
    $cell = $table.DataBodyRange.Cells($r,$statusCol)
    $status = ([string]$cell.Value2).Trim()
    switch ($status) {
      { $_ -in @('已发现','短名单','候选','待人工复核') } { $cell.Value2 = '待复核'; break }
      '已拒绝' { if (-not $alreadyV2) { $cell.Value2 = '待修改' }; break }
    }
  }

  $headerRow = $table.HeaderRowRange.Row
  $sheetColumn = $table.Range.Column + $statusCol - 1
  $validationRange = $sheet.Range($sheet.Cells($headerRow + 1, $sheetColumn), $sheet.Cells(1000, $sheetColumn))
  $validationRange.Validation.Delete()
  $validationRange.Validation.Add(3, 1, 1, '待复核,待修改,已接受,已入库,已拒绝,重复')
  $validationRange.Validation.IgnoreBlank = $true
  $validationRange.Validation.InCellDropdown = $true
  $validationRange.Validation.ErrorTitle = '状态无效'
  $validationRange.Validation.ErrorMessage = '请选择工作流定义的六种审核状态。'
  $validationRange.Validation.ShowError = $true

  $nextColumn = $sheetColumn + 1
  try {
    $nextRange = $sheet.Range($sheet.Cells($headerRow + 1, $nextColumn), $sheet.Cells(1000, $nextColumn))
    $formula = [string]$nextRange.Cells(1,1).Validation.Formula1
    if ($formula -match '已发现|短名单|待人工复核') { $nextRange.Validation.Delete() }
  } catch {}

  foreach ($component in @($book.VBProject.VBComponents)) {
    if ($component.Name -eq 'CombatAtlasReviewer') { $book.VBProject.VBComponents.Remove($component); break }
  }
  $book.VBProject.VBComponents.Import((Resolve-Path -LiteralPath (Join-Path $PSScriptRoot 'CombatAtlasReviewer.bas')).Path) | Out-Null
  $book.Save(); $book.Close($true); $book = $null
} finally {
  if ($book) { try { $book.Close($false) } catch {} }
  if ($excel) { try { $excel.Quit() } catch {} }
  if ($null -ne $old) { Set-ItemProperty -Path $securityPath -Name AccessVBOM -Type DWord -Value $old.AccessVBOM }
  else { Remove-ItemProperty -Path $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Output "review 工作流已迁移；备份：$backupPath"
