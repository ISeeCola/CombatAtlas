param(
  [string]$ReviewTemplate = (Join-Path (Split-Path $PSScriptRoot -Parent) 'outputs\review-build\combat_atlas_review.xlsx'),
  [string]$ReviewPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'source\combat_atlas_review.xlsm'),
  [string]$MainPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'source\combat_atlas_main.xlsm')
)
$ErrorActionPreference = 'Stop'
$securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $securityPath)) { New-Item -Path $securityPath -Force | Out-Null }
$old = Get-ItemProperty -Path $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue
$excel = $null
try {
  Set-ItemProperty -Path $securityPath -Name AccessVBOM -Type DWord -Value 1
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false; $excel.DisplayAlerts = $false

  $review = $excel.Workbooks.Open((Resolve-Path -LiteralPath $ReviewTemplate).Path)
  $review.VBProject.VBComponents.Import((Resolve-Path -LiteralPath (Join-Path $PSScriptRoot 'CombatAtlasReviewer.bas')).Path) | Out-Null
  $sheet = $review.Worksheets.Item('待审批文章')
  $target = $sheet.Range('A4:C5')
  $button = $sheet.Buttons().Add($target.Left, $target.Top, $target.Width, $target.Height)
  $button.Caption = '批准入库'; $button.OnAction = 'ApproveIntoMainWorkbook'; $button.Font.Name = 'Arial'; $button.Font.Size = 12; $button.Font.Bold = $true
  $review.SaveAs([System.IO.Path]::GetFullPath($ReviewPath), 52)
  $review.Close($true); $review = $null

  $main = $excel.Workbooks.Open((Resolve-Path -LiteralPath $MainPath).Path)
  foreach ($component in @($main.VBProject.VBComponents)) { if ($component.Name -eq 'CombatAtlasPublisher') { $main.VBProject.VBComponents.Remove($component); break } }
  $main.VBProject.VBComponents.Import((Resolve-Path -LiteralPath (Join-Path $PSScriptRoot 'CombatAtlasPublisher.bas')).Path) | Out-Null
  $main.Save(); $main.Close($true); $main = $null
} finally {
  if ($review) { try { $review.Close($false) } catch {} }
  if ($main) { try { $main.Close($false) } catch {} }
  if ($excel) { try { $excel.Quit() } catch {} }
  if ($null -ne $old) { Set-ItemProperty -Path $securityPath -Name AccessVBOM -Type DWord -Value $old.AccessVBOM }
  else { Remove-ItemProperty -Path $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Output "已更新本地双工作簿：$MainPath；$ReviewPath"
