param(
  [string]$InputPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'outputs\combat-atlas-excel-source\CombatAtlas.xlsx'),
  [string]$OutputPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'source\combat_atlas_main.xlsm')
)

$ErrorActionPreference = 'Stop'
$InputPath = (Resolve-Path -LiteralPath $InputPath).Path
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$modulePath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot 'CombatAtlasPublisher.bas')).Path
$securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $securityPath)) { New-Item -Path $securityPath -Force | Out-Null }
$oldProperty = Get-ItemProperty -Path $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue
$hadOldProperty = $null -ne $oldProperty
$oldValue = if ($hadOldProperty) { $oldProperty.AccessVBOM } else { $null }

$excel = $null
$workbook = $null
try {
  Set-ItemProperty -Path $securityPath -Name AccessVBOM -Type DWord -Value 1
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($InputPath)
  $workbook.VBProject.VBComponents.Import($modulePath) | Out-Null
  $sheet = $workbook.Worksheets.Item('知识库文章')
  $target = $sheet.Range('A5:C6')
  $button = $sheet.Buttons().Add($target.Left, $target.Top, $target.Width, $target.Height)
  $button.Caption = '发布到网页'
  $button.OnAction = 'PublishCombatAtlas'
  $button.Font.Name = 'Arial'
  $button.Font.Size = 12
  $button.Font.Bold = $true
  $workbook.SaveAs($OutputPath, 52)
  $workbook.Close($true)
  $workbook = $null
} finally {
  if ($workbook) { $workbook.Close($false) }
  if ($excel) { $excel.Quit() }
  if ($hadOldProperty) { Set-ItemProperty -Path $securityPath -Name AccessVBOM -Type DWord -Value $oldValue }
  else { Remove-ItemProperty -Path $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

Write-Output $OutputPath


