param(
  [string]$MainPath = '',
  [string]$ReviewPath = '',
  [switch]$Reopen
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $MainPath) { $MainPath = Join-Path $root 'source\combat_atlas_main.xlsm' }
if (-not $ReviewPath) { $ReviewPath = Join-Path $root 'source\combat_atlas_review.xlsm' }
$backupDir = Join-Path $root 'source\backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class CombatAtlasRunningObjects {
  [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable table);
  [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved, out IBindCtx context);
  public static object[] GetObjects() {
    var found = new List<object>(); IRunningObjectTable table; IBindCtx context;
    if (GetRunningObjectTable(0, out table) != 0 || CreateBindCtx(0, out context) != 0) return found.ToArray();
    IEnumMoniker iterator; table.EnumRunning(out iterator); iterator.Reset(); var monikers = new IMoniker[1];
    while (iterator.Next(1, monikers, IntPtr.Zero) == 0) { object value = null; try { table.GetObject(monikers[0], out value); } catch { continue; } if (value != null) found.Add(value); }
    return found.ToArray();
  }
}
'@

function Get-OpenWorkbook([string]$Path) {
  $fullPath = [System.IO.Path]::GetFullPath($Path)
  foreach ($candidate in [CombatAtlasRunningObjects]::GetObjects()) {
    try { if ([System.IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $fullPath) { return $candidate } } catch { }
  }
  return $null
}

$specs = @(
  @{ Path = [System.IO.Path]::GetFullPath($MainPath); Module = 'CombatAtlasPublisher'; File = (Join-Path $PSScriptRoot 'CombatAtlasPublisher.bas'); Table = 'CombatAtlasSources' },
  @{ Path = [System.IO.Path]::GetFullPath($ReviewPath); Module = 'CombatAtlasReviewer'; File = (Join-Path $PSScriptRoot 'CombatAtlasReviewer.bas'); Table = 'CombatAtlasInbox' }
)
$wasOpen = @{}
foreach ($spec in $specs) {
  if (-not (Test-Path -LiteralPath $spec.Path)) { throw "找不到工作簿：$($spec.Path)" }
  $open = Get-OpenWorkbook $spec.Path
  $wasOpen[$spec.Path] = $null -ne $open
  $backup = Join-Path $backupDir ("{0}-before-macro-{1}.xlsm" -f [System.IO.Path]::GetFileNameWithoutExtension($spec.Path),(Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
  if ($open) { $open.Save(); $open.SaveCopyAs($backup); $open.Close($true) } else { Copy-Item -LiteralPath $spec.Path -Destination $backup }
}

$securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $securityPath)) { New-Item -Path $securityPath -Force | Out-Null }
$old = Get-ItemProperty -Path $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue
$excel = $null; $book = $null
$results = @()
try {
  Set-ItemProperty -Path $securityPath -Name AccessVBOM -Type DWord -Value 1
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false; $excel.DisplayAlerts = $false
  foreach ($spec in $specs) {
    $book = $excel.Workbooks.Open($spec.Path)
    foreach ($component in @($book.VBProject.VBComponents)) {
      if ($component.Name -eq $spec.Module) { $book.VBProject.VBComponents.Remove($component); break }
    }
    $book.VBProject.VBComponents.Import((Resolve-Path -LiteralPath $spec.File).Path) | Out-Null
    $compile = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($compile -and $compile.Enabled) { $compile.Execute() }
    if ($compile -and $compile.Enabled) { throw "$($spec.Module) 未能通过 VBA 原生编译" }
    $table = $null; $sheet = $null
    foreach ($candidateSheet in @($book.Worksheets)) {
      foreach ($candidateTable in @($candidateSheet.ListObjects)) { if ($candidateTable.Name -eq $spec.Table) { $sheet = $candidateSheet; $table = $candidateTable; break } }
      if ($table) { break }
    }
    if (-not $table) { throw "找不到表：$($spec.Table)" }
    $formulaErrors = 0; try { $formulaErrors = $sheet.UsedRange.SpecialCells(-4123,16).Count } catch { }
    $rowCount = if ($table.DataBodyRange) { $table.DataBodyRange.Rows.Count } else { 0 }
    $columnCount = $table.ListColumns.Count
    $book.Save(); $book.Close($true); $book = $null
    $results += [ordered]@{ workbook=$spec.Path; module=$spec.Module; rows=$rowCount; columns=$columnCount; formulaErrors=$formulaErrors }
  }
}
finally {
  if ($book) { try { $book.Close($false) } catch { } }
  if ($excel) { try { $excel.Quit() } catch { } }
  if ($null -ne $old) { Set-ItemProperty -Path $securityPath -Name AccessVBOM -Type DWord -Value $old.AccessVBOM } else { Remove-ItemProperty -Path $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if ($Reopen) { foreach ($spec in $specs) { if ($wasOpen[$spec.Path]) { Start-Process -FilePath $spec.Path } } }
$results | ConvertTo-Json -Compress
