param(
  [string]$WorkbookPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'source\combat_atlas_main.xlsm'),
  [string]$ModulePath = (Join-Path $PSScriptRoot 'CombatAtlasPublisher.bas'),
  [switch]$Reopen
)

$ErrorActionPreference = 'Stop'
$WorkbookPath = [System.IO.Path]::GetFullPath($WorkbookPath)
$ModulePath = (Resolve-Path -LiteralPath $ModulePath).Path

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class CombatAtlasRot {
  [DllImport("ole32.dll")]
  private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable table);
  [DllImport("ole32.dll")]
  private static extern int CreateBindCtx(int reserved, out IBindCtx context);

  public static object[] GetObjects() {
    var objects = new List<object>();
    IRunningObjectTable table;
    IBindCtx context;
    if (GetRunningObjectTable(0, out table) != 0 || CreateBindCtx(0, out context) != 0) return objects.ToArray();
    IEnumMoniker enumerator;
    table.EnumRunning(out enumerator);
    enumerator.Reset();
    var monikers = new IMoniker[1];
    while (enumerator.Next(1, monikers, IntPtr.Zero) == 0) {
      object value = null;
      try { table.GetObject(monikers[0], out value); }
      catch { continue; }
      if (value != null) objects.Add(value);
    }
    return objects.ToArray();
  }
}
'@

$backupDir = Join-Path (Split-Path $WorkbookPath -Parent) 'backups'
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
$backupPath = Join-Path $backupDir ("combat_atlas_main-before-publisher-{0}.xlsm" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

$openWorkbook = $null
foreach ($candidate in [CombatAtlasRot]::GetObjects()) {
  try {
    if ([System.IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $WorkbookPath) {
      $openWorkbook = $candidate
      break
    }
  }
  catch { }
}

if ($openWorkbook) {
  $openWorkbook.Save()
  $openWorkbook.SaveCopyAs($backupPath)
  $openWorkbook.Close($true)
}
else {
  Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath
}

$securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $securityPath)) { New-Item -Path $securityPath -Force | Out-Null }
$old = Get-ItemProperty -Path $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue
$excel = $null
$workbook = $null
try {
  Set-ItemProperty -Path $securityPath -Name AccessVBOM -Type DWord -Value 1
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($WorkbookPath)
  foreach ($component in @($workbook.VBProject.VBComponents)) {
    if ($component.Name -eq 'CombatAtlasPublisher') {
      $workbook.VBProject.VBComponents.Remove($component)
      break
    }
  }
  $workbook.VBProject.VBComponents.Import($ModulePath) | Out-Null

  $compile = $excel.VBE.CommandBars.FindControl(1, 578)
  if ($compile -and $compile.Enabled) { $compile.Execute() }
  if ($compile -and $compile.Enabled) { throw 'VBA native compile did not complete.' }

  $sheet = $null
  $table = $null
  foreach ($candidateSheet in @($workbook.Worksheets)) {
    foreach ($candidateTable in @($candidateSheet.ListObjects)) {
      if ($candidateTable.Name -eq 'CombatAtlasSources') {
        $sheet = $candidateSheet
        $table = $candidateTable
        break
      }
    }
    if ($table) { break }
  }
  if (-not $table) { throw 'CombatAtlasSources table was not found.' }
  $formulaErrors = 0
  try { $formulaErrors = $sheet.UsedRange.SpecialCells(-4123, 16).Count } catch { }
  $result = [ordered]@{
    rows = if ($table.DataBodyRange) { $table.DataBodyRange.Rows.Count } else { 0 }
    columns = $table.ListColumns.Count
    buttons = $sheet.Buttons().Count
    formulaErrors = $formulaErrors
    backup = $backupPath
  }
  $workbook.Save()
  $workbook.Close($true)
  $workbook = $null
}
finally {
  if ($workbook) { try { $workbook.Close($false) } catch { } }
  if ($excel) { try { $excel.Quit() } catch { } }
  if ($null -ne $old) { Set-ItemProperty -Path $securityPath -Name AccessVBOM -Type DWord -Value $old.AccessVBOM }
  else { Remove-ItemProperty -Path $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

if ($Reopen) { Start-Process -FilePath $WorkbookPath }
$result | ConvertTo-Json -Compress
