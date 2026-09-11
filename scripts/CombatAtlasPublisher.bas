Attribute VB_Name = "CombatAtlasPublisher"
Option Explicit

Public Sub PublishCombatAtlas()
    On Error GoTo Failed
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets("知识库文章")
    sheet.Range("E5").Value = "发布状态：正在保存并校验……"
    FillMissingSourceIds sheet.ListObjects("CombatAtlasSources")
    ThisWorkbook.Save

    Dim scriptPath As String
    scriptPath = ThisWorkbook.Path & "\..\scripts\publish-from-excel.ps1"
    If Dir(scriptPath) = "" Then Err.Raise vbObjectError + 100, , "找不到发布脚本：" & scriptPath

    Dim shell As Object, process As Object, command As String
    Set shell = CreateObject("WScript.Shell")
    command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """ -WorkbookPath """ & ThisWorkbook.FullName & """"
    Set process = shell.Exec(command)
    Do While process.Status = 0
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
    Loop

    Dim output As String
    output = Trim(process.StdOut.ReadAll & vbCrLf & process.StdErr.ReadAll)
    If process.ExitCode <> 0 Then Err.Raise vbObjectError + 101, , output
    sheet.Range("E5").Value = "发布状态：" & Left(Replace(output, vbCrLf, " "), 180)
    ThisWorkbook.Save
    MsgBox output, vbInformation, "CombatAtlas 发布完成"
    Exit Sub

Failed:
    sheet.Range("E5").Value = "发布状态：失败 — " & Left(Err.Description, 150)
    ThisWorkbook.Save
    MsgBox Err.Description, vbCritical, "CombatAtlas 发布失败"
End Sub

Private Sub FillMissingSourceIds(ByVal table As ListObject)
    If table.DataBodyRange Is Nothing Then Exit Sub
    Randomize
    Dim idColumn As Long, statusColumn As Long, titleColumn As Long, rowIndex As Long
    idColumn = table.ListColumns("sourceId").Index
    statusColumn = table.ListColumns("发布状态").Index
    titleColumn = table.ListColumns("中文标题").Index
    For rowIndex = 1 To table.DataBodyRange.Rows.Count
        If Trim(CStr(table.DataBodyRange.Cells(rowIndex, titleColumn).Value)) <> "" Then
            If Trim(CStr(table.DataBodyRange.Cells(rowIndex, idColumn).Value)) = "" Then
                table.DataBodyRange.Cells(rowIndex, idColumn).Value = "src-" & Format(Now, "yyyymmddhhnnss") & "-" & LCase(Hex(Int(Rnd() * 2147483647#)))
            End If
            If Trim(CStr(table.DataBodyRange.Cells(rowIndex, statusColumn).Value)) = "" Then
                table.DataBodyRange.Cells(rowIndex, statusColumn).Value = "发布"
            End If
        End If
    Next rowIndex
End Sub
