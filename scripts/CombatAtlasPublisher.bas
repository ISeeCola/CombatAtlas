Attribute VB_Name = "CombatAtlasPublisher"
Option Explicit

Public Sub PublishCombatAtlas()
    On Error GoTo Failed
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets(Utf16(&H77E5, &H8BC6, &H5E93, &H6587, &H7AE0))
    sheet.Range("E5").Value = Utf16(&H53D1, &H5E03, &H72B6, &H6001, &HFF1A, &H6B63, &H5728, &H4FDD, &H5B58, &H5E76, &H6821, &H9A8C, &H2026, &H2026)
    FillMissingSourceIds sheet.ListObjects("CombatAtlasSources")
    ThisWorkbook.Save

    Dim scriptPath As String
    scriptPath = ThisWorkbook.Path & "\..\scripts\publish-from-excel.ps1"
    If Dir(scriptPath) = "" Then Err.Raise vbObjectError + 100, , Utf16(&H627E, &H4E0D, &H5230, &H53D1, &H5E03, &H811A, &H672C, &HFF1A) & scriptPath

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
    sheet.Range("E5").Value = Utf16(&H53D1, &H5E03, &H72B6, &H6001, &HFF1A) & Left(Replace(output, vbCrLf, " "), 180)
    ThisWorkbook.Save
    MsgBox output, vbInformation, "CombatAtlas " & Utf16(&H53D1, &H5E03, &H5B8C, &H6210)
    Exit Sub

Failed:
    Dim failureMessage As String
    failureMessage = Err.Description
    On Error Resume Next
    If Not sheet Is Nothing Then sheet.Range("E5").Value = Utf16(&H53D1, &H5E03, &H72B6, &H6001, &HFF1A, &H5931, &H8D25, &H20, &H2014, &H20) & Left(failureMessage, 150)
    ThisWorkbook.Save
    On Error GoTo 0
    MsgBox failureMessage, vbCritical, "CombatAtlas " & Utf16(&H53D1, &H5E03, &H5931, &H8D25)
End Sub

Private Sub FillMissingSourceIds(ByVal table As ListObject)
    If table.DataBodyRange Is Nothing Then Exit Sub
    Randomize
    Dim idColumn As Long, statusColumn As Long, titleColumn As Long, rowIndex As Long
    idColumn = table.ListColumns("sourceId").Index
    statusColumn = table.ListColumns(Utf16(&H53D1, &H5E03, &H72B6, &H6001)).Index
    titleColumn = table.ListColumns(Utf16(&H4E2D, &H6587, &H6807, &H9898)).Index
    For rowIndex = 1 To table.DataBodyRange.Rows.Count
        If Trim(CStr(table.DataBodyRange.Cells(rowIndex, titleColumn).Value)) <> "" Then
            If Trim(CStr(table.DataBodyRange.Cells(rowIndex, idColumn).Value)) = "" Then
                table.DataBodyRange.Cells(rowIndex, idColumn).Value = "src-" & Format(Now, "yyyymmddhhnnss") & "-" & LCase(Hex(Int(Rnd() * 2147483647#)))
            End If
            If Trim(CStr(table.DataBodyRange.Cells(rowIndex, statusColumn).Value)) = "" Then
                table.DataBodyRange.Cells(rowIndex, statusColumn).Value = Utf16(&H53D1, &H5E03)
            End If
        End If
    Next rowIndex
End Sub

Private Function Utf16(ParamArray codepoints() As Variant) As String
    Dim index As Long
    For index = LBound(codepoints) To UBound(codepoints)
        Utf16 = Utf16 & ChrW$(CLng(codepoints(index)))
    Next index
End Function
