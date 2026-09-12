Attribute VB_Name = "CombatAtlasReviewer"
Option Explicit

Public Sub ApproveIntoMainWorkbook()
    On Error GoTo Failed
    ThisWorkbook.Save
    Dim scriptPath As String, resultPath As String, updatePath As String
    scriptPath = ThisWorkbook.Path & "\..\scripts\promote-approved.ps1"
    If Dir(scriptPath) = "" Then Err.Raise vbObjectError + 200, , "Promotion script not found: " & scriptPath
    resultPath = TempFile("promotion", ".txt")
    updatePath = TempFile("updates", ".tsv")

    Dim command As String, exitCode As Long, output As String
    command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """ -ReviewPath """ & ThisWorkbook.FullName & """ -ResultPath """ & resultPath & """ -ReviewUpdatePath """ & updatePath & """"
    exitCode = RunAndWait(command)
    output = ReadResult(resultPath)
    If exitCode <> 0 Then
        MsgBox output, vbExclamation, "CombatAtlas promotion blocked"
        GoTo Cleanup
    End If
    If Dir(updatePath) <> "" Then ApplyReviewUpdates updatePath
    ThisWorkbook.Save

    DeleteIfExists resultPath
    command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """ -SyncDocsOnly -ReviewPath """ & ThisWorkbook.FullName & """ -ResultPath """ & resultPath & """"
    exitCode = RunAndWait(command)
    output = ReadResult(resultPath)
    If exitCode <> 0 Then
        MsgBox output, vbExclamation, "CombatAtlas docs sync needs retry"
    Else
        MsgBox output, vbInformation, "CombatAtlas promotion completed"
    End If
Cleanup:
    DeleteIfExists resultPath
    DeleteIfExists updatePath
    Exit Sub
Failed:
    MsgBox Err.Description, vbCritical, "CombatAtlas promotion failed"
    Resume Cleanup
End Sub

Private Sub ApplyReviewUpdates(ByVal filePath As String)
    Dim lines() As String, fields() As String, meta() As String
    lines = Split(Replace(ReadUnicodeText(filePath), vbCrLf, vbLf), vbLf)
    If UBound(lines) < 0 Or Left$(lines(0), 8) <> "#columns" Then Err.Raise vbObjectError + 201, , "Invalid review update control file."
    meta = Split(lines(0), vbTab)
    If UBound(meta) < 4 Then Err.Raise vbObjectError + 202, , "Missing review column map."
    Dim candidateCol As Long, sourceCol As Long, statusCol As Long, noteCol As Long
    candidateCol = CLng(meta(1)): sourceCol = CLng(meta(2)): statusCol = CLng(meta(3)): noteCol = CLng(meta(4))
    Dim table As ListObject, i As Long, rowIndex As Long
    Set table = ThisWorkbook.Worksheets(1).ListObjects("CombatAtlasInbox")
    For i = 1 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            fields = Split(lines(i), vbTab)
            If UBound(fields) >= 3 Then
                rowIndex = FindDataRow(table, candidateCol, fields(0))
                If rowIndex = 0 Then Err.Raise vbObjectError + 203, , "Candidate not found while applying updates: " & fields(0)
                table.DataBodyRange.Cells(rowIndex, sourceCol).Value2 = fields(1)
                table.DataBodyRange.Cells(rowIndex, statusCol).Value2 = fields(2)
                table.DataBodyRange.Cells(rowIndex, noteCol).Value2 = Replace(fields(3), "\n", vbLf)
            End If
        End If
    Next i
End Sub

Private Function FindDataRow(ByVal table As ListObject, ByVal columnIndex As Long, ByVal candidateId As String) As Long
    Dim i As Long
    For i = 1 To table.DataBodyRange.Rows.Count
        If Trim$(CStr(table.DataBodyRange.Cells(i, columnIndex).Value2)) = candidateId Then FindDataRow = i: Exit Function
    Next i
    FindDataRow = 0
End Function

Private Function RunAndWait(ByVal command As String) As Long
    Dim shell As Object, process As Object
    Set shell = CreateObject("WScript.Shell")
    Set process = shell.Exec(command)
    Do While process.Status = 0
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
    Loop
    RunAndWait = process.ExitCode
End Function

Private Function ReadResult(ByVal filePath As String) As String
    If Dir(filePath) = "" Then ReadResult = "Operation failed without diagnostics.": Exit Function
    ReadResult = Trim$(ReadUnicodeText(filePath))
    If Len(ReadResult) = 0 Then ReadResult = "Operation failed without diagnostics."
End Function

Private Function TempFile(ByVal label As String, ByVal extension As String) As String
    Randomize
    TempFile = RuntimeFolder() & "\CombatAtlas-" & label & "-" & Format(Now, "yyyymmddhhnnss") & "-" & CStr(Int(Rnd() * 100000)) & extension
End Function

Private Function RuntimeFolder() As String
    Dim fileSystem As Object, rootPath As String, folderPath As String
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    rootPath = ThisWorkbook.Path & "\..\.runtime"
    folderPath = rootPath & "\ipc"
    If Not fileSystem.FolderExists(rootPath) Then fileSystem.CreateFolder rootPath
    If Not fileSystem.FolderExists(folderPath) Then fileSystem.CreateFolder folderPath
    RuntimeFolder = folderPath
End Function

Private Sub DeleteIfExists(ByVal filePath As String)
    On Error Resume Next
    If Len(filePath) > 0 And Dir(filePath) <> "" Then Kill filePath
    On Error GoTo 0
End Sub

Private Function ReadUnicodeText(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "unicode"
    stream.Open
    stream.LoadFromFile filePath
    ReadUnicodeText = stream.ReadText
    stream.Close
End Function
