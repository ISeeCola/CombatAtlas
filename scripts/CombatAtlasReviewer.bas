Attribute VB_Name = "CombatAtlasReviewer"
Option Explicit

Public Sub ApproveIntoMainWorkbook()
    On Error GoTo Failed
    ThisWorkbook.Save
    Dim scriptPath As String, resultPath As String
    scriptPath = ThisWorkbook.Path & "\..\scripts\promote-approved.ps1"
    If Dir(scriptPath) = "" Then Err.Raise vbObjectError + 200, , "Promotion script not found: " & scriptPath
    resultPath = Environ$("TEMP") & "\CombatAtlas-promotion-" & Format(Now, "yyyymmddhhnnss") & ".txt"
    Dim shell As Object, process As Object, command As String
    Set shell = CreateObject("WScript.Shell")
    command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """ -ReviewPath """ & ThisWorkbook.FullName & """ -ResultPath """ & resultPath & """"
    Set process = shell.Exec(command)
    Do While process.Status = 0
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
    Loop
    Dim output As String
    If Dir(resultPath) <> "" Then
        output = Trim(ReadUnicodeText(resultPath))
    Else
        output = Trim(process.StdOut.ReadAll & vbCrLf & process.StdErr.ReadAll)
    End If
    On Error Resume Next
    Kill resultPath
    On Error GoTo Failed
    If process.ExitCode <> 0 Then
        MsgBox output, vbExclamation, "CombatAtlas promotion blocked"
        Exit Sub
    End If
    MsgBox output, vbInformation, "CombatAtlas promotion completed"
    Exit Sub
Failed:
    MsgBox Err.Description, vbCritical, "CombatAtlas promotion failed"
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
