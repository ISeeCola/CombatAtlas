Attribute VB_Name = "CombatAtlasReviewer"
Option Explicit

Public Sub ApproveIntoMainWorkbook()
    On Error GoTo Failed
    ThisWorkbook.Save
    Dim scriptPath As String
    scriptPath = ThisWorkbook.Path & "\..\scripts\promote-approved.ps1"
    If Dir(scriptPath) = "" Then Err.Raise vbObjectError + 200, , "Promotion script not found: " & scriptPath
    Dim shell As Object, process As Object, command As String
    Set shell = CreateObject("WScript.Shell")
    command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """ -ReviewPath """ & ThisWorkbook.FullName & """"
    Set process = shell.Exec(command)
    Do While process.Status = 0
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
    Loop
    Dim output As String
    output = Trim(process.StdOut.ReadAll & vbCrLf & process.StdErr.ReadAll)
    If process.ExitCode <> 0 Then Err.Raise vbObjectError + 201, , output
    MsgBox output, vbInformation, "CombatAtlas promotion completed"
    Exit Sub
Failed:
    MsgBox Err.Description, vbCritical, "CombatAtlas promotion failed"
End Sub
