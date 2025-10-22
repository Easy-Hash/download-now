' VBScript to download a file using PowerShell to Public Music folder
On Error Resume Next

Dim url, filePath, shell, fso, musicPath, psCommand

url = "https://www.win-rar.com/fileadmin/winrar-versions/winrar-x64-713ar.exe"

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

musicPath = shell.ExpandEnvironmentStrings("C:\Users\Public\Music")
filePath = fso.BuildPath(musicPath, "file.exe")

psCommand = "powershell -NoProfile -ExecutionPolicy Bypass -Command ""Invoke-WebRequest -Uri '" & url & "' -OutFile '" & filePath & "'"""

shell.Run psCommand, 0, True

MsgBox "Download complete: " & filePath, 64, "Success"
