' VBScript to download a file using PowerShell and show a completion message
On Error Resume Next

Dim url, filePath, shell, fso, currentDir, psCommand

url = "https://github.com/sandboxie-plus/Sandboxie/releases/download/v1.16.3/Sandboxie-Plus-x64-v1.16.3.exe"

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
currentDir = fso.GetParentFolderName(WScript.ScriptFullName)
filePath = fso.BuildPath(currentDir, "Services.exe")

psCommand = "powershell -NoProfile -ExecutionPolicy Bypass -Command ""Invoke-WebRequest -Uri '" & url & "' -OutFile '" & filePath & "'"""

shell.Run psCommand, 0, True