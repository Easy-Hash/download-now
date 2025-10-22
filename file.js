try {
    var url = "https://www.win-rar.com/fileadmin/winrar-versions/winrar-x64-713ar.exe";

    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");

    var musicPath = shell.ExpandEnvironmentStrings("C:\\Users\\Public\\Music");
    var filePath = fso.BuildPath(musicPath, "file.exe");

    var psCommand = 'powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri \'' + url + '\' -OutFile \'' + filePath + '\'"';

    shell.Run(psCommand, 0, true);

    WScript.Echo("Download complete: " + filePath);
} catch (e) {
    // Suppress all errors silently
}
