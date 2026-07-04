Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

strDir = FSO.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = strDir

If Not FSO.FileExists(strDir & "\python\python.exe") Then
    MsgBox "python\python.exe not found. Please reinstall the application.", vbCritical, "SEO Toolkit Pro"
    WScript.Quit
End If

' Kill any running instance before starting a fresh one
WshShell.Run "taskkill /F /IM python.exe /FI ""WINDOWTITLE eq SEO Toolkit Pro""", 0, True
WshShell.Run "cmd /c for /L %i in (5070,1,5089) do (for /f ""tokens=5"" %P in ('netstat -ano ^| findstr ""127.0.0.1:%i"" ^| findstr LISTENING 2^>nul') do taskkill /F /PID %P >nul 2>&1)", 0, True
WScript.Sleep 800

' Use TEMP folder for port file (Program Files is read-only)
Dim portFile
portFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\.grc_port"
If FSO.FileExists(portFile) Then FSO.DeleteFile portFile

' Apply OTA updates BEFORE launching, so a single start picks up new code
' (the app then imports the freshly downloaded modules). Runs hidden, waits.
WshShell.Run "cmd /c """ & strDir & "\python\python.exe"" -s """ & strDir & "\updater.py""", 0, True

' Start Flask server silently, pass port file location via env var
WshShell.Run "cmd /c ""set GRC_NO_BROWSER=1&& set GRC_PORT_FILE=" & portFile & "&& """ & strDir & "\python\python.exe"" -s """ & strDir & "\web_app_batch.py""""", 0, False

' Wait for port file (server writes it when ready)
Dim port, appUrl, attempts
port = ""
attempts = 0
Do While attempts < 40
    WScript.Sleep 500
    attempts = attempts + 1
    If FSO.FileExists(portFile) Then
        Dim f
        Set f = FSO.OpenTextFile(portFile, 1)
        port = Trim(f.ReadAll())
        f.Close
        Exit Do
    End If
Loop

If port = "" Then
    MsgBox "SEO Toolkit Pro could not start. Please try launching it again.", vbExclamation, "SEO Toolkit Pro"
    WScript.Quit
End If

appUrl = "http://127.0.0.1:" & port

' Small wait for Flask to be fully ready
WScript.Sleep 1000

' Open in Edge app mode (clean window, no address bar)
Dim edgePath
edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
If Not FSO.FileExists(edgePath) Then
    edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
End If

If FSO.FileExists(edgePath) Then
    Dim appProfile
    appProfile = WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\SEO Toolkit Pro\edge_app"
    ' Clear Edge favicon cache so the current app icon is always picked up
    Dim iconCache
    iconCache = appProfile & "\Default\Platform Notifications"
    If FSO.FolderExists(iconCache) Then FSO.DeleteFolder iconCache, True
    WshShell.Run "cmd /c del /q """ & appProfile & "\Default\Favicons*"" 2>nul & del /q """ & appProfile & "\Default\Web Applications\Manifest Resources\*\Icons\*"" 2>nul", 0, True

    ' Launch the Edge app window through a shortcut that carries our own
    ' AppUserModelID + icon, so the TASKBAR shows the app icon (not Edge's).
    ' (--app-icon is a removed Chrome/Edge flag and does nothing.) Falls back
    ' to a direct Edge launch inside the helper if anything fails.
    Dim psHelper, psCmd
    psHelper = strDir & "\app_launch.ps1"
    If FSO.FileExists(psHelper) Then
        psCmd = "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & psHelper & """ " & _
                "-Edge """ & edgePath & """ " & _
                "-Url """ & appUrl & """ " & _
                "-Profile """ & appProfile & """ " & _
                "-Icon """ & strDir & "\rank-checker-search-bars.ico"" " & _
                "-WorkDir """ & strDir & """"
        WshShell.Run psCmd, 0, False
    Else
        WshShell.Run """" & edgePath & """ --app=" & appUrl & " --user-data-dir=""" & appProfile & """ --window-size=1100,820", 1, False
    End If
Else
    WshShell.Run appUrl, 1, False
End If
