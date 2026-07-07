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

' Start the Flask server FIRST so the window opens quickly. The updater used to run
' (and block) BEFORE the server; on slower machines that pushed startup past the wait
' and showed "could not start", and the long wait made users re-launch (the 2-3 times).
' Updates are now pulled AFTER the server is up (below) and by the app's own auto-check.
WshShell.Run "cmd /c ""set GRC_NO_BROWSER=1&& set GRC_PORT_FILE=" & portFile & "&& """ & strDir & "\python\python.exe"" -s """ & strDir & "\web_app_batch.py""""", 0, False

' Show a lightweight "Starting..." splash so the user has feedback while the server
' comes up. It self-closes when the port file appears (or times out). Cosmetic only.
If FSO.FileExists(strDir & "\splash.ps1") Then
    WshShell.Run "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strDir & "\splash.ps1""", 0, False
End If

' Wait for port file (server writes it when ready) — generous for slower machines.
Dim port, appUrl, attempts
port = ""
attempts = 0
Do While attempts < 80
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

' Server is up — pull OTA updates in the BACKGROUND (does not delay the window; new
' files apply on the next launch, and the app shows an "update installed" note).
WshShell.Run "cmd /c """ & strDir & "\python\python.exe"" -s """ & strDir & "\updater.py""", 0, False

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

    ' Open the app window DIRECTLY in Edge (no PowerShell helper). One antivirus
    ' (Quick Heal) flagged app_launch.ps1's window-icon code as a threat and quarantined
    ' it, which broke launching. The title bar + favicon already show the app icon; the
    ' taskbar button groups under Edge, a harmless cosmetic Edge --app limitation.
    WshShell.Run """" & edgePath & """ --app=" & appUrl & " --user-data-dir=""" & appProfile & """ --window-size=1100,820 --no-first-run --no-default-browser-check", 1, False
Else
    WshShell.Run appUrl, 1, False
End If
