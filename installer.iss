; Inno Setup Script for SEO Toolkit Pro
; Download Inno Setup from: https://jrsoftware.org/isdl.php

#define MyAppName "SEO Toolkit Pro"
#define MyAppVersion "4.9.4"
#define MyAppPublisher "Vishal Chhipa"
#define MyAppExeName "Start Tool.vbs"

[Setup]
AppId={{E8F3A1B2-7C4D-4E5F-9A6B-1C2D3E4F5A6B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer_output
OutputBaseFilename=SEOToolkitPro_Setup
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\rank-checker-search-bars.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
DisableProgramGroupPage=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UsePreviousAppDir=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"
Name: "startmenu"; Description: "Create a &Start Menu shortcut"; GroupDescription: "Additional shortcuts:"; Flags: checkedonce

[Files]
; Main application files
Source: "web_app_batch.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "engine.py";        DestDir: "{app}"; Flags: ignoreversion
Source: "da_checker.py";    DestDir: "{app}"; Flags: ignoreversion
Source: "health_audit.py";  DestDir: "{app}"; Flags: ignoreversion
Source: "auth.py";          DestDir: "{app}"; Flags: ignoreversion
Source: "updater.py";       DestDir: "{app}"; Flags: ignoreversion
Source: "gsc_audit.py";      DestDir: "{app}"; Flags: ignoreversion
Source: "brief_analysis.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "google_ads_keywords.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "site_audit.py";     DestDir: "{app}"; Flags: ignoreversion
Source: "config.json";       DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist
Source: "Start Tool.vbs";   DestDir: "{app}"; Flags: ignoreversion
Source: "app_launch.ps1";   DestDir: "{app}"; Flags: ignoreversion
Source: "splash.ps1";       DestDir: "{app}"; Flags: ignoreversion
Source: "rank-checker-search-bars.ico"; DestDir: "{app}"; Flags: ignoreversion

; Embedded Python
Source: "python\*"; DestDir: "{app}\python"; Flags: ignoreversion recursesubdirs createallsubdirs

; Templates
Source: "templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs

; Static files
Source: "static\*"; DestDir: "{app}\static"; Flags: ignoreversion recursesubdirs createallsubdirs

; SEO On-Page scripts
Source: "scripts\*"; DestDir: "{app}\scripts"; Flags: ignoreversion recursesubdirs createallsubdirs

; Extensions (Buster, Urban VPN, etc.)
Source: "extensions\*"; DestDir: "{app}\extensions"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
Name: "{app}\uploads"
Name: "{app}\profiles"

[Icons]
; Desktop shortcut
Name: "{autodesktop}\{#MyAppName}"; Filename: "wscript.exe"; Parameters: """{app}\Start Tool.vbs"""; WorkingDir: "{app}"; IconFilename: "{app}\rank-checker-search-bars.ico"; Tasks: desktopicon
; Start Menu shortcuts
Name: "{group}\{#MyAppName}";           Filename: "wscript.exe"; Parameters: """{app}\Start Tool.vbs"""; WorkingDir: "{app}"; IconFilename: "{app}\rank-checker-search-bars.ico"; Tasks: startmenu
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"; Tasks: startmenu

; Remove artifacts from older versions on upgrade (Inno keeps orphaned files otherwise)
[InstallDelete]
Type: files; Name: "{group}\Stop {#MyAppName}.lnk"
Type: files; Name: "{app}\Stop Tool.bat"

[Run]
Filename: "wscript.exe"; Parameters: """{app}\Start Tool.vbs"""; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "taskkill"; Parameters: "/F /IM python.exe";      Flags: runhidden; RunOnceId: "KillPython"
Filename: "taskkill"; Parameters: "/F /IM msedgedriver.exe"; Flags: runhidden; RunOnceId: "KillEdgeDriver"
Filename: "taskkill"; Parameters: "/F /IM chromedriver.exe"; Flags: runhidden; RunOnceId: "KillChromeDriver"

[UninstallDelete]
Type: filesandordirs; Name: "{app}\chrome_profile*"
Type: filesandordirs; Name: "{app}\profiles"
Type: filesandordirs; Name: "{app}\__pycache__"
Type: filesandordirs; Name: "{app}\*.log"
Type: filesandordirs; Name: "{app}\*.csv"
Type: filesandordirs; Name: "{app}\autosave_results.json"
Type: filesandordirs; Name: "{app}\.grc_port"
Type: filesandordirs; Name: "{app}\downloads"
Type: filesandordirs; Name: "{app}\uploads"
Type: filesandordirs; Name: "{app}\python"
Type: filesandordirs; Name: "{app}\templates"
Type: filesandordirs; Name: "{app}\static"
Type: filesandordirs; Name: "{app}\extensions"
Type: filesandordirs; Name: "{localappdata}\SEO Toolkit Pro\edge_app"
Type: dirifempty;     Name: "{app}"

[Code]
// Kill running instances before install/update
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssInstall then
  begin
    Exec('taskkill', '/F /IM python.exe',        '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    Exec('taskkill', '/F /IM msedgedriver.exe',   '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    Exec('taskkill', '/F /IM chromedriver.exe',   '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    Sleep(1000);
  end;
end;
