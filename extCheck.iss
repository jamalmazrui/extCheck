; ============================================================================
; extCheck.iss  --  Inno Setup script for extCheck
;
; Builds extCheck_setup.exe, a 64-bit Windows installer for extCheck.
;
; Compile with Inno Setup 6.x (https://jrsoftware.org/isinfo.php). The
; .iss file expects to be opened from the dist\ folder produced by
; buildExtCheck.cmd; that folder must contain extCheck.exe, extCheck.csv,
; extCheck.ico, the documentation, and the source.
;
; The installer:
;   * Targets 64-bit Windows 10 (and later) only.
;   * Installs to C:\Program Files\extCheck by default.
;   * Adds a Start Menu group with shortcuts to extCheck, the README,
;     and the uninstaller.
;   * Adds a desktop shortcut whose hotkey is Alt+Ctrl+X. The shortcut
;     launches extCheck in GUI mode (-g) with saved-configuration
;     loading (-u). WorkingDir is the user's Documents folder so any
;     output folders or log files land somewhere writable.
;   * Adds no right-click Explorer verbs and no file associations.
;   * On uninstall, removes the program files but leaves
;     %LOCALAPPDATA%\extCheck\extCheck.ini intact (the user's saved
;     settings -- their filesystem, their call).
; ============================================================================

#define sAppName       "extCheck"
#define sAppVersion    "2.0"
#define sAppPublisher  "Jamal Mazrui"
#define sAppUrl        "https://github.com/JamalMazrui/extCheck"
#define sAppExeName    "extCheck.exe"
#define sAppCopyright  "Copyright (c) 2026 Jamal Mazrui. MIT License."

[Setup]
AppId={{E8C3D0A2-5B1F-4D8E-9A4C-2C3F4D7E8B9A}
AppName={#sAppName}
AppVersion={#sAppVersion}
AppVerName={#sAppName} {#sAppVersion}
AppPublisher={#sAppPublisher}
AppPublisherURL={#sAppUrl}
AppSupportURL={#sAppUrl}
AppUpdatesURL={#sAppUrl}
AppCopyright={#sAppCopyright}

DefaultDirName={autopf}\{#sAppName}
DefaultGroupName={#sAppName}
DisableProgramGroupPage=yes
DisableDirPage=auto
DisableReadyPage=no

OutputDir=.
OutputBaseFilename={#sAppName}_setup
Compression=lzma2
SolidCompression=yes

ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

WizardStyle=modern
SetupIconFile=extCheck.ico

Uninstallable=yes
UninstallDisplayIcon={app}\{#sAppExeName}
UninstallDisplayName={#sAppName} {#sAppVersion}

MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Note: extCheck.ico is NOT copied to {app} because the icon is embedded
; in extCheck.exe at build time (via /win32icon=extCheck.ico in the csc
; invocation). Shortcut icons in [Icons] inherit from the exe's embedded
; icon by default. The .ico file IS still needed at COMPILE time for the
; SetupIconFile= directive above, which gives extCheck_setup.exe itself
; an icon -- but that's a compile-time dependency only and does not need
; to ship with the installed program.
Source: "extCheck.exe";       DestDir: "{app}"; Flags: ignoreversion
Source: "extCheck.csv";       DestDir: "{app}"; Flags: ignoreversion
Source: "ReadMe.htm";         DestDir: "{app}"; Flags: ignoreversion
Source: "ReadMe.md";          DestDir: "{app}"; Flags: ignoreversion
Source: "license.htm";        DestDir: "{app}"; Flags: ignoreversion
Source: "announce.md";        DestDir: "{app}"; Flags: ignoreversion
Source: "announce.htm";       DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist
Source: "extCheck.cs";        DestDir: "{app}"; Flags: ignoreversion
Source: "extCheck.iss";       DestDir: "{app}"; Flags: ignoreversion
Source: "buildExtCheck.cmd";  DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu group. WorkingDir is the user's Documents folder so output
; CSV files and the optional extCheck.log land somewhere writable (the
; install dir under Program Files is not writable for non-admins).
Name: "{group}\{#sAppName}"; \
  Filename: "{app}\{#sAppExeName}"; \
  Parameters: "-g -u"; \
  WorkingDir: "{userdocs}"; \
  Comment: "Check Office and Markdown files for accessibility problems"

Name: "{group}\{#sAppName} README"; \
  Filename: "{app}\ReadMe.htm"; \
  WorkingDir: "{app}"; \
  Comment: "Documentation for {#sAppName}"

Name: "{group}\Uninstall {#sAppName}"; \
  Filename: "{uninstallexe}"; \
  Comment: "Remove {#sAppName} from this computer"

; Desktop shortcut with the Alt+Ctrl+X hotkey. Launches extCheck in
; GUI mode (-g) with saved-configuration loading (-u). The hotkey is
; not used by Windows or major office applications by default, but
; individual applications may intercept it when they have focus.
; WorkingDir is the user's Documents folder for the same writability
; reason as the Start Menu shortcut above.
Name: "{userdesktop}\{#sAppName}"; \
  Filename: "{app}\{#sAppExeName}"; \
  WorkingDir: "{userdocs}"; \
  Parameters: "-g -u"; \
  HotKey: Alt+Ctrl+X; \
  Comment: "Check accessibility (Alt+Ctrl+X)"

[Run]
; Post-install checkboxes shown on the final wizard page. Both default
; to checked; the user can uncheck either to skip.

; Launch extCheck (GUI mode). WorkingDir is the user's Documents folder
; so any output files or log file land somewhere writable.
FileName: "{app}\{#sAppExeName}"; \
  Parameters: "-g"; \
  WorkingDir: "{userdocs}"; \
  Description: "Launch {#sAppName} now"; \
  Flags: nowait postinstall skipifsilent

; Open the HTML documentation.
FileName: "{app}\ReadMe.htm"; \
  Description: "Read documentation for {#sAppName}"; \
  Flags: nowait postinstall skipifsilent shellexec
