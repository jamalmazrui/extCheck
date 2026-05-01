; ============================================================================
; extCheck.iss  --  Inno Setup script for extCheck
;
; Compile with the Inno Setup IDE (ISCC.exe) to produce extCheck_setup.exe.
; The resulting installer:
;   - Targets 64-bit Windows 10 (and later) only.
;   - Requires administrator privileges.
;   - Prompts the user for the installation directory; default is
;     C:\Program Files\extCheck.
;   - Shows a brief MIT license summary on the welcome page (no extra
;     wizard screen). The full license text is installed alongside
;     the program as License.htm.
;   - Adds a Start Menu group with shortcuts to extCheck, the README,
;     and the uninstaller.
;   - Adds a desktop shortcut whose hotkey is Alt+Ctrl+X. The shortcut
;     launches extCheck in GUI mode (-g) with saved-configuration
;     loading (-u). WorkingDir is the user's Documents folder so any
;     output folders or log files land somewhere writable.
;   - Adds no right-click Explorer verbs and no file associations.
;   - On uninstall, removes the program files but leaves
;     %LOCALAPPDATA%\extCheck\extCheck.ini intact (the user's saved
;     settings -- their filesystem, their call).
;
; This installer ships only the runtime distribution (the .exe, the
; runtime CSV rule registry, the documentation in HTML form, and the
; license). The Markdown sources, the C# source, the build script,
; and this .iss script live in the GitHub repository.
; ============================================================================

#define sAppName       "extCheck"
#define sAppVersion    "2.0"
#define sAppPublisher  "Jamal Mazrui"
#define sAppUrl        "https://github.com/JamalMazrui/extCheck"
#define sAppExeName    "extCheck.exe"
#define sAppCopyright  "Copyright (c) 2026 Jamal Mazrui. MIT License."
#define sHotKey        "Alt+Ctrl+X"

[Setup]
AppId={{E8C3D0A2-5B1F-4D8E-9A4C-2C3F4D7E8B9A}
AppName={#sAppName}
AppVersion={#sAppVersion}
AppVerName={#sAppName} {#sAppVersion}
AppPublisher={#sAppPublisher}
AppPublisherURL={#sAppUrl}
AppSupportURL={#sAppUrl}
AppUpdatesURL={#sAppUrl}/releases
AppCopyright={#sAppCopyright}
VersionInfoVersion={#sAppVersion}

; Install under Program Files. {autopf} resolves to "Program Files"
; on 64-bit Windows when the installer runs in 64-bit mode (see
; ArchitecturesInstallIn64BitMode below). The user can override this
; default on the wizard's directory page.
DefaultDirName={autopf}\{#sAppName}
DefaultGroupName={#sAppName}
DisableProgramGroupPage=yes
UsePreviousAppDir=yes

; Force the "Select Destination Location" page to always be shown,
; even on reinstall. Without this, DisableDirPage defaults to "auto",
; which means: hide the directory page if a prior install of the same
; AppId is detected. We want the page shown every time so the user
; can review the install location, and so it is obviously editable.
; UsePreviousAppDir=yes pre-fills the field with the previous
; directory, so the user just presses Next on a reinstall to keep the
; same path -- but they can also change it.
DisableDirPage=no
UsePreviousGroup=yes

OutputDir=.
OutputBaseFilename={#sAppName}_setup
Compression=lzma2
SolidCompression=yes
SetupIconFile={#sAppName}.ico
WizardStyle=modern

; Installer requires admin to write to Program Files.
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=

; 64-bit Windows only.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

Uninstallable=yes
UninstallDisplayIcon={app}\{#sAppExeName}
UninstallDisplayName={#sAppName} {#sAppVersion}

MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
; Replace the default welcome-page body text with one that includes a
; brief MIT license notice. This satisfies the requirement that the
; license summary appear on an existing wizard screen rather than on
; an additional dedicated page (which is what LicenseFile= would
; produce). The full license text is installed alongside the program.
WelcomeLabel2=This will install [name/ver] on your computer.%n%n[name] is licensed under the MIT License: free to use, copy, modify, and distribute; provided "as is" with no warranty. The full license text will be installed as License.htm in the program folder.%n%nIt is recommended that you close all other applications before continuing.

[Files]
; The runtime distribution: the executable, the rule-registry CSV
; (used by extCheck -rules), the HTML docs, and the license. The icon
; is embedded in extCheck.exe at build time (csc /win32icon flag),
; so the .ico does not need to ship in the install directory.
Source: "{#sAppName}.exe";    DestDir: "{app}"; Flags: ignoreversion
Source: "{#sAppName}.csv";    DestDir: "{app}"; Flags: ignoreversion
Source: "ReadMe.htm";         DestDir: "{app}"; Flags: ignoreversion
Source: "Announce.htm";       DestDir: "{app}"; Flags: ignoreversion
Source: "License.htm";        DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu group. WorkingDir is the user's Documents folder so output
; CSV files and the optional extCheck.log land somewhere writable (the
; install dir under Program Files is not writable for non-admins).
Name: "{group}\{#sAppName}"; \
  Filename: "{app}\{#sAppExeName}"; \
  Parameters: "-g -u"; \
  WorkingDir: "{userdocs}"; \
  Comment: "Check Office and Markdown files for accessibility problems"

Name: "{group}\{#sAppName} ReadMe"; \
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
  HotKey: {#sHotKey}; \
  Comment: "Check accessibility ({#sHotKey})"

[Registry]
; File Explorer "Report via extCheck" right-click menu entry.
; Registered under HKLM\Software\Classes\*\shell\extCheck so the
; verb appears for all file types. The user is trusted to invoke it
; only on supported files (.docx, .xlsx, .pptx, .md); if invoked on
; an unsupported extension, extCheck reports that and exits.
;
; Registering once on "*" rather than four times under
; SystemFileAssociations\<ext> keeps the registry footprint small,
; matches the approach 2htm uses, and gives the user a single,
; consistent entry to find regardless of file type.
;
; The verb passes "%1" (the full absolute path of the right-clicked
; file) to extCheck.exe with -f to force overwriting any existing
; CSV report so repeated right-clicks refresh rather than skip.
;
; The Shift+F10 keyboard shortcut in File Explorer opens the same
; context menu, so this entry is reachable both by mouse and by
; keyboard. The "e&xtCheck" ampersand makes "x" the accelerator
; letter, matching the desktop hotkey Alt+Ctrl+X.
;
; Selecting multiple files and invoking the verb runs extCheck.exe
; once per selected file (the standard shell behavior for verb
; commands that include "%1"). Each invocation processes one file
; and writes its CSV next to the source.
;
; The uninsdeletekey flag on the parent verb key causes Inno Setup
; to remove the entire subtree (including the command subkey)
; automatically on uninstall.

Root: HKLM; Subkey: "SOFTWARE\Classes\*\shell\extCheck"; \
  ValueType: string; ValueName: ""; ValueData: "Report via e&xtCheck"; \
  Flags: uninsdeletekey
Root: HKLM; Subkey: "SOFTWARE\Classes\*\shell\extCheck"; \
  ValueType: string; ValueName: "Icon"; ValueData: """{app}\{#sAppExeName}"",0"
Root: HKLM; Subkey: "SOFTWARE\Classes\*\shell\extCheck\command"; \
  ValueType: string; ValueName: ""; \
  ValueData: """{app}\{#sAppExeName}"" -f ""%1"""

[Run]
; Post-install checkboxes shown on the final wizard page. Both default
; to checked; the user can uncheck either to skip. The launch checkbox
; label includes a reminder of the desktop hotkey so the user notices
; and remembers it.

FileName: "{app}\{#sAppExeName}"; \
  Parameters: "-g"; \
  WorkingDir: "{userdocs}"; \
  Description: "Launch {#sAppName} now (desktop hotkey: {#sHotKey})"; \
  Flags: nowait postinstall skipifsilent

FileName: "{app}\ReadMe.htm"; \
  Description: "Read documentation for {#sAppName}"; \
  Flags: nowait postinstall skipifsilent shellexec
