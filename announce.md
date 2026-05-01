# extCheck Release Notes

## Version 2.0

This is the first release of extCheck as a 64-bit GUI/CLI hybrid program.

### What's new

- **64-bit build.** extCheck is now built `/platform:x64`. Microsoft Office's COM automation requires bitness matching between the controller and Office; modern Office (Microsoft 365, Office 2019+, Office 2024) is 64-bit by default. If a user has 32-bit Office, `com.createApp` surfaces a clear bitness-mismatch error.
- **STAThread.** `Main` is decorated with `[STAThread]`. This satisfies Office COM's single-threaded apartment requirement and is also a precondition for WinForms common dialogs.
- **Parameter dialog.** Launching extCheck without arguments from a GUI shell (Explorer double-click, Start menu, desktop hotkey) now shows a parameter dialog with controls for source files, output directory, and the option checkboxes (force replacements, view output, log session, use configuration). The dialog can also be invoked explicitly with `-g`.
- **CLI options.** New flags: `-g/--gui-mode`, `-o/--output-dir`, `-f/--force`, `-l/--log`, `-u/--use-configuration`, and `--view-output`. These mirror the dialog controls one-to-one, so a workflow prototyped in the GUI is straightforward to translate to a batch file.
- **Configuration persistence.** With the `-u` flag (or the dialog's "Use configuration" checkbox), extCheck reads and writes its settings at `%LOCALAPPDATA%\extCheck\extCheck.ini`. Without the flag, extCheck leaves no settings on disk.
- **Diagnostic log.** With `-l` (or "Log session"), extCheck writes a fresh `extCheck.log` in the output directory. The log captures program version, command-line arguments, GUI auto-detection, per-file open/check/quit events, and any errors. Each session truncates any prior log.
- **Desktop hotkey Alt+Ctrl+X.** The installer adds a desktop shortcut whose hotkey is Alt+Ctrl+X. Pressing it from anywhere in Windows opens the extCheck parameter dialog.
- **Camel Type coding standard.** All variable and identifier names follow the project's Camel Type style. The `o` prefix is reserved for COM objects (the Office automation bridge).
- **Cross-program naming.** Identifier names for shared concepts match across the three companion tools (urlCheck, extCheck, 2htm). The program-name and version constants are `sProgramName` and `sProgramVersion`; the config and log filename constants are `sConfigDirName`, `sConfigFileName`, `sLogFileName`; the source and output-directory variables are `sSource` and `sOutputDir`; the GUI layout constants are all `iLayout*`. The COM helper class is `comHelper` (was `com`). The `logger` surface is now uniform: `open`, `close`, `info`, `warn`, `error`, `debug`.
- **Picker initial directory.** The Browse source and Choose output buttons now open at the directory derived from the text-field value when that value points to an existing path (whether the user just typed it or it was loaded from a saved configuration), and at the user's Documents folder otherwise.

### Removed

- The `-strip-images` and `-plain-text` options from earlier versions are gone. extCheck is exclusively an accessibility checker now; it does not transform documents.

### Technical notes

- `comHelper.createApp` reports both the calling process's bitness and the likely Office bitness in its error message, so a mismatch surfaces with an actionable diagnosis.
- The icon is embedded in `extCheck.exe` at build time via `/win32icon`. Shortcuts inherit the icon; the .ico file does not need to ship with the installed program.
- The build script (`buildExtCheck.cmd`) writes `extCheck.exe` to the current working directory rather than a `dist\` subfolder, and it no longer copies `extCheck.csv` or `extCheck.ico` aside (those files already live in the working directory next to `extCheck.iss`). The script is straightforward enough that no forward `call :label` is used; it ships with CRLF line endings as defense-in-depth.

### Installer (`extCheck_setup.exe`)

- 64-bit only.
- Prompts for the installation directory (default: `C:\Program Files\extCheck`).
- Includes a brief MIT-license summary on the welcome page.
- Installs only HTML versions of the documentation (`ReadMe.htm`, `Announce.htm`, `License.htm`); the Markdown counterparts and source/build/installer scripts live in the GitHub repository.
- The "Launch extCheck now" checkbox on the final page reminds the user that the desktop hotkey is Alt+Ctrl+X.
