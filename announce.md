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
- **Friendlier source-field parsing.** When you supply a single path, you no longer need to put quotes around it just because the path contains spaces. extCheck tests the entire trimmed source field as a single spec first; only when it is not a usable single spec does it fall back to space-tokenization. Quotes are only required when supplying multiple specs and at least one contains a space.
- **Output-directory create prompt.** If you press OK with an output directory that does not yet exist, extCheck asks "Create [path]?" with default Yes. Choosing No keeps the dialog open with focus on the output field.
- **Office automation alerts disabled.** Word, Excel, and PowerPoint application objects created by extCheck now have their `DisplayAlerts` property set to none, plus other prompt-suppression options including `Word.Application.Options.DoNotPromptForConvert = true` (suppresses the PDF-conversion dialog) and `AutomationSecurity = msoAutomationSecurityForceDisable` (blocks macros silently).
- **GUI progress display.** When extCheck runs in GUI mode, a small "Checking" status form now shows the current file's basename and the count of files **already completed**. When checking a single file, you see "report.docx — 0 of 1, 0%" while it is being processed (rather than no feedback while Word opens a multi-megabyte file).
- **Pre-pruning + structured results summary.** Before the check loop runs and before the progress UI opens, the file list is pruned in two passes: (1) unsupported extensions are silently dropped (logged when `-l` is on); (2) files whose CSV target already exists are dropped unless **Force replacements** is checked — these are counted as "skipped." The progress counter denominator is the post-pruning count, so percentages reflect actual work. The final summary is structured as up to three sections — `Checked N file(s):`, `Failed to check N file(s):`, and `Skipped N file(s). Check "Force replacements" to overwrite.` — each shown only when its count is non-zero, with singular "file" / plural "files" inflection. Failed entries include a short reason after the basename when one is available (`report.docx: could not open`); the full exception and stack trace go to the log when `-l` is on. The per-issue detail list is in the `<basename>.csv` file.
- **CLI vs GUI output styles.** In CLI mode (real console attached) basenames print inline as the loop runs — natural progress feedback for the console user. In GUI mode (and right-click invocations) the loop is silent on stdout; the progress status form shows the current file, and the structured summary is the final MessageBox. The structured summary is printed in both modes, but the per-name lists are suppressed in CLI mode (they would just repeat what already scrolled by).
- **Log header.** When `-l` (Log session) is enabled, the log file now begins with a clean header before the timestamped processing notifications: program name and version, a friendly run timestamp (`Run on May 1, 2026 at 2:30 PM`), and a `Parameters:` block listing each setting with both explicit and defaulted values resolved (Source, Output directory, Force replacements, View output, Use configuration, Log session, Show rules, GUI mode, Working directory, Command line). The header is followed by the normal timestamped log entries.
- **File Explorer right-click menu.** The installer now adds a "Check accessibility with extCheck" entry to the right-click menu (also reachable via Shift+F10) for `.docx`, `.xlsx`, `.pptx`, and `.md` files. Selecting a supported file and choosing this entry checks that one file and writes the CSV next to the source.

### Removed

- The `-strip-images` and `-plain-text` options from earlier versions are gone. extCheck is exclusively an accessibility checker now; it does not transform documents.

### Technical notes

- `comHelper.createApp` reports both the calling process's bitness and the likely Office bitness in its error message, so a mismatch surfaces with an actionable diagnosis.
- The icon is embedded in `extCheck.exe` at build time via `/win32icon`. Shortcuts inherit the icon; the .ico file does not need to ship with the installed program.
- The build script (`buildExtCheck.cmd`) writes `extCheck.exe` to the current working directory rather than a `dist\` subfolder, and it no longer copies `extCheck.csv` or `extCheck.ico` aside (those files already live in the working directory next to `extCheck.iss`). The script is straightforward enough that no forward `call :label` is used; it ships with CRLF line endings as defense-in-depth.

### Installer (`extCheck_setup.exe`)

- 64-bit only.
- Prompts for the installation directory on every run (default: `C:\Program Files\extCheck`). The directory page is now explicitly enabled (`DisableDirPage=no` in the .iss); previously it was at the Inno Setup default of `auto`, which silently skipped the page on reinstalls of the same `AppId`. The previous directory is pre-filled, so on a reinstall the user just presses Next to keep the same path.
- Includes a brief MIT-license summary on the welcome page.
- Installs only HTML versions of the documentation (`ReadMe.htm`, `Announce.htm`, `License.htm`); the Markdown counterparts and source/build/installer scripts live in the GitHub repository.
- The "Launch extCheck now" checkbox on the final page reminds the user that the desktop hotkey is Alt+Ctrl+X.
- Adds a "Report via e**x**tCheck" entry to the File Explorer right-click menu for all file types (registered under `HKLM\SOFTWARE\Classes\*\shell\extCheck`). The user is trusted to invoke the verb only on supported file types (`.docx`, `.xlsx`, `.pptx`, `.md`); on an unsupported file, extCheck reports that and exits. Registering once on `*` keeps the registry footprint small and matches the approach 2htm uses. The accelerator letter `x` matches the desktop hotkey accelerator. Uninstall removes the registry entries.
