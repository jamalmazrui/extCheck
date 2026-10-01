---
title: "extCheck Developer Notes"
author: "Jamal Mazrui"
---

# extCheck Developer Notes

How extCheck is built, released and laid out. Since 2.1.0 it is built on the Homer Development Kit (HomerDev), kit 1.43.4 or later, in `C:\HomerDev`.

## The project folder

`C:\extCheck` mirrors the installed tree:

- At the top: `extCheck.cs` (the program), `build.cmd`, `extCheck_setup.iss`, `extCheck.cmd`, `extCheck.ico`, `accept.inix`, `RepoFiles.txt`, `LocalFiles.txt`, `ReadMe` and `License`.
- `exec` — the built `extCheck.exe`. Never in git.
- `help` — this document and the others: `extCheck` (the guide), `Announce`, `Developer`, `History`, `Hotkeys`, each as `.md` and `.htm`, and `extCheck.csv`, the rule registry, which the build writes by running `exec\extCheck.exe -rules -o help`.
- `logs` — one log per run of the build or any tool.
- `scripts` — the kit's tools, refreshed from `C:\HomerDev\scripts` by every build.

`version.txt` holds the version and lives only on this machine. The build writes `Version.cs` from it, and the installer reads it directly.

## The four steps

1. `build` — steps the version (`build nobump` keeps it), compiles `exec\extCheck.exe`, writes the rule registry into `help`, writes each `.htm` from its `.md`, puts the project's files in the Homer encoding, checks that the installer ships every file in `help`, and builds `extCheck_setup.exe`. Its log is `logs\extCheck-build-yyyyMMdd-HHmmss.log`.
2. `scripts\push "message"` — rewrites the whitelist `.gitignore` from `RepoFiles.txt`, commits and pushes.
3. `scripts\tidy` and `scripts\tidy --do-it` — the periodic clean.
4. `scripts\release` — runs `scripts\check`, then tags the pushed commit with the version stamped in `extCheck_setup.exe` and publishes the installer.

Try the fresh build with `extCheck -g` (the `extCheck.cmd` at the top runs `exec\extCheck.exe`).

## The compiler and the kit's classes

The build finds the Roslyn C# compiler with `vswhere`, or installs the free Visual Studio Build Tools with winget. The Framework's own `csc.exe`, which extCheck used to be built with, stops at C# 5 and cannot compile the kit.

extCheck compiles eight of the kit's classes straight from `C:\HomerDev\CSharp`: Elevate, Inix, Lbc, Log, Paths, Say, Util and Web. Its dialog, `guiDialog.show`, is an LbcDialog, built the way urlFido's is: bands for a field and its button, checkboxes below a separator, and `runWithButtons` with OK, Guide, Default settings and Cancel, looping back to the dialog after Guide, Default settings, or an output folder the user declined to create. Lbc needs Elevate, Inix, Log, Say and Util.

- `Log` keeps the session log in `%LOCALAPPDATA%\extCheck\logs`. extCheck's own `logger` writes every line to it as well as to the optional `extCheck.log`, and `Main` records anything `program.run` did not catch. Log needs Paths and Say.
- `Paths.configs()` is where `extCheck.inix` lives, and `Paths.installedFolder` is how Help finds `help\extCheck.htm` from `exec`.
- `InixCodec.writeValue` saves each setting in place.
- `Elevate` answers F11 (and needs Web).

## Source layout

The whole program is one C# file: `extCheck.cs`. It uses standard `System.Windows.Forms` for the parameter dialog and the COM `dynamic` keyword to drive Office. There are no third-party dependencies. The classes inside `extCheck.cs` are arranged as a shared infrastructure layer (`issue`, `results`, `shared`, `comHelper`, `logger`, `configManager`, `guiDialog`) plus per-format modules (`docxModule`, `xlsxModule`, `pptxModule`, `mdModule`), with a top-level `program` class that parses arguments, optionally shows the dialog, and dispatches to the right module per file extension.

## Threading and bitness

`Main` is decorated with `[STAThread]`. This is required for two reasons:

- Office COM automation requires a single-threaded apartment. Without it, Word/Excel/PowerPoint COM servers can disconnect mid-operation with HRESULT 0x80010108 (RPC_E_DISCONNECTED) or 0x80010114 (OLE_E_OBJNOTCONNECTED).
- WinForms common dialogs (`OpenFileDialog`, `FolderBrowserDialog`) require an STA thread.

The build is `/platform:x64`. Office COM automation requires the controller process and the installed Office to share the same bitness. Modern Office is 64-bit by default; if a user has 32-bit Office, `com.createApp` surfaces a clear error message pointing at the mismatch and recommending a 32-bit rebuild.


## Conventions

- Camel Type for C#, as `C:\HomerDev\help\CamelType_CSharp.md` describes. Constants added since the move to the kit carry the `c_` prefix. The `o` prefix is reserved for COM objects.
- extCheck shares identifier names with urlCheck, urlFido and 2htm: `sProgramName`, `sProgramVersion`, `sConfigFileName`, `sLogFileName`, `sOutputDir`, the `iLayout*` constants, and the `logger` surface `open`, `close`, `info`, `warn`, `error`, `debug`.
