---
title: "extCheck — Accessibility Checker for Office and Markdown Files"
author: "Jamal Mazrui"
date: "April 29, 2026"
description: "Produce accessibility reports for popular file formats from the command line or a small dialog."
---

# extCheck

**Author:** Jamal Mazrui
**License:** MIT

`extCheck` is a Windows tool that checks Microsoft Word, Excel, PowerPoint, and Pandoc Markdown files for accessibility problems. For each file you give it, extCheck writes a CSV report listing issues with rule IDs, locations, problem descriptions, and remediation guidance. It can be used from the command line or from a small parameter dialog. The dialog is designed to be friendly to screen readers.

The whole extCheck project may be downloaded as a single zip archive from:

<https://github.com/JamalMazrui/extCheck/archive/main.zip>

---

## What you need

- Windows 10 or later (64-bit)
- Microsoft Word, Excel, or PowerPoint installed to check the corresponding `.docx`, `.xlsx`, or `.pptx` files
- No Office installation needed for `.md` files
- An internet connection is **not** required

You do **not** need to install .NET separately. The .NET Framework 4.8.1 used by extCheck ships in-box with Windows 10 (since version 22H2) and Windows 11.

**Bitness note.** extCheck is built as a 64-bit program, and Microsoft Office automation requires the controller process and the installed Office to share the same bitness. Modern Office (Microsoft 365, Office 2019+, Office 2024) is 64-bit by default, so this matches the common case. If you have 32-bit Office on your machine, extCheck will surface a clear error pointing at the bitness mismatch; you can either install 64-bit Office or rebuild extCheck with `/platform:x86` (see Development below).

---

## Installing

Run `extCheck_setup.exe` (the setup wizard) and follow the prompts. By default extCheck installs to `C:\Program Files\extCheck`, adds a Start-menu shortcut, and adds a desktop shortcut whose hotkey is **Alt+Ctrl+X**. Pressing Alt+Ctrl+X from anywhere in Windows opens the extCheck dialog.

The installer never adds a right-click verb or any other Explorer integration. It does not change file associations.

---

## Running extCheck

There are two ways to run it.

### From the dialog (easiest)

Launch extCheck from any of these:

- The desktop shortcut (or its **Alt+Ctrl+X** hotkey)
- The Start-menu shortcut
- Double-clicking `extCheck.exe` in File Explorer
- A Run dialog (`Win+R`) typing `extCheck`

The parameter dialog appears. Fill in the fields you want and press OK to start checking. Press F1 inside the dialog for in-context help.

The dialog has these controls. Each label has an underlined letter that you can press with **Alt** to jump straight to that control (the underlined letter is shown in brackets below):

- **Source files** [S] — what to check. Enter one file path, a wildcard pattern (e.g., `*.docx`, `reports\*.md`), or several of either separated by spaces. Wrap paths containing spaces in double quotes. The Browse source [B] button opens a file picker.
- **Output directory** [O] — where the per-file CSV reports are written. Blank means the current working directory. The Choose output [C] button opens a folder picker.
- **Force replacements** [F] — overwrite an existing `<basename>.csv` instead of skipping the input. Without this, extCheck skips an input whose CSV already exists in the output directory.
- **View output** [V] — when all files have been checked, open the output directory in File Explorer.
- **Log session** [L] — write a fresh `extCheck.log` file in the output directory (or current directory if no output directory is set). Useful for diagnostics; replaces any prior log.
- **Use configuration** [U] — load these field values from a saved configuration file at startup, and save them back when you press OK. The configuration lives at `%LOCALAPPDATA%\extCheck\extCheck.ini`. Without this checkbox, extCheck leaves no settings on disk.
- **Help** [H] — show this help summary and offer to open the full README in your browser. F1 also shows Help.
- **Default settings** [D] — clear all fields, uncheck all boxes, and delete the saved configuration if any.
- **OK** / **Cancel** — start the check, or cancel without checking. Enter is OK; Esc is Cancel.

When all files have been checked, a final results dialog summarizes what was processed.

### From the command line

Open a Command Prompt, navigate to the folder containing the files you want to check, and run:

```cmd
extCheck report.docx
```

A `report.csv` is written to the current directory and a summary prints to the console.

Show help:

```cmd
extCheck -h
```

Show version:

```cmd
extCheck -v
```

Check several files at once:

```cmd
extCheck *.docx *.md
extCheck docs\*.docx data\*.xlsx slides\*.pptx
```

Write CSV reports to a specific directory:

```cmd
extCheck *.md -o reports
```

Open the output directory in File Explorer when done:

```cmd
extCheck *.docx -o reports --view-output
```

Write a fresh diagnostic log next to the reports:

```cmd
extCheck *.docx -o reports -l
```

Generate the rule registry (every rule extCheck knows about, with WCAG criteria and remediation guidance):

```cmd
extCheck -rules
```

Or write the rule registry to a specific directory:

```cmd
extCheck -rules -o references
```

The dialog can also be launched from the command line with `-g`:

```cmd
extCheck -g
```

When invoked without arguments from a GUI shell (Explorer double-click, Start-menu shortcut, desktop hotkey), extCheck shows the dialog automatically. When invoked without arguments from a console shell, it prints help and exits. The `-g` flag forces GUI mode regardless.

---

## Command-line options

| Option | Long form | Description |
|---|---|---|
| `-h` | `--help` | Show usage and exit |
| `-v` | `--version` | Show version and exit |
| `-g` | `--gui-mode` | Show the parameter dialog |
| `-o <d>` | `--output-dir <d>` | Write reports to `<d>` (created if missing); defaults to current directory |
| `-rules` | | Write the rule registry as `extCheck.csv` to the output directory and exit |
| `-f` | `--force` | Overwrite an existing CSV instead of skipping the input |
|   | `--view-output` | After checking, open the output directory in File Explorer |
| `-l` | `--log` | Write `extCheck.log` (UTF-8 with BOM) in the output directory; replaced each session |
| `-u` | `--use-configuration` | Read saved defaults from `%LOCALAPPDATA%\extCheck\extCheck.ini` |

Every option in the GUI corresponds one-to-one with a command-line flag, so a workflow prototyped in the dialog can be translated to a batch file without surprises.

---

## Supported file formats

| Extension | Format |
|-----------|--------|
| .docx | Microsoft Word document |
| .xlsx | Microsoft Excel workbook |
| .pptx | Microsoft PowerPoint presentation |
| .md | Pandoc Markdown file |

`temp.*` (with a literal asterisk for the extension) expands to all supported extensions, so you can write `extCheck temp.*` to check `temp.docx`, `temp.xlsx`, `temp.pptx`, and `temp.md` if any exist.

---

## Output

For each file evaluated, a CSV named `<basename>.csv` is written to the output directory (`-o`, or the current working directory if no `-o`). The CSV columns are:

- **RuleID** — unique identifier for the rule (e.g., `MissingAltText`, `DuplicateHeadingText`)
- **Source** — `MSAC` (Microsoft Office Accessibility Checker categories) or `AXE` (axe-core WCAG equivalents)
- **Category** — high-level grouping (e.g., "Image", "Heading", "Link")
- **Location** — sheet name, slide label, line number, or `(Document)`
- **Context** — short snippet of the offending content
- **Message** — what the rule found and why it matters
- **Remediation** — how to fix it

Results are also printed to the console. Each issue is printed as:

```
[RuleID] (Source) Location | Category | Context
  Problem:   ...
  Remediate: ...
```

At the end of a multi-file run, a total count is printed.

---

## The rule registry

Running `extCheck -rules` writes `extCheck.csv` containing every rule extCheck knows about. Each row includes:

- **RuleID** — unique identifier
- **MSOfficeCategory** — the Office Accessibility Checker category this maps to
- **WCAGCriteria** — the WCAG 2.1 success criterion number
- **Severity** — `Error` (definite barrier) or `Warning` (likely barrier, context-dependent)
- **AppliesTo** — which file formats the rule checks
- **Description** — what the rule looks for and why it matters
- **Remediation** — how to fix it

The rules come from two complementary sources:

**MSAC** rules mirror the categories used by the built-in Microsoft Office Accessibility Checker: missing alternative text, missing table headers, heading issues, repeated blank characters, blank cells used for formatting, merged cells, complex tables, use of color alone, object titles, list usage, text contrast.

**AXE** rules are adapted from the axe-core open-source accessibility engine maintained by Deque Systems. These cover areas the Office checker does not address, including hyperlink distinguishability, duplicate heading text, non-descriptive link text, form field labels, slide reading order, animation timing, and code block language specifiers for Markdown.

---

## Configuration file

When **Use configuration** is checked in the dialog (or `-u` is on the command line), extCheck reads and writes a small INI file at:

```
%LOCALAPPDATA%\extCheck\extCheck.ini
```

It stores the source files string, the output directory, and the option checkboxes (force, view output, log session). Without **Use configuration**, extCheck leaves nothing on disk between runs. **Default settings** in the dialog deletes this file.

---

## Log file

When **Log session** is checked (or `-l` is on the command line), extCheck writes a fresh `extCheck.log` to the output directory (or current directory if no output directory is set). Any prior log is deleted at the start of the run, so the file always reflects only the current session.

The log captures: program version, command-line arguments, GUI auto-detection, the resolved output directory, per-file open/check/quit events, and any errors.

The log is UTF-8 with a byte-order mark, so Notepad opens it correctly.

---

## Notes

- extCheck only reports automatically-detectable violations. It does not replace manual testing.
- **PowerPoint** requires a visible application window; headless mode is not supported. The window is minimized automatically when extCheck launches PowerPoint, and closed when checking is complete.
- **Markdown** checking requires no Office software. The checker reads the file directly and evaluates Pandoc-flavored Markdown conventions.
- **False positives** are possible. For example, the empty-alt-text rule for Markdown flags all empty `[]` alt attributes, but empty alt is correct for purely decorative images. The all-caps rule ignores sequences of six characters or fewer to allow common acronyms. Review each flagged item in context before remediating.

---

## Development

This section is for developers who want to build `extCheck.exe` from source.

### Source layout

The whole program is one C# file: `extCheck.cs`. It uses standard `System.Windows.Forms` for the parameter dialog and the COM `dynamic` keyword to drive Office. There are no third-party dependencies.

The classes inside `extCheck.cs` are arranged as a shared infrastructure layer (`issue`, `results`, `shared`, `com`, `logger`, `configManager`, `guiDialog`) plus per-format modules (`docxModule`, `xlsxModule`, `pptxModule`, `mdModule`), with a top-level `program` class that parses arguments, optionally shows the dialog, and dispatches to the right module per file extension.

### Coding style

The source uses what the author calls "Camel Type" (C# variant): Hungarian prefix notation for variables (`b` for boolean, `i` for integer, `s` for string, `ls` for `List<T>`, `d` for dictionary, `o` for other object types, etc.), lower camelCase for everything other than where the language requires PascalCase (class names; `public` API surfaces). Constants follow the same naming as variables — only the `const` or `static readonly` keyword conveys constant-ness. The style is tuned for screen-reader productivity (predictable token shapes that read well aloud).

### Threading and bitness

`Main` is decorated with `[STAThread]` for two reasons:

- Office COM automation requires a single-threaded apartment. Without it, Word/Excel/PowerPoint COM servers can disconnect mid-operation with HRESULT 0x80010108 (RPC_E_DISCONNECTED) or 0x80010114 (OLE_E_OBJNOTCONNECTED). PowerPoint shape iteration and Excel `UsedRange.Value2` are particularly sensitive.
- WinForms common dialogs (`OpenFileDialog`, `FolderBrowserDialog`) require an STA thread.

The build is `/platform:x64`. Office COM automation requires the controller process and the installed Office to share the same bitness. Modern Office is 64-bit by default; if a user has 32-bit Office, `com.createApp` surfaces a clear error message pointing at the mismatch and recommending a 32-bit rebuild.

### Prerequisites

- The .NET Framework 4.8.1 Developer Pack (provides `csc.exe` and the 4.8.1 reference assemblies). Install from <https://dotnet.microsoft.com/download/dotnet-framework/net481>.
- Inno Setup 6.x to compile the installer from `extCheck.iss`. Download from <https://jrsoftware.org/isinfo.php>.

### Building the executable

Run the included script. It auto-detects `csc.exe`, compiles `extCheck.cs` with the icon embedded, and copies the rule registry CSV into the `dist\` folder:

```cmd
buildExtCheck.cmd
```

The result is `dist\extCheck.exe` plus `dist\extCheck.csv` and `dist\extCheck.ico`.

To build manually:

```cmd
csc /target:exe /platform:x64 /optimize+ ^
    /win32icon:extCheck.ico ^
    /reference:System.Windows.Forms.dll ^
    /reference:System.Drawing.dll ^
    /out:dist\extCheck.exe ^
    extCheck.cs
```

### Building the installer

After `buildExtCheck.cmd` produces `dist\extCheck.exe`, copy the rest of the source tree into `dist\` so that Inno Setup can find them:

```
dist\extCheck.exe         (built; the icon is embedded inside it)
dist\extCheck.csv         (rule registry, copied by buildExtCheck.cmd)
dist\extCheck.ico         (copied by buildExtCheck.cmd; needed only at
                           installer compile time for the setup wizard's
                           own icon, not shipped with the installed program)
dist\ReadMe.htm           (you provide; rendered from ReadMe.md)
dist\ReadMe.md            (you provide)
dist\license.htm          (you provide)
dist\announce.htm         (you provide; release notes)
dist\announce.md          (you provide)
dist\extCheck.cs          (you provide; the source)
dist\buildExtCheck.cmd    (you provide)
dist\extCheck.iss         (you provide)
```

Open `extCheck.iss` in Inno Setup and click Compile. The result is `dist\extCheck_setup.exe`.

The installer:

- Installs to `C:\Program Files\extCheck` by default
- Adds a desktop shortcut whose hotkey is **Alt+Ctrl+X**
- Adds a Start-menu group with shortcuts to extCheck, the README, and the uninstaller
- Adds no right-click Explorer verbs and no file associations
- Is 64-bit only

The runtime distribution from the installer's perspective is minimal: `extCheck.exe` (with embedded icon) plus `extCheck.csv`. Everything else in the install directory (README, license, source, etc.) is documentation and reference material.

### Running from source

```cmd
csc /platform:x64 extCheck.cs
extCheck.exe report.docx
```

### Uninstalling

Use Apps & Features in Windows Settings, or run the uninstaller from the extCheck Start-menu group. The uninstaller removes the program files. It does not touch `%LOCALAPPDATA%\extCheck\extCheck.ini` or any `extCheck.log` files in working directories — delete those manually if you want a fully clean removal.

---

## License

MIT License. See `license.htm`.
