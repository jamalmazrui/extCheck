---
title: "extCheck — Accessibility Checker for Office and Markdown Files"
author: "Jamal Mazrui"
description: "Accessibility Checker for Office and Markdown Files"
---

# extCheck

**Author:** Jamal Mazrui
**License:** MIT

`extCheck` is one of three companion accessibility tools by Jamal Mazrui:

- **2htm** — convert documents (Word, Excel, PowerPoint, PDF, Markdown) to accessible HTML
- **extCheck** — check Office and Markdown files for accessibility problems
- **urlCheck** — check web pages for accessibility problems

The three tools share a common command-line and GUI layout, so learning one makes the others easy to pick up.

`extCheck` is a Windows tool that checks Microsoft Word, Excel, PowerPoint, and Pandoc Markdown files for accessibility problems. For each file you give it, extCheck writes a CSV report listing issues with rule IDs, locations, problem descriptions, and remediation guidance.

Like its companion tools, `extCheck` runs in two modes: a **GUI mode** (a small parameter dialog launched by double-clicking the program, pressing its desktop hotkey, or running with `-g`) and a **command-line mode** (any other invocation, suitable for batch files and pipelines). Both modes accept the same options.

---

## What you need

- Windows 10 or later (64-bit)
- Microsoft Word, Excel, or PowerPoint installed to check the corresponding `.docx`, `.xlsx`, or `.pptx` files
- No Office installation needed for `.md` files

You do **not** need to install .NET separately. The .NET Framework 4.8.1 used by `extCheck` ships in-box with Windows 10 (since version 22H2) and Windows 11.

**Bitness note.** `extCheck` is built as a 64-bit program, and Microsoft Office automation requires the controller process and the installed Office to share the same bitness. Modern Office (Microsoft 365, Office 2019+, Office 2024) is 64-bit by default, so this matches the common case. If you have 32-bit Office on your machine, `extCheck` will surface a clear error pointing at the bitness mismatch; you can either install 64-bit Office or rebuild `extCheck` with `/platform:x86` (see Development below).

---

## Installing

Download `extCheck_setup.exe` from the [GitHub repository](https://github.com/JamalMazrui/extCheck) and run it. The setup wizard:

- Prompts you for the installation directory (default: `C:\Program Files\extCheck`).
- Includes a brief MIT license summary on the welcome page; the full license text is installed alongside the program as `License.htm`.
- Adds a Start-menu shortcut and a desktop shortcut whose hotkey is **Alt+Ctrl+X**. Pressing **Alt+Ctrl+X** from anywhere in Windows opens the `extCheck` dialog.

The final wizard page offers two checkboxes (both checked by default): launch `extCheck` (with a hotkey reminder) and read the HTML documentation.

---

## Running extCheck

### From the dialog (easiest)

Launch `extCheck` from any of these:

- The desktop shortcut (or its **Alt+Ctrl+X** hotkey)
- The Start-menu shortcut
- Double-clicking `extCheck.exe` in File Explorer
- A Run dialog (`Win+R`) typing `extCheck`

The parameter dialog has these controls. Each label has an underlined letter that you can press with **Alt** to jump straight to that control:

- **Source files** [S] — a single file path, a wildcard pattern (e.g., `*.docx`), or several of either separated by spaces. Wrap paths containing spaces in double quotes.
- **Browse source...** [B] — pick a single source from a file picker
- **Output directory** [O] — where the output is written. Blank means the current working directory.
- **Choose output...** [C] — pick the output directory from a folder picker
- **Force replacements** [F] — overwrite an existing `<basename>.csv` instead of skipping the input. Without this, extCheck skips an input whose CSV already exists in the output directory.
- **View output** [V] — open the output directory in File Explorer when the run is done
- **Log session** [L] — write a fresh `extCheck.log` in the output directory (or current directory if no output directory is set)
- **Use configuration** [U] — load these field values from the saved configuration at startup, and save them back when you press OK
- **Help** [H] — show this help summary and offer to open the full README. F1 also shows Help.
- **Default settings** [D] — clear all fields, uncheck all boxes, and delete the saved configuration if any
- **OK** / **Cancel** — start the run, or cancel without running. Enter is OK; Esc is Cancel.

The Browse source and Choose output pickers open at the directory derived from the corresponding text field's current value when that value points to an existing path; otherwise they open at your Documents folder. With **Use configuration** checked, those text fields are pre-populated from your last session, so the pickers naturally pick up where you left off.

When all files have been processed, a final results dialog summarizes what was done.


### From the command line

Open a Command Prompt and run `extCheck` with the source as an argument:

```cmd
# Check one file:
extCheck report.docx

# Several files at once:
extCheck *.docx *.md

# Files in different folders:
extCheck docs\*.docx data\*.xlsx slides\*.pptx

# Write reports to a specific directory:
extCheck *.md -o reports

# Show the rule registry:
extCheck -rules

# Open the GUI:
extCheck -g

```

When invoked without arguments from a GUI shell (Explorer double-click, Start-menu shortcut, desktop hotkey), `extCheck` shows the dialog automatically. When invoked without arguments from a console shell, it prints help and exits. The `-g` flag forces GUI mode regardless.

---

## Command-line options

| Option | Long form | Description |
|---|---|---|
| `-h` | `--help` | Show usage and exit |
| `-v` | `--version` | Show version and exit |
| `-g` | `--gui-mode` | Show the parameter dialog |
| `-o <d>` | `--output-dir <d>` | Write output to `<d>` (created if missing); defaults to current directory |
| `-f` | `--force` | overwrite an existing `<basename> |
|   | `--view-output` | After the run, open the output directory in File Explorer |
| `-l` | `--log` | Write `extCheck.log` (UTF-8 with BOM) in the output directory; replaced each session |
| `-u` | `--use-configuration` | Read saved defaults from `%LOCALAPPDATA%\extCheck\extCheck.ini` |
| `-rules` | | Write the rule registry as `extCheck.csv` to the output directory and exit |

Every option in the GUI corresponds one-to-one with a command-line flag, so a workflow prototyped in the dialog can be translated to a batch file without surprises.

---

## Supported file formats

| Extension | Format |
|---|---|
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

Results are also printed to the console. At the end of a multi-file run, a total count is printed.

## The rule registry

Running `extCheck -rules` writes `extCheck.csv` containing every rule extCheck knows about: rule ID, the Office Accessibility Checker category it maps to, the WCAG 2.1 success criterion number, severity, applicable file formats, description, and remediation guidance. Rules come from two complementary sources:

**MSAC** rules mirror the categories used by the built-in Microsoft Office Accessibility Checker: missing alternative text, missing table headers, heading issues, repeated blank characters, blank cells used for formatting, merged cells, complex tables, use of color alone, object titles, list usage, text contrast.

**AXE** rules are adapted from the axe-core open-source accessibility engine maintained by Deque Systems. These cover areas the Office checker does not address, including hyperlink distinguishability, duplicate heading text, non-descriptive link text, form field labels, slide reading order, animation timing, and code block language specifiers for Markdown.

---

## Configuration file

When **Use configuration** is checked in the dialog (or `-u` is on the command line), `extCheck` reads and writes a small INI file at:

```
%LOCALAPPDATA%\extCheck\extCheck.ini
```

It stores the source field, the output directory, and the option checkboxes. Without **Use configuration**, `extCheck` leaves nothing on disk between runs. **Default settings** in the dialog deletes this file.

---

## Log file

When **Log session** is checked (or `-l` is on the command line), `extCheck` writes a fresh `extCheck.log` to the output directory (or current directory if no output directory is set). Any prior log is replaced at the start of the run, so the file always reflects only the current session.

The log captures: program version, command-line arguments, GUI auto-detection, the resolved output directory, per-file events, and any errors (including tracebacks for unexpected failures).

Without **Log session**, `extCheck` does not create any log or error file on disk. Errors are reported only to the console (and the GUI results dialog, in GUI mode).

The log is UTF-8 with a byte-order mark, so Notepad opens it correctly.

---

## Notes

- extCheck reports automatically-detectable violations only. It does not replace manual testing.
- **PowerPoint** requires a visible application window; headless mode is not supported. The window is minimized automatically when extCheck launches PowerPoint, and closed when checking is complete.
- **Markdown** checking requires no Office software. The checker reads the file directly and evaluates Pandoc-flavored Markdown conventions.
- **False positives** are possible. For example, the empty-alt-text rule for Markdown flags all empty `[]` alt attributes, but empty alt is correct for purely decorative images. The all-caps rule ignores sequences of six characters or fewer to allow common acronyms. Review each flagged item in context before remediating.

---

## Development

This section is for developers who want to build the executable from source. End users can skip it.

### Distribution layout

The runtime distribution shipped by `extCheck_setup.exe` is just a few files: `extCheck.exe` plus the HTML documentation (`ReadMe.htm`, `Announce.htm`, `License.htm`). The Markdown sources, the build script, the installer script, the icon, the program source, and the coding-style guide live in the GitHub repository (and in this `extCheck.zip` archive).

### Source layout

The whole program is one C# file: `extCheck.cs`. It uses standard `System.Windows.Forms` for the parameter dialog and the COM `dynamic` keyword to drive Office. There are no third-party dependencies. The classes inside `extCheck.cs` are arranged as a shared infrastructure layer (`issue`, `results`, `shared`, `com`, `logger`, `configManager`, `guiDialog`) plus per-format modules (`docxModule`, `xlsxModule`, `pptxModule`, `mdModule`), with a top-level `program` class that parses arguments, optionally shows the dialog, and dispatches to the right module per file extension.

### Coding style

The source uses what the author calls "Camel Type" (C# variant): Hungarian prefix notation for variables (`b` for boolean, `i` for integer, `s` for string, `ls` for `List<T>`, `d` for dictionary, etc.), lower camelCase for everything other than where the language requires PascalCase. The `o` prefix is reserved for COM objects only; managed C# objects use the lowercase class name as their prefix (e.g., `Form form`, `OpenFileDialog dialog`). Constants follow the same naming as variables — only the `const` or `static readonly` keyword conveys constant-ness. See `Camel_Type_C#.md` in this archive for the full guidelines.

### Threading and bitness

`Main` is decorated with `[STAThread]`. This is required for two reasons:

- Office COM automation requires a single-threaded apartment. Without it, Word/Excel/PowerPoint COM servers can disconnect mid-operation with HRESULT 0x80010108 (RPC_E_DISCONNECTED) or 0x80010114 (OLE_E_OBJNOTCONNECTED).
- WinForms common dialogs (`OpenFileDialog`, `FolderBrowserDialog`) require an STA thread.

The build is `/platform:x64`. Office COM automation requires the controller process and the installed Office to share the same bitness. Modern Office is 64-bit by default; if a user has 32-bit Office, `com.createApp` surfaces a clear error message pointing at the mismatch and recommending a 32-bit rebuild.

### Prerequisites

- The .NET Framework 4.8.1 Developer Pack (provides `csc.exe` and the 4.8.1 reference assemblies). Install from <https://dotnet.microsoft.com/download/dotnet-framework/net481>.
- Inno Setup 6.x to compile the installer.

### Building the executable

Run the included script:

```cmd
buildExtCheck.cmd
```

It auto-detects the compiler, verifies the build environment, embeds the icon into `extCheck.exe`, and produces the runtime distribution in `dist\`.

### Building the installer

Open `extCheck.iss` in Inno Setup and click Compile. The result is `dist\extCheck_setup.exe`.

The installer ships only the runtime files: `extCheck.exe` plus the HTML documentation (`ReadMe.htm`, `Announce.htm`, `License.htm`), plus app-specific runtime files (`extCheck.csv` for the rule registry). Markdown sources, the build script, this `.iss` script, the icon, the source file, and any coding-style guideline files live in the GitHub repository.

### Uninstalling

Use Apps & Features in Windows Settings, or run the uninstaller from the `extCheck` Start-menu group. The uninstaller removes the program files. It does not touch `%LOCALAPPDATA%\extCheck\extCheck.ini` or any `extCheck.log` files in working directories — delete those manually if you want a fully clean removal.


## License

MIT License. See `License.htm` (installed alongside the program) or `License.md` (in the GitHub repository).