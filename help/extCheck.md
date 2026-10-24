---
title: "extCheck Guide"
author: "Jamal Mazrui"
description: "Accessibility Checker for Office and Markdown Files"
---

# extCheck Guide

This is the full guide. For a quick start, see ReadMe.

**Author:** Jamal Mazrui
**Copyright:** © 2026 Jamal Mazrui
**License:** [MIT](https://opensource.org/license/mit/)
**Project home:** <https://github.com/JamalMazrui/extCheck>

`extCheck` is a Windows tool that checks Microsoft Word, Excel, PowerPoint, and Pandoc Markdown files for accessibility problems. For each file you give it, extCheck writes a CSV report listing issues with rule IDs, locations, problem descriptions, and remediation guidance.

Like its companion tools `urlCheck` and `2htm` (see the Announce file for a description of the family), `extCheck` runs in two modes: a **GUI mode** (a small parameter dialog launched by double-clicking the program, pressing its desktop hotkey, or running with `-g`) and a **command-line mode** (any other invocation, suitable for batch files and pipelines). Both modes accept the same options.

---

## What you need

- Windows 10 or later, 64-bit
- Microsoft Word, Excel or PowerPoint, to check `.docx`, `.xlsx` or `.pptx` files
- No Office at all for `.md` files

You do **not** need to install .NET yourself. `extCheck` uses .NET Framework 4.8.1, which comes with Windows 10 (since version 22H2) and Windows 11.

**A note on 32-bit and 64-bit.** `extCheck` is built as a 64-bit program. To work with Office, a program must match Office's bitness. Modern Office (Microsoft 365, Office 2019 and later, Office 2024) is 64-bit by default, so this fits the usual case. If your Office is 32-bit, `extCheck` gives a clear error that names the mismatch. You can then install 64-bit Office, or rebuild `extCheck` with `/platform:x86` (see Development below).

---

## Installing

Download `extCheck_setup.exe` from the [extCheck releases page](https://github.com/JamalMazrui/extCheck/releases) and run it. It needs administrator rights, because it installs for everyone on the computer. The setup wizard:

- Asks for the installation folder the first time (default: `C:\Program Files\extCheck`). An update goes to the same folder without asking.
- Shows a short MIT license summary on the welcome page; the full license is installed as `License.htm`.
- Adds a Start-menu group and a desktop shortcut whose hotkey is **Alt+Control+X**. Pressing **Alt+Control+X** from anywhere in Windows opens the `extCheck` dialog.
- Adds a **Report via extCheck** entry to the File Explorer right-click menu for all file types. Right-clicking any file (or pressing **Shift+F10**) and choosing this entry checks that file and writes the CSV report next to it. extCheck's supported formats are `.docx`, `.xlsx`, `.pptx`, and `.md`; on any other file it says so and stops. The access letter is **x**, matching the desktop hotkey, so from a file in File Explorer the keys are **Shift+F10**, then **x**.

The last page offers two checkboxes. **Launch extCheck now** is checked; **Open the user guide** is not. When you press Finish, a short results box says what was installed and where the logs are. extCheck starts after you close that box, so the box is never hidden behind it.

The program itself is in the `exec` folder of the installation, and this guide, the other documents and the rule registry are in the `help` folder. `ReadMe.htm`, `License.htm` and `extCheck.cmd`, which runs the program from a command prompt there, are at the top.

### Updating

In the dialog, press **F11**. extCheck asks GitHub whether a newer version exists. If one does, Yes is the default: press Enter and the new setup program is downloaded and started. If you already have the newest version, No is the default. If the check fails, for example with no internet connection, a message says so.

---

## Running extCheck

### From the dialog (easiest)

Launch `extCheck` from any of these:

- The desktop shortcut (or its **Alt+Control+X** hotkey)
- The Start-menu shortcut
- Double-clicking `extCheck.exe` in the `exec` folder in File Explorer

The dialog is the one every Homer launch app shows, in rows called bands that read left to right. Tab moves through its controls in this order. Each label has an underlined letter that you can press with **Alt** to jump straight to that control:

- **Source paths** [S] — a file path, a folder path, a wildcard pattern (such as `*.docx`), or several of these separated by spaces. A folder on its own means every supported file in it, the same as `<dir>\*.*`, so you need not add the wildcard. One path with spaces in it needs no quotes: when the whole field names a file or folder that exists, extCheck takes it as one path. You need quotes only when you give several, and one of them has a space.
- **Browse local source** [B] — pick a single source from a file picker
- **Output directory** [O] — where the output is written. Blank means the current working directory.
- **Choose output directory** [C] — pick the output directory from a folder picker
- **Force replacements** [F] — overwrite an existing `<basename>.csv` instead of skipping the input. Without this, extCheck skips an input whose CSV already exists in the output directory.
- **Use configuration** [U] — load these field values from the saved configuration at startup, and save them back when you press OK
- **View output directory** [V] — open the output directory in File Explorer when the run is done, after you close the Results box
- **Default settings** [D] — clear all fields, uncheck all boxes, and delete the saved configuration if any
- **Help** [H] — list every field with its tip, the dialog's keys, where the logs are kept, and the version check. F1 also shows Help.
- **OK** / **Cancel** — alone at the bottom right: start the run, or close without running. Control+Enter is OK and Escape is Cancel, from any control.

More keys work anywhere in the dialog. **F1** shows Help. **F11** checks the web, from any control, for a newer version of extCheck (see Updating). **Control+Enter** is OK from any control. **Shift+F1** says the tip for the field you are on, and **F7** lists the dialog's controls so you can move to one. The dialog is built with the Homer Lbc classes, so its text fields also have the Lbc editing keys, which Help lists.

**Default settings** returns you to the dialog with the fields cleared and any saved configuration deleted. Every run ends with the extCheck Results box. Its only button is OK, so Enter, Space or Escape closes it; then the output directory opens if View output directory is checked, and extCheck closes. A long report shows its first 15 lines and names a file with the rest.

The Browse local source and Choose output directory pickers open at the folder named in their text field, when it exists. Otherwise they open at your Documents folder. With **Use configuration** checked, those fields start with what you used last time, so the pickers pick up where you left off.

If you press OK with an output directory that does not yet exist, extCheck prompts to create it (default Yes). Choosing No keeps the dialog open with focus on the output field so you can correct it.

When all files have been processed, a final results dialog summarizes what was done.


### From the command line

Open a Command Prompt in the extCheck program folder, or put that folder on your PATH, and run `extCheck` with the source as an argument. `extCheck.cmd` there runs the program in the `exec` folder.

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

- `-d`, `--default-settings` — forget the saved settings and exit.
- `-f`, `--force-replacements` — overwrite an existing `<basename>.csv` instead of skipping the input.
- `-g`, `--gui-mode` — show the parameter dialog.
- `-h`, `--help` — show usage and exit.
- `-o <folder>`, `--output-directory <folder>` — write the reports to `<folder>` (made if missing); the default is the current directory.
- `-rules` — write the rule registry as `extCheck.csv` to the output directory and exit.
- `-u`, `--use-configuration` — read saved settings from `%LOCALAPPDATA%\extCheck\configs\extCheck.inix`.
- `--version` — show the version and exit.
- `-v`, `--view-output-directory` — after the run, open the output directory in File Explorer.

Every option in the dialog has a command-line switch named after its label, whose letter is the control's Alt letter, and switches may be written with /, - or --, in any case, so a workflow prototyped in the dialog can be translated to a batch file without surprises.

---

## Supported file formats

- `.docx` — Microsoft Word document
- `.xlsx` — Microsoft Excel workbook
- `.pptx` — Microsoft PowerPoint presentation
- `.md` — Pandoc Markdown file

`temp.*` (with a literal asterisk for the extension) expands to all supported extensions, so you can write `extCheck temp.*` to check `temp.docx`, `temp.xlsx`, `temp.pptx`, and `temp.md` if any exist.

---

## Output

For each file evaluated, a CSV named `<basename>.csv` is written to the output directory (`-o`, or the current working directory if no `-o`). The CSV columns are:

- **RuleID** — unique identifier for the rule (e.g., `MissingAltText`, `DuplicateHeadingText`)
- **Source** — always `extCheck`: every check is extCheck's own code
- **Modeled On** — the checker a rule follows: `Microsoft Accessibility Checker` or `axe-core`, or empty for a rule of extCheck's own. extCheck does not run either of them; it checks the same things its own way
- **Category** — high-level grouping (e.g., "Image", "Heading", "Link")
- **Location** — sheet name, slide label, line number, or `(Document)`
- **Context** — short snippet of the offending content
- **Message** — what the rule found and why it matters
- **Remediation** — how to fix it

**A place extCheck could not check is a finding of its own.** Sometimes Office cannot report a detail a rule needs, such as whether a chart has a title. extCheck then writes a `CouldNotInspect` row, in the category "Not checked", that names the detail, the rule that needed it and the error Office gave. It is not an accessibility problem in itself; it marks a place to check by hand. So a report with no problems means every check ran, and a report with `CouldNotInspect` rows says exactly where one did not.

**A chart's title and axis titles are where its check starts, not where it ends.** A chart with no axes, such as a pie or doughnut chart, is not checked for axis titles. And a titled chart still needs its figures, or its main point, in text: a table beside it, or alternative text that states the trend. That part is for a person to judge.

Results are also printed to the console. At the end of a multi-file run, a total count is printed.

## The rule registry

The installed `help` folder holds `extCheck.csv`, the rule registry. The program itself writes it when it is built, and running `extCheck -rules` writes the same file wherever you like. It lists every rule extCheck knows: the rule ID, the category of the Office Accessibility Checker it matches, the number of the WCAG 2.1 success criterion, the severity, the file formats it applies to, a description and advice on fixing it. Rules come from two sources that complement each other:

**MSAC** rules follow the categories of the Microsoft Office Accessibility Checker built into Office: missing alternative text, missing table headers, heading issues, repeated blank characters, blank cells used for formatting, merged cells, complex tables, color used alone, object titles, how lists are used, and text contrast.

**AXE** rules are adapted from axe-core, the open-source accessibility engine kept by Deque Systems. They cover areas the Office checker does not. These include whether links can be told apart, repeated heading text, link text that does not say where it goes, labels on form fields, slide reading order, animation timing, and the language named on a block of code in Markdown.

---

## Configuration file

When **Use configuration** is checked in the dialog (or `-u` is on the command line), `extCheck` reads and writes its settings in:

```
%LOCALAPPDATA%\extCheck\configs\extCheck.inix
```

It stores the source field, the output directory, and the option checkboxes, one per line under a `[Settings]` heading, such as `ViewOutput=yes`. You can edit it in any text editor; a comment you add is kept when extCheck saves. **Default settings** in the dialog deletes the file.

Settings saved by versions before 2.1, in `%LOCALAPPDATA%\extCheck\extCheck.ini`, are read when there is no `.inix` yet. The old file is removed the first time the new one is saved.

---

## Log files

extCheck always keeps a log of each session, whatever the options:

```
%LOCALAPPDATA%\extCheck\logs\extCheck-yyyyMMdd-HHmmss.log
```

Each run gets its own file, named for the moment it started, so sorting the folder by name also sorts it by time. The 30 newest are kept. The log starts with the version, where the program ran from, the Windows version, the working folder, the command line and every setting, then records each file checked and any error with its full details. If something goes wrong, zip the `logs` folder and send it.

Logging is automatic: there is no checkbox or switch for it, and Help names the folder. The log is UTF-8 with a byte-order mark, so Notepad opens it correctly.

---

## Notes

- extCheck reports automatically-detectable violations only. It does not replace manual testing.
- **PowerPoint** requires a visible application window; headless mode is not supported. The window is minimized automatically when extCheck launches PowerPoint, and closed when checking is complete.
- **Markdown** checking requires no Office software. The checker reads the file directly and evaluates Pandoc-flavored Markdown conventions.
- **False positives** are possible. For example, the empty-alt-text rule for Markdown flags all empty `[]` alt attributes, but empty alt is correct for purely decorative images. The all-caps rule ignores sequences of six characters or fewer to allow common acronyms. Review each flagged item in context before remediating.

---

## Uninstalling

Use Installed apps in Windows Settings, or the Uninstall shortcut in the extCheck Start-menu group. The uninstaller removes the program, its logs and the File Explorer entry. Your saved settings, `configs\extCheck.inix`, are left in place, as they always have been; delete them by hand if you want extCheck gone completely.

---

## For developers

How extCheck is built, released and structured is in Developer, in this folder.

## License

MIT License. See `License.htm` at the top of the program folder.
