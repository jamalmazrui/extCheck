---
title: "extCheck — Accessibility Checker for Office and Markdown Files"
author: "Jamal Mazrui"
date: "April 2, 2026"
lang: en-US
description: "User guide for extCheck.exe, a command-line accessibility checker for .docx, .xlsx, .pptx, and .md files."
---

# extCheck

A command-line accessibility checker for Microsoft Word (.docx), Excel (.xlsx), PowerPoint (.pptx), and Pandoc Markdown (.md) files. It evaluates each file against accessibility rules drawn from the Microsoft Office Accessibility Checker and the axe-core WCAG 2.1 rule set, then writes results to a CSV file and prints a summary to the console.

## Requirements

- Windows with .NET Framework 4.8.1 (included in Windows 10 and 11)
- Microsoft Word, Excel, or PowerPoint installed to check the corresponding file types
- No Office installation needed to check Markdown files

## Quick Start

Open a Command Prompt, navigate to the folder containing the files you want to check, and run:

```
extCheck.exe report.docx
```

Results appear in the console and a CSV file named `report.csv` is written to the current directory.

## Usage

```
extCheck.exe [-h] [-rules] <filespec> [<filespec> ...]
```

### Arguments

- `<filespec>` — Path to one or more files. Wildcards are supported.

### Options

- `-h` — Show help and exit.
- `-rules` — Write `extCheck.csv` to the current directory. This file is the complete rule registry, listing every rule ID, its WCAG criterion, severity, applicable formats, description, and remediation guidance.

### Supported Formats

| Extension | Format |
|-----------|--------|
| .docx | Microsoft Word document |
| .xlsx | Microsoft Excel workbook |
| .pptx | Microsoft PowerPoint presentation |
| .md | Pandoc Markdown file |

## Examples

Check a single Word document:

```
extCheck.exe report.docx
```

Check all Markdown files in the current directory:

```
extCheck.exe *.md
```

Check all Excel files in a subdirectory:

```
extCheck.exe quarterly\*.xlsx
```

Check multiple formats in one run:

```
extCheck.exe docs\*.docx slides\*.pptx notes\*.md
```

Export the complete rule list:

```
extCheck.exe -rules
```

## Output

For each file checked, a CSV file named `<filename>.csv` is written to the current working directory. If you check a file in another folder, the CSV still lands where you ran the command. If two files share the same base name but different extensions (for example `report.docx` and `report.md`), the CSV from the second one will overwrite the first, so check them in separate runs if you need both.

### CSV Columns

| Column | Contents |
|--------|----------|
| RuleID | Short identifier for the accessibility rule, e.g. `MissingAltText` |
| Source | `MSAC` (Office Accessibility Checker) or `AXE` (axe-core WCAG rules) |
| Category | The MS Office Accessibility Checker category name |
| Location | Where in the file the issue was found (sheet name, slide number, line number, or `(Document)`) |
| Context | The specific object or text where the issue occurs |
| Message | Plain-English description of the problem |
| Remediation | Step-by-step guidance for fixing it |

### Console Output

The console shows the same information in a readable format. Each issue is printed as:

```
[RuleID] (Source) Location | Category | Context
  Problem:   ...
  Remediate: ...
```

At the end of a multi-file run, a total count is printed.

## The Rule Registry

Running `extCheck.exe -rules` writes `extCheck.csv` with all 73 accessibility rules. Each row includes:

- **RuleID** — Unique identifier
- **MSOfficeCategory** — The Office Accessibility Checker category this maps to
- **WCAGCriteria** — The WCAG 2.1 success criterion number
- **Severity** — `Error` (definite barrier) or `Warning` (likely barrier, context-dependent)
- **AppliesTo** — Which file formats the rule checks
- **Description** — What the rule looks for and why it matters
- **Remediation** — How to fix it

## Rule Sources

Rules come from two complementary sources:

**MSAC** rules mirror the categories used by the built-in Microsoft Office Accessibility Checker: Missing alternative text, Missing table header, Heading issues, Repeated blank characters, Blank cells used for formatting, Merged cells, Complex table, Use of color alone, Object has no title or name, Object not inline, List not used correctly, and Text.

**AXE** rules are adapted from the axe-core open-source accessibility engine maintained by Deque Systems. These cover areas the Office checker does not address, including hyperlink distinguishability (`HyperlinkNotUnderlined`), duplicate heading text (`DuplicateHeadingText`), non-descriptive link text (`NonDescriptiveLinkText`), color used as the only indicator (`ColorUsedAlone`), form field labels (`UnlabeledFormField`), slide reading order (`TitleNotFirstInReadingOrder`), animation timing (`RepeatingAnimation`, `RapidAutoAnimation`, `FastAutoAdvance`), and code block language specifiers for Markdown (`CodeBlockMissingLanguage`).

## Notes on Specific Formats

**PowerPoint** requires a visible application window; headless mode is not supported. The window is minimized automatically when extCheck launches PowerPoint, and closed when checking is complete.

**Markdown** checking requires no Office software. The checker reads the file directly and evaluates Pandoc-flavored Markdown conventions including YAML front matter, ATX and setext headings, pipe tables, fenced code blocks, and raw HTML elements.

**False positives** are possible. The empty alt text rule for Markdown flags all empty `[]` alt attributes, but empty alt is correct for purely decorative images. The all-caps rule ignores sequences of six characters or fewer to allow common acronyms. Review each flagged item in context before remediating.

## Compiling from Source

The program is distributed as a compiled executable. If you have the source file `extCheck.cs` and want to recompile:

```
csc extCheck.cs /platform:x86
```

The `/platform:x86` flag is required because Office COM automation is 32-bit. The compiler (`csc.exe`) is included with .NET Framework and is typically found in `C:\Windows\Microsoft.NET\Framework\v4.0.30319\`.

## Contact

Jamal Mazrui  
Chief Accessibility Officer, CurbCutOS  
Seattle, Washington
