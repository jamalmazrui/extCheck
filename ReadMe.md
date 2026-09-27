---
title: "extCheck ReadMe"
author: "Jamal Mazrui"
---

# extCheck ReadMe

extCheck checks Microsoft Word, Excel, PowerPoint and Markdown files for accessibility problems. For each file, it writes a CSV report listing each problem, where it is, and how to fix it.

This is the quick start. The full guide is `help\extCheck.htm`.

## Install

1. Download `extCheck_setup.exe` from the [extCheck releases page](https://github.com/JamalMazrui/extCheck/releases).
2. Run it. It asks for administrator rights, because it installs for everyone on the computer.
3. On the last page, leave **Launch extCheck now** checked and press Finish. Read the results box, then close it; extCheck opens.

You need Windows 10 or 11, 64-bit, and the Office program for each kind of file you check: Word for `.docx`, Excel for `.xlsx`, PowerPoint for `.pptx`. Markdown needs nothing more.

## Check a file

1. Press **Alt+Control+X** from anywhere in Windows. The extCheck dialog opens with focus in **Source files**.
2. Type a file name, a folder, or a pattern such as `*.docx`.
3. Press Enter.

A results box says what was checked. Each report is named after its file, such as `report.csv`, in the output directory.

Or, in File Explorer, press **Shift+F10** on a file and then **x**, for Report via extCheck. The report is written next to the file.

## From the command line

In a Command Prompt in the program folder:

```cmd
extCheck report.docx
extCheck *.docx *.md -o reports
extCheck -h
```

## Keys in the dialog

- **Alt** with an underlined letter moves to that control.
- **Enter** starts the check; **Escape** closes the dialog.
- **F1** shows Help.
- **F11** checks the web for a newer version of extCheck and offers to install it.

All the keys are listed in `help\Hotkeys.htm`.

## When something goes wrong

Every run keeps a log in `%LOCALAPPDATA%\extCheck\logs`, one file per run. Zip that folder and send it with a description of what happened.

## License

extCheck is free and open source under the MIT License. See `License.htm`.
