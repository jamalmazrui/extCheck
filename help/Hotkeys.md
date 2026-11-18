---
title: "extCheck Hotkeys"
author: "Jamal Mazrui"
---

# extCheck Hotkeys

The keys of the extCheck dialog, which checks Word, Excel, PowerPoint and Markdown files for accessibility. The keys are given three ways: by command name, by key, and in the order Tab reaches the controls. Each Alt key is the letter underlined in its control's label, so a screen reader says it with the control. OK and Cancel have no letter: Control+Enter is OK and Escape is Cancel, from any control. Shift+F1, Control+Slash, F7 and F11 work in every Homer dialog, the messages extCheck shows included.

## By Command Name

Commands in alphabetical order.

- **Browse local source**, Alt+B -- Pick one or more source files in a standard dialog; it opens where Source paths points, else in your Documents folder.
- **Cancel**, Escape -- Close the dialog without starting, from any control.
- **Choose output directory**, Alt+C -- Pick the output directory in a standard folder dialog.
- **Default settings**, Alt+D -- Put every field back as it was out of the box, and forget any saved configuration.
- **Elevate Version**, F11 -- Download and install the latest extCheck.
- **Force replacements**, Alt+F -- Do the work again for a source whose result is already there, replacing the old result once the new one is complete.
- **Help**, Alt+H or F1 -- List each field with its tip, the dialog keys, where the logs are kept, and the version check.
- **List Controls**, F7 -- List the dialog's controls, and move to the one chosen.
- **OK**, Control+Enter or Enter -- Start the check: Control+Enter from any control, Enter from any but a button.
- **Open extCheck**, Alt+Control+X -- Open the extCheck dialog from anywhere in Windows, with its desktop shortcut.
- **Output directory**, Alt+O -- Type where each report is written, named after its file; blank means the current folder.
- **Report via extCheck**, Shift+F10, then X -- Check a file from its context menu in File Explorer.
- **Source paths**, Alt+S -- Type the files, folders or wildcard patterns such as *.docx to check, separated by spaces; a folder means every supported file in it.
- **Use configuration**, Alt+U -- Load the saved settings at startup, and save them when you press OK.
- **View output directory**, Alt+V -- Open the output directory in File Explorer when the run is done.
- **What This Does**, Shift+F1 -- Say what the control with focus does.
- **Where Am I**, Control+Slash -- Say where you are, what has focus and its keys.

## By Key

Keys in order of how many are pressed together, then alphabetically.

- **Enter** -- OK: Start the check: Control+Enter from any control, Enter from any but a button.
- **Escape** -- Cancel: Close the dialog without starting, from any control.
- **F1** -- Help: List each field with its tip, the dialog keys, where the logs are kept, and the version check.
- **F7** -- List Controls: List the dialog's controls, and move to the one chosen.
- **F11** -- Elevate Version: Download and install the latest extCheck.
- **Alt+B** -- Browse local source: Pick one or more source files in a standard dialog; it opens where Source paths points, else in your Documents folder.
- **Alt+C** -- Choose output directory: Pick the output directory in a standard folder dialog.
- **Alt+D** -- Default settings: Put every field back as it was out of the box, and forget any saved configuration.
- **Alt+F** -- Force replacements: Do the work again for a source whose result is already there, replacing the old result once the new one is complete.
- **Alt+H** -- Help: List each field with its tip, the dialog keys, where the logs are kept, and the version check.
- **Alt+O** -- Output directory: Type where each report is written, named after its file; blank means the current folder.
- **Alt+S** -- Source paths: Type the files, folders or wildcard patterns such as *.docx to check, separated by spaces; a folder means every supported file in it.
- **Alt+U** -- Use configuration: Load the saved settings at startup, and save them when you press OK.
- **Alt+V** -- View output directory: Open the output directory in File Explorer when the run is done.
- **Control+Enter** -- OK: Start the check: Control+Enter from any control, Enter from any but a button.
- **Control+Slash** -- Where Am I: Say where you are, what has focus and its keys.
- **Shift+F1** -- What This Does: Say what the control with focus does.
- **Shift+F10, then X** -- Report via extCheck: Check a file from its context menu in File Explorer.
- **Alt+Control+X** -- Open extCheck: Open the extCheck dialog from anywhere in Windows, with its desktop shortcut.

## In Tab Order

The dialog is in bands, rows read left to right, in this order:

1. Source paths -- Alt+S
2. Browse local source -- Alt+B
3. Output directory -- Alt+O
4. Choose output directory -- Alt+C
5. Force replacements -- Alt+F
6. Use configuration -- Alt+U
7. View output directory -- Alt+V
8. Default settings -- Alt+D
9. Help -- Alt+H or F1
10. OK and Cancel, alone at the bottom right: Control+Enter is OK and Escape is Cancel, from any control

## Text Box Keys

Every text box in the dialog is a Homer text box, with these keys too. A dialog's own letter comes first: where a control's underlined letter is the same as one of these, as Alt+C is Choose output directory, Alt and that letter goes to the control.

- **Control+A** -- Select All: Select all the text in the box.
- **Control+Shift+A** -- Unselect All: Clear the selection, keeping the cursor where it is.
- **Control+C** -- Copy: Copy the selection, or the current line when nothing is selected.
- **Control+X** -- Cut: Cut the selection, or the current line when nothing is selected.
- **Alt+X** -- Cut Append: Cut the selection or line and add it to what the clipboard holds, unless a control in the dialog has X as its letter.
- **Alt+C** -- Copy Append: Copy the selection or line and add it to what the clipboard holds, unless a control in the dialog has C as its letter.
- **Control+D** -- Delete Line: Delete the current line, without the clipboard.
- **F8** -- Start Selection: Mark where a selection starts.
- **Shift+F8** -- Complete Selection: Select from the start mark to the cursor.
- **Control+F8** -- Copy All: Copy all the text in the box.
- **Alt+F8** -- Read All: Read all the text in the box aloud.
- **Alt+Y** -- Say Yield: Say how many lines and characters the box holds, unless a control in the dialog has Y as its letter.
- **Alt+Apostrophe** -- Say Clipboard: Say what the clipboard holds.

## Why These Keys

F1 is help in every Windows program. F11 is the Homer key for updating a program to its newest release, because "elevate" sounds like "eleven". Alt+Control+X belongs to the desktop shortcut; Alt+Control keys are for desktop shortcuts. X is the sound that starts extCheck's name, and the same letter is the access key of Report via e**x**tCheck in File Explorer, so one letter reaches extCheck both ways.
