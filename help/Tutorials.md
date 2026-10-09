# Tutorials

## Contents

- [1 - User Interface](#1-user-interface)
- [2 - Install and Launch](#2-install-and-launch)
- [3 - Check a Word Document](#3-check-a-word-document)
- [4 - Check a PowerPoint Presentation](#4-check-a-powerpoint-presentation)
- [5 - Check an Excel Workbook](#5-check-an-excel-workbook)
- [6 - Check a Markdown File](#6-check-a-markdown-file)
- [7 - Check a Whole Folder](#7-check-a-whole-folder)
- [8 - Check from Explorer and the Command Line](#8-check-from-explorer-and-the-command-line)
- [9 - Conclusion](#9-conclusion)
- [1 - User Interface](#1-user-interface)
- [2 - Install and Launch](#2-install-and-launch)
- [3 - Check a Word Document](#3-check-a-word-document)
- [4 - Check a PowerPoint Presentation](#4-check-a-powerpoint-presentation)
- [5 - Check an Excel Workbook](#5-check-an-excel-workbook)
- [6 - Check a Markdown File](#6-check-a-markdown-file)
- [7 - Check a Whole Folder](#7-check-a-whole-folder)
- [8 - Check from Explorer and the Command Line](#8-check-from-explorer-and-the-command-line)
- [9 - Conclusion](#9-conclusion)
- [1 - User Interface](#1-user-interface)
- [2 - Install and Launch](#2-install-and-launch)
- [3 - Check a Word Document](#3-check-a-word-document)
- [4 - Check a PowerPoint Presentation](#4-check-a-powerpoint-presentation)
- [5 - Check an Excel Workbook](#5-check-an-excel-workbook)
- [6 - Check a Markdown File](#6-check-a-markdown-file)
- [7 - Check a Whole Folder](#7-check-a-whole-folder)
- [8 - Check from Explorer and the Command Line](#8-check-from-explorer-and-the-command-line)
- [9 - Conclusion](#9-conclusion)
- [0 - Overview](#0-overview)
- [1 - User Interface](#1-user-interface)
- [2 - Install and Launch](#2-install-and-launch)
- [3 - Check a Word Document](#3-check-a-word-document)
- [4 - Check a PowerPoint Presentation](#4-check-a-powerpoint-presentation)
- [5 - Check an Excel Workbook](#5-check-an-excel-workbook)
- [6 - Check a Markdown File](#6-check-a-markdown-file)
- [7 - Check a Whole Folder](#7-check-a-whole-folder)
- [8 - Check from Explorer and the Command Line](#8-check-from-explorer-and-the-command-line)
- [9 - Conclusion](#9-conclusion)
- [1. Where to Go Next](#1-where-to-go-next)

<!-- walkthrough: written by makeTutorials.py, do not edit between the markers -->

## 0 - Overview

What extCheck is, the two ways to run it, and the ten walks that teach it.

**Before you start:** Nothing to set up for this walk; it only listens.

### Step 1

Welcome. This is the first of ten short walks through extCheck. I am the host, and the other voice you hear is the screen reader, speaking as it would on your own computer.

### Step 2: Insert+Up Arrow

Two screen reader keys help in every walk. If you miss something the reader said, Insert plus Up Arrow says the current line again.

Screen reader:

- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Insert+Tab

And whenever you are unsure where you are, Insert plus Tab says what has the focus: its name, its kind, and its value.

Screen reader:

- Source files edit, blank

Confirm the wording with a live run: buildTutorials -live.

### Step 4

extCheck checks documents for accessibility problems before you share them: Word documents, Excel workbooks, PowerPoint presentations, and Markdown files. For each one it writes a report, a spreadsheet file listing every problem it found, where it is, why it matters, and how to fix it.

### Step 5

Word, Excel and PowerPoint files are opened by Microsoft Office in the background, so Office must be installed to check them. Markdown files need nothing else at all.

### Step 6

There are two ways to run it. The dialog is a small window of fields and check boxes, opened from the desktop with one key. The command line runs the same checks from a command prompt or a batch file, which suits checking many files on a schedule.

### Step 7

A third way sits in between: in File Explorer, the context menu of any file holds Report via extCheck, which checks that one file and writes its report beside it.

### Step 8

The reports are plain CSV files. Open one in Excel, or in any editor, and each row is one problem, with seven columns: the rule, its source, the category, the location, the text concerned, the message, and the fix.

### Step 9

Each rule comes from one of two families. Some mirror the categories of the Accessibility Checker built into Microsoft Office. Others are adapted from axe-core, the open-source engine many web checkers use, and cover things the Office checker does not.

### Step 10

Here are the ten walks. Walk one: the user interface, the dialog and its keys. Walk two: install and launch. Walk three: check a Word document. Walk four: check a PowerPoint presentation. Walk five: check an Excel workbook.

### Step 11

Walk six: check a Markdown file. Walk seven: check a whole folder and keep the reports together. Walk eight: check from File Explorer and the command line. Walk nine: the conclusion, a glossary, and where to get help.

### Step 12

Each walk is a few minutes long, and each builds on the ones before it, so they are best heard in order the first time. Walk one comes next.

**Something to try:** Think of one document you share with other people. Walk three, four, five or six will check one like it.

## 1 - User Interface

The extCheck dialog: its fields, its underlined letters, and the keys that work anywhere in it, each shown as the screen reader speaks it.

**Before you start:** extCheck installed, as walk two shows; the dialog open with Alt+Control+X.

### Step 1

One want: to move around the extCheck dialog quickly, without hunting. Everything in it can be reached by a letter.

### Step 2: Alt+Control+X

I press the desktop key for extCheck. The dialog opens on its first field.

Screen reader:

- extCheck dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3

Source files is where the documents to check are named: one file, a folder, a pattern such as star dot docx, or several of these with spaces between. The underlined letter is S.

### Step 4: Tab

Tab moves to the next control, the button that opens a file picker.

Screen reader:

- Browse source button

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Alt+O

Every label has an underlined letter, and Alt with that letter jumps straight to it from anywhere in the dialog. Output directory is O.

Screen reader:

- Output directory edit

Confirm the wording with a live run: buildTutorials -live.

### Step 6

Output directory is where the reports go. Left blank, they go to the current folder. Choose output, C, opens a folder picker instead of typing.

### Step 7: Alt+F

Next come four check boxes. Force replacements, F, replaces a report that already exists; without it, a document whose report is there is skipped.

Screen reader:

- Force replacements check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Space

Space checks it, and Space again clears it.

Screen reader:

- checked

Confirm the wording with a live run: buildTutorials -live.

### Step 9: Space

I clear it again, since most runs do not need it.

Screen reader:

- not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 10

View output, V, opens the reports' folder in File Explorer when the run is done. Log session, L, writes a log of the run beside the reports. Use configuration, U, remembers these fields for next time.

### Step 11: Shift+F1

Not sure what a field is for? Shift+F1 says the tip for the one you are on.

Screen reader:

- Force replacements: replace an existing report instead of skipping its document

Confirm the wording with a live run: buildTutorials -live.

### Step 12: F1

F1 shows Help: every field with its tip, and every key.

Screen reader:

- extCheck Help dialog

Confirm the wording with a live run: buildTutorials -live.

### Step 13: Escape

Escape closes Help and returns to the same field.

Screen reader:

- Force replacements check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 14: F7

F7 lists every control in the dialog. Arrow to one, press Enter, and the focus moves there.

Screen reader:

- Controls list

Confirm the wording with a live run: buildTutorials -live.

### Step 15: Enter

I choose Source files from the list.

Screen reader:

- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 16: Tab

Enter starts the check. From any control, even a check box, Control+Enter does the same.

Screen reader:

- OK button

Confirm the wording with a live run: buildTutorials -live.

### Step 17: Alt+G

Two more buttons sit at the end. Guide, G, opens the full guide in the browser and then returns here.

Screen reader:

- Guide button

Confirm the wording with a live run: buildTutorials -live.

### Step 18: Alt+D

Default settings, D, clears every field and box, and forgets anything saved. It asks first, since it cannot be undone.

Screen reader:

- Default settings button

Confirm the wording with a live run: buildTutorials -live.

### Step 19: Tab

Escape, from anywhere in the dialog, closes it without checking anything, the same as the Cancel button.

Screen reader:

- Cancel button

Confirm the wording with a live run: buildTutorials -live.

### Step 20

The dialog never speaks over the screen reader. It says a field's name once, as the reader would, and adds speech only for what the reader cannot know, such as a run's progress.

### Step 21: Alt+U

And the dialog remembers nothing unless asked. With Use configuration ticked, OK saves the fields; next time, they are already filled in.

Screen reader:

- Use configuration check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 22

Every key you have heard is listed in the Hotkeys document in the help folder, by key, by command, and by the field it belongs to, so you can review them without the dialog open.

### Step 23: Shift+Tab

Tab and Shift with Tab move forward and back through every control, in the order they appear, so nothing can be missed by moving one step at a time.

Screen reader:

- Use configuration check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 24: Alt+S

An underlined letter works whether or not the screen reader announces it, since it is plain Windows: the Alt key with the letter that the label shows.

Screen reader:

- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 25

That is the whole dialog: a field for what to check, a field for where reports go, four boxes, and a key for everything. The next walk installs extCheck.

**Something to try:** Open the dialog, press F7, and choose a control from the list. Then press Shift+F1 there and hear its tip.

## 2 - Install and Launch

Installing extCheck for everyone on the computer, the desktop key that opens it, and the File Explorer entry that checks one file.

**Before you start:** extCheck_setup.exe downloaded from the extCheck releases page on GitHub.

### Step 1

One want: extCheck on this computer, ready from the desktop and from File Explorer, in one sitting.

### Step 2: Enter

I open the downloaded setup program. It installs for everyone on the computer, so Windows asks for administrator rights. I say yes.

Screen reader:

- Setup - extCheck dialog
- Welcome to the extCheck Setup Wizard

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Enter

The welcome page shows a short summary of the license. extCheck is free and open source, under the MIT license. Enter takes Next.

Screen reader:

- Select Destination Location

Confirm the wording with a live run: buildTutorials -live.

### Step 4: Enter

The first time, it asks for a folder. The default, in Program Files, is right. An update later goes to the same folder without asking.

Screen reader:

- Ready to Install

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Enter

Enter again, and it copies the program and its documents.

Screen reader:

- Installing

Confirm the wording with a live run: buildTutorials -live.

### Step 6

The last page offers two check boxes. Launch extCheck now is checked; open the user guide is not.

Screen reader:

- Launch extCheck now check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 7: Enter

I press Finish. A short box says what was installed and where the logs are.

Screen reader:

- extCheck setup complete

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Enter

When I close that box, extCheck starts, its dialog ready on Source files.

Screen reader:

- extCheck dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 9

Setup added three ways to start it. The first is a desktop shortcut whose key is Alt+Control+X, which opens the dialog from anywhere in Windows.

### Step 10

The second is the Start menu, where extCheck has its own group with the program, the guide and the uninstaller.

### Step 11: Escape

The third is in File Explorer. I close the dialog and go to a document there.

Screen reader:

- Desktop

Confirm the wording with a live run: buildTutorials -live.

### Step 12: Shift+F10

Shift+F10 opens the context menu for the file I am on.

Screen reader:

- Context menu

Confirm the wording with a live run: buildTutorials -live.

### Step 13: X

X is Report via extCheck. It checks this one file at once and writes its report in the same folder, with the same name ending in dot csv.

Screen reader:

- extCheck: 1 file checked

Confirm the wording with a live run: buildTutorials -live.

### Step 14

Keeping extCheck current takes one key. In the dialog, F11 asks GitHub whether a newer version exists. If one does, Yes is the default, and Enter downloads and starts the new setup.

### Step 15

If this copy is already the newest, No is the default, and Enter simply returns to the dialog.

### Step 16

Installing for everyone has a reason: the program sits in Program Files, out of reach of changes made by mistake, and each person's settings and logs stay in their own folders.

Screen reader:

- C colon backslash Program Files backslash extCheck

Confirm the wording with a live run: buildTutorials -live.

### Step 17

Those per-person folders are under your own application data: configs for the saved fields, and logs for a record of every session, kept whatever the dialog's Log session box says.

Screen reader:

- configs folder, logs folder

Confirm the wording with a live run: buildTutorials -live.

### Step 18

If setup ever stops partway, its own log says why, and running the setup again simply finishes the job.

### Step 19

To remove extCheck, use Installed apps in Windows Settings, or the uninstaller in its Start-menu group. Your reports are never touched, since they are wherever you saved them.

Screen reader:

- extCheck, Uninstall button

Confirm the wording with a live run: buildTutorials -live.

### Step 20

One more thing the installer adds: extCheck dot cmd at the top of the installation, which runs extCheck from any command prompt. Walk eight uses it.

Screen reader:

- extCheck dot cmd

Confirm the wording with a live run: buildTutorials -live.

### Step 21: Alt+Control+X

Back in the dialog, the title bar shows the version, so you can always tell which extCheck is installed.

Screen reader:

- extCheck 1.4 dialog

Confirm the wording with a live run: buildTutorials -live.

### Step 22: Windows

And the Start-menu group holds the guide as well, so help is there even before the dialog is opened.

Screen reader:

- Start
- extCheck folder

Confirm the wording with a live run: buildTutorials -live.

### Step 23

Installed, launched three ways, and kept current with F11. The next walk checks a real document.

**Something to try:** After installing, right-click a document in File Explorer and find Report via extCheck.

## 3 - Check a Word Document

Checking a Word handout before it is emailed, and reading its report: the most common use of extCheck.

**Before you start:** A Word document of your own, such as a handout or a letter, and Microsoft Word installed.

### Step 1

One want: a Word handout that everyone can read, checked before I email it to a group.

### Step 2: Alt+Control+X

I open the dialog. The focus is on Source files.

Screen reader:

- extCheck dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Alt+B

Rather than type the path, Browse source opens a file picker.

Screen reader:

- Open dialog
- File name edit

Confirm the wording with a live run: buildTutorials -live.

### Step 4: Enter

I type the start of the handout's name and press Enter. The picker closes, and its full path is in Source files.

Screen reader:

- Source files edit
- C colon backslash Users backslash Documents backslash Handout dot docx

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Alt+V

I leave Output directory blank, so the report lands in the current folder, and I tick View output, so that folder opens when the check is done.

Screen reader:

- View output check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 6: Enter

Enter starts the check. Word opens the document invisibly in the background, so nothing appears on screen.

Screen reader:

- Checking Handout dot docx

Confirm the wording with a live run: buildTutorials -live.

### Step 7

A few seconds later the results box says what was done.

Screen reader:

- 1 file checked, 6 issues
- OK button

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Enter

Six issues. I press Enter, and File Explorer opens on the folder, on the new report, Handout dot csv.

Screen reader:

- Handout dot csv

Confirm the wording with a live run: buildTutorials -live.

### Step 9: Enter

I open it in Excel. Each row is one problem, and the columns are rule, source, category, location, context, message, and remediation, which is the fix.

Screen reader:

- Handout dot csv - Excel
- Rule ID, A1

Confirm the wording with a live run: buildTutorials -live.

### Step 10: Down Arrow

Down Arrow moves to the first problem. Control with Right Arrow reads along its row.

Screen reader:

- Missing Alt Text, A2

Confirm the wording with a live run: buildTutorials -live.

### Step 11

Missing alt text: a picture with no description, so a screen reader can say only that a picture is there. The location says which picture, and the remediation says how to add its description in Word.

### Step 12: Down Arrow

The next row is a heading problem.

Screen reader:

- Heading Skipped Level, A3

Confirm the wording with a live run: buildTutorials -live.

### Step 13

A heading skipped a level, from heading one straight to heading three. Readers who move by headings use those levels as an outline, so a gap misleads them about the structure.

### Step 14

Others in a typical Word report: a table with no header row, a link whose text says click here rather than where it goes, and blank paragraphs used for spacing, each heard as an empty line.

### Step 15

Each finding is a fix I can make in Word, in a minute or two: add the description, correct the heading level, mark the first table row as its header.

### Step 16: Alt+F

After fixing, I check again. The old report is there, so I tick Force replacements, which replaces it rather than skipping the document.

Screen reader:

- Force replacements check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 17: Enter

Enter, and the new report has fewer rows. When the count reaches zero, the handout is ready to send.

Screen reader:

- 1 file checked, 0 issues

Confirm the wording with a live run: buildTutorials -live.

### Step 18: Alt+Down Arrow

A report can be long. In Excel, a filter on the Rule ID column shows one kind of problem at a time, which makes a long report manageable.

Screen reader:

- Filter, Rule ID

Confirm the wording with a live run: buildTutorials -live.

### Step 19

Or sort by Location, to fix the document from top to bottom in one pass.

Screen reader:

- Sort A to Z, Location

Confirm the wording with a live run: buildTutorials -live.

### Step 20

The Message column always says why a problem matters, not just what it is, so a fix is never a matter of guessing.

Screen reader:

- Message, A picture has no description, so a screen reader can only say that a picture is there

Confirm the wording with a live run: buildTutorials -live.

### Step 21

Some rules are certain, such as a picture with no description at all. Others are judgments, such as a description that may be too short to be useful; the message says which kind a finding is.

Screen reader:

- Alt Text Too Short, Picture 3

Confirm the wording with a live run: buildTutorials -live.

### Step 22

If Word is already open with other documents, extCheck uses its own hidden copy, and leaves yours alone.

### Step 23

One document, checked, fixed and checked again. The next walk does the same for a slide presentation, where reading order matters most.

**Something to try:** Fix one problem the report names, check the document again with Force replacements on, and confirm that row is gone.

## 4 - Check a PowerPoint Presentation

Checking slides before a talk, and the problems presentations have that documents do not: slide titles, reading order, and pictures without descriptions.

**Before you start:** A PowerPoint presentation of your own, and Microsoft PowerPoint installed.

### Step 1

One want: slides for a talk that the audience can follow afterwards with a screen reader, when I share the file.

### Step 2: Alt+Control+X

I open the dialog, and type the presentation's path in Source files. Typing works as well as browsing.

Screen reader:

- extCheck dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Enter

I press Enter. PowerPoint opens the slides in the background, and extCheck reads each one.

Screen reader:

- Checking Talk dot pptx

Confirm the wording with a live run: buildTutorials -live.

### Step 4

The results box gives the count.

Screen reader:

- 1 file checked, 9 issues

Confirm the wording with a live run: buildTutorials -live.

### Step 5

Nine issues, in Talk dot csv. Presentations have three kinds of problem that documents do not.

### Step 6

The first is the slide title. A screen reader user moves through a presentation by its titles, the way a sighted reader glances at the top of each slide. A slide with no title is a slide with no name.

### Step 7

In the report, its location names the slide by number, so I know exactly where to look.

Screen reader:

- Slide Missing Title, Slide 4

Confirm the wording with a live run: buildTutorials -live.

### Step 8

The second is reading order. A screen reader reads the things on a slide from back to front, in the order they were added, not top to bottom as they look. A title added last is read last.

### Step 9

extCheck reports a slide whose title is not first to be read. In PowerPoint, the Selection Pane lists a slide's objects in that order, and dragging the title to the bottom of that list makes it read first.

Screen reader:

- Title Not First In Reading Order, Slide 2

Confirm the wording with a live run: buildTutorials -live.

### Step 10

The third is pictures and charts with no description. A chart that shows a trend says nothing unless its description says what the trend is.

Screen reader:

- Missing Alt Text, Slide 6, Chart 1

Confirm the wording with a live run: buildTutorials -live.

### Step 11

Others in a typical presentation: slides that advance on a timer, which may move before a listener has finished, and text too small to read when the slides are shown.

### Step 12

Some findings are judgments rather than certainties, such as text that may be too small. The message says why each matters, and the fix is mine to decide.

### Step 13

I fix the titles in PowerPoint first, since they are the presentation's outline, then the reading order, then the descriptions.

### Step 14: Alt+F

Then I check again with Force replacements on, so the new report replaces the old.

Screen reader:

- Force replacements check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 15: Enter

Enter, and the count comes down. Presentations are the files most often shared without a check, and the ones that need it most.

Screen reader:

- 1 file checked, 1 issue

Confirm the wording with a live run: buildTutorials -live.

### Step 16: Tab

A slide's reading order can be checked by ear, too: in PowerPoint, Tab moves through a slide's objects in exactly the order a screen reader reads them.

Screen reader:

- Title 1, title placeholder

Confirm the wording with a live run: buildTutorials -live.

### Step 17

The first thing heard should be the title. Here it is, so this slide is in order.

### Step 18

Charts deserve the most care. A good description says what the chart shows, such as sales doubling over three years, not that it is a bar chart.

### Step 19

Pictures used only for decoration should be marked decorative in PowerPoint, so a screen reader skips them; extCheck does not report a picture marked that way.

### Step 20

Speaker notes are a good home for anything said aloud in the talk but not on a slide, so a reader of the shared file gets it too.

Screen reader:

- Notes, Slide 3

Confirm the wording with a live run: buildTutorials -live.

### Step 21

Before checking, close the presentation in PowerPoint, so the copy extCheck reads is the one you last saved.

Screen reader:

- Talk dot pptx, saved

Confirm the wording with a live run: buildTutorials -live.

### Step 22: F5

After fixing, a quick test by ear helps too: in PowerPoint, start the slide show, and listen to each slide's title as it arrives.

Screen reader:

- Slide 1, Quarterly Results

Confirm the wording with a live run: buildTutorials -live.

### Step 23

Titles, reading order and descriptions: the three that matter most in slides. The next walk checks a workbook.

**Something to try:** Open a presentation's Selection Pane in PowerPoint, and see the order a screen reader follows on one slide.

## 5 - Check an Excel Workbook

Checking a shared spreadsheet: headers, merged cells, blank cells used for spacing, and sheet names, the problems that make a workbook hard to navigate by screen reader.

**Before you start:** An Excel workbook of your own, such as a budget or a schedule, and Microsoft Excel installed.

### Step 1

One want: a budget spreadsheet the whole team can work in, including the colleague who uses a screen reader.

### Step 2: Alt+Control+X

I open the dialog and put the workbook's path in Source files.

Screen reader:

- extCheck dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Enter

Enter. Excel opens the workbook in the background, and extCheck reads every sheet.

Screen reader:

- Checking Budget dot xlsx

Confirm the wording with a live run: buildTutorials -live.

### Step 4

The results box gives the count.

Screen reader:

- 1 file checked, 7 issues

Confirm the wording with a live run: buildTutorials -live.

### Step 5

Seven issues, in Budget dot csv. A spreadsheet is read cell by cell, so its problems are about knowing where you are.

### Step 6

The location column names the sheet and, where it matters, the cell, so each finding can be found again in a moment.

Screen reader:

- Merged Cells, Sheet1, A1 to D1

Confirm the wording with a live run: buildTutorials -live.

### Step 7

Merged cells: one cell spread across several columns, often a title across the top. Moving through it, a screen reader can lose its place in the columns below.

### Step 8

Next, a table with no header row.

Screen reader:

- Missing Table Header, Sheet1

Confirm the wording with a live run: buildTutorials -live.

### Step 9

A header row names each column, so a screen reader can say Amount, not just a number, as it moves along a row. In Excel, formatting the range as a table with headers fixes this.

### Step 10

Then a sheet name.

Screen reader:

- Default Sheet Name, Sheet1

Confirm the wording with a live run: buildTutorials -live.

### Step 11

Sheet1 says nothing. A name such as Monthly Budget tells a screen reader user which sheet they are on, the moment they switch to it.

### Step 12

Others in a typical workbook: blank rows and columns used only for spacing, which are heard as empty cells; colour used alone to mean something, such as red for overdue; and hidden sheets someone may need.

### Step 13

Colour alone is a judgment the report flags for me to decide: a word beside the colour, such as Overdue, says the same thing to everyone.

### Step 14

I fix the names and headers first, since every reader relies on them, then the merged and blank cells.

### Step 15: Alt+F

Then I check again with Force replacements on.

Screen reader:

- Force replacements check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 16: Enter

Enter, and the count comes down.

Screen reader:

- 1 file checked, 2 issues

Confirm the wording with a live run: buildTutorials -live.

### Step 17: Control+Page Down

In Excel, a workbook's structure is heard by moving, not by looking. Control with Page Down moves to the next sheet, and the reader says its name.

Screen reader:

- Monthly Budget

Confirm the wording with a live run: buildTutorials -live.

### Step 18

A named sheet tells me at once where I am. That is why sheet names matter.

### Step 19

Frozen panes help as well: a frozen header row stays in view, so its names are always there to be read with a cell.

Screen reader:

- Amount, B2

Confirm the wording with a live run: buildTutorials -live.

### Step 20

Formulas are fine to keep. extCheck reads what a cell shows, so a total computed by a formula is checked like any other value.

Screen reader:

- Total, 4,250 dollars, formula

Confirm the wording with a live run: buildTutorials -live.

### Step 21

A workbook shared for others to fill in benefits most from a check, since every person who opens it meets the same structure.

Screen reader:

- Budget dot csv

Confirm the wording with a live run: buildTutorials -live.

### Step 22

Before checking, save and close the workbook, so the copy extCheck reads is the latest one.

Screen reader:

- Budget dot xlsx, saved

Confirm the wording with a live run: buildTutorials -live.

### Step 23: Right Arrow

After fixing the headers, moving along a row says each column's name before its value, which is what a header row is for.

Screen reader:

- Amount, B5, 1,200

Confirm the wording with a live run: buildTutorials -live.

### Step 24: Control+Home

A short description of the workbook on its first sheet, saying what each sheet holds, helps every reader find the right one first.

Screen reader:

- Monthly Budget, A1, About this workbook

Confirm the wording with a live run: buildTutorials -live.

### Step 25

Headers, names and plain cells: a workbook anyone can navigate. The next walk checks a Markdown file, which needs no Office at all.

**Something to try:** Give each sheet of a workbook a meaningful name, then check it again and see those rows disappear.

## 6 - Check a Markdown File

Checking a Markdown file, such as a project's ReadMe on GitHub: headings, lists, links and pictures, with no Office installed.

**Before you start:** A Markdown file of your own, such as a ReadMe. Office is not needed.

### Step 1

One want: a ReadMe for my project on GitHub that reads well by screen reader, since it is the first thing a visitor sees.

### Step 2

Markdown is plain text with simple marks: a number sign for a heading, a dash for a list item, square brackets for a link. extCheck reads it directly, so no Office program is needed.

### Step 3: Alt+Control+X

I open the dialog and put the file's path in Source files.

Screen reader:

- extCheck dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 4: Enter

Enter, and the check takes a moment, since nothing has to start in the background.

Screen reader:

- 1 file checked, 5 issues

Confirm the wording with a live run: buildTutorials -live.

### Step 5

Five issues, in ReadMe dot csv. The location here is a line number, so I can go straight to it in my editor.

Screen reader:

- Heading Skipped Level, Line 12

Confirm the wording with a live run: buildTutorials -live.

### Step 6

A heading skipped a level, as in a Word document. Two number signs, then four, leaves a gap in the outline a reader moves through.

### Step 7

Next, a link.

Screen reader:

- Non Descriptive Link Text, Line 20

Confirm the wording with a live run: buildTutorials -live.

### Step 8

The link text says click here. A screen reader user often lists a page's links on their own, and a list of click here, click here tells them nothing. The text should say where the link goes.

### Step 9

Then a picture.

Screen reader:

- Missing Alt Text, Line 31

Confirm the wording with a live run: buildTutorials -live.

### Step 10

An image whose square brackets are empty has no description. Whatever goes between them is what a screen reader says.

### Step 11

Then a list that is not a list.

Screen reader:

- Fake Numbered List, Line 40

Confirm the wording with a live run: buildTutorials -live.

### Step 12

Lines numbered with a number in brackets, or a number and a dash, look like a list but are plain paragraphs to Markdown, so a screen reader does not announce a list or its length. A number followed by a full stop makes a real one.

### Step 13

Others in a typical Markdown file: more than one top heading, bare web addresses read out character by character, and tables with no header row.

### Step 14

Because a Markdown check needs nothing installed, it suits running on every change, before each commit, as a habit.

### Step 15: Alt+F

I fix the file and check again with Force replacements on.

Screen reader:

- Force replacements check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 16: Enter

Enter, and the report comes back clean.

Screen reader:

- 1 file checked, 0 issues

Confirm the wording with a live run: buildTutorials -live.

### Step 17

Markdown is also how many sites and documentation tools are written, so these habits travel well beyond GitHub.

### Step 18

Code blocks are treated with care: text between the fences is left alone, since a command or a sample is meant to look exactly as it does.

Screen reader:

- Line 50, code block, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 19

Pandoc's own Markdown extensions are understood too, such as a title block at the top and attributes in curly braces.

### Step 20

A table in Markdown needs its header row and the line of dashes under it, or Pandoc and GitHub both read it as plain text.

Screen reader:

- Missing Table Header, Line 60

Confirm the wording with a live run: buildTutorials -live.

### Step 21

Since the report gives line numbers, an editor's Go to line command takes me straight to each one.

### Step 22

Pictures in Markdown take their description between the square brackets, just before the address in round brackets.

Screen reader:

- Line 31, image, Screenshot of the main dialog

Confirm the wording with a live run: buildTutorials -live.

### Step 23

And the heading at the very top, one number sign, is the document's title: there should be exactly one.

Screen reader:

- Heading level 1, Project Name

Confirm the wording with a live run: buildTutorials -live.

### Step 24

Headings, links, pictures and lists: the Markdown the web reads best. The next walk checks a whole folder at once.

**Something to try:** Check a ReadMe of your own, fix one heading or link it names, and check it again.

## 7 - Check a Whole Folder

Checking every document in a folder at once, keeping the reports together in their own folder, and keeping a log of the run.

**Before you start:** A folder holding several documents of different kinds, such as course materials.

### Step 1

One want: every handout, slide deck and spreadsheet for a course checked before term starts, with the reports kept together.

### Step 2: Alt+Control+X

I open the dialog. In Source files, a folder's path is enough: extCheck checks every Word, Excel, PowerPoint and Markdown file in it.

Screen reader:

- extCheck dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3

I type the folder's path.

Screen reader:

- C colon backslash Courses backslash Fall

Confirm the wording with a live run: buildTutorials -live.

### Step 4

Several sources work too, separated by spaces, and a pattern such as star dot pptx checks only the presentations.

### Step 5: Alt+O

I want the reports in their own folder, not mixed in with the documents, so I go to Output directory.

Screen reader:

- Output directory edit

Confirm the wording with a live run: buildTutorials -live.

### Step 6

Choose output opens a folder picker, but I simply type a new folder's path, Reports, beside the course folder.

Screen reader:

- C colon backslash Courses backslash Reports

Confirm the wording with a live run: buildTutorials -live.

### Step 7: Alt+L

I tick Log session, so a log of the whole run is written beside the reports, and View output, so their folder opens at the end.

Screen reader:

- Log session check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Enter

Enter. The Reports folder does not exist yet, so extCheck asks before making it. Yes is the default.

Screen reader:

- Create the output directory?
- Yes button

Confirm the wording with a live run: buildTutorials -live.

### Step 9: Enter

Enter again, and the run begins. A small window counts the files as they are checked, so a long run is never silent.

Screen reader:

- Checking 3 of 12: Week 3 dot pptx

Confirm the wording with a live run: buildTutorials -live.

### Step 10

The results box gives the totals for the whole folder.

Screen reader:

- 12 files checked, 48 issues

Confirm the wording with a live run: buildTutorials -live.

### Step 11

Twelve files, forty-eight issues, in twelve reports, one per document, each named for its document.

### Step 12

If two documents share a name, such as Notes dot docx and Notes dot md, the second report takes its kind into its name, Notes dash md dot csv, so neither replaces the other.

### Step 13

When I run the same check again tomorrow, documents whose reports are already there are skipped, so only new ones are checked. Force replacements checks them all again.

### Step 14: Alt+U

Use configuration, U, saves these fields, so tomorrow the dialog opens with the same folders already filled in.

Screen reader:

- Use configuration check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 15

The log is the record of the run: every file, every finding, and anything that went wrong, such as a document Office could not open.

### Step 16

If a document fails to open, the others are still checked, and the run says which failed and why.

### Step 17

Checking a folder also checks the folders inside it? Not by itself. Name each subfolder, or a pattern, to include them, so a run never wanders where you did not mean it to.

### Step 18

Office lock files, the ones whose names begin with a tilde, are skipped, since they are not documents.

Screen reader:

- Skipped tilde dollar Week 3 dot docx

Confirm the wording with a live run: buildTutorials -live.

### Step 19

The progress window can be left alone; the run carries on, and its count says how far it has come at any moment.

Screen reader:

- Checking 9 of 12: Budget dot xlsx

Confirm the wording with a live run: buildTutorials -live.

### Step 20

At the end, the results box names any file that could not be checked, with the reason, such as a document protected by a password.

Screen reader:

- 11 files checked, 1 failed

Confirm the wording with a live run: buildTutorials -live.

### Step 21

Opened together in Excel, the reports make a list of the work to do before term, document by document.

### Step 22

A report for each document means a fix to one never hides a problem in another: each report stands on its own.

Screen reader:

- Week 3 dot csv

Confirm the wording with a live run: buildTutorials -live.

### Step 23

The totals at the end make progress measurable: fewer issues each time the folder is checked.

Screen reader:

- 12 files checked, 19 issues

Confirm the wording with a live run: buildTutorials -live.

### Step 24

A whole term's materials, checked in one run, with one folder of reports. The last task walk does the same from File Explorer and the command line.

**Something to try:** Check a folder twice, the second time without Force replacements, and hear the documents already reported skipped.

## 8 - Check from Explorer and the Command Line

Checking without the dialog: from File Explorer's context menu, and from a command prompt or a batch file, with the same options as the dialog.

**Before you start:** A document in File Explorer, and a command prompt.

### Step 1

One want: to check a document the moment I am looking at it, and to check a folder every week without opening anything.

### Step 2

The quickest way is in File Explorer. I am on a document there.

Screen reader:

- Handout dot docx

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Shift+F10

Shift+F10 opens its context menu.

Screen reader:

- Context menu

Confirm the wording with a live run: buildTutorials -live.

### Step 4: X

X is Report via extCheck.

Screen reader:

- extCheck: 1 file checked, 2 issues

Confirm the wording with a live run: buildTutorials -live.

### Step 5

The report is written beside the document, so I find it in the same folder, with the same name ending in dot csv.

### Step 6

The command line suits work done often. In a command prompt, extCheck followed by a file's name checks that file.

Screen reader:

- extCheck Handout dot docx

Confirm the wording with a live run: buildTutorials -live.

### Step 7

Every check box in the dialog has a command-line option. Dash o followed by a folder sends the reports there; dash f replaces existing reports; dash l writes the session log.

### Step 8

So one line checks a whole folder into a Reports folder, replacing last week's reports.

Screen reader:

- extCheck dash f dash o Reports C colon backslash Courses backslash Fall

Confirm the wording with a live run: buildTutorials -live.

### Step 9

Saved in a batch file, that line runs in a moment, or on a schedule with Windows Task Scheduler.

### Step 10

Dash u reads the settings saved by the dialog's Use configuration box, so a workflow tried out in the dialog can be run the same way from a batch file.

### Step 11

When a run ends, extCheck returns an exit code a batch file can test: zero when every file was checked, and one when a file could not be checked or nothing was found to check.

### Step 12

Dash rules writes the rule registry, the list of every rule extCheck knows, with its source, category and fix, as a spreadsheet.

Screen reader:

- extCheck dash rules

Confirm the wording with a live run: buildTutorials -live.

### Step 13

The registry is the place to look when a report names a rule I have not met before. A copy is also in the help folder of the installation.

### Step 14

Dash h shows every option, and dash v the version.

Screen reader:

- extCheck dash h

Confirm the wording with a live run: buildTutorials -live.

### Step 15

Dash g opens the dialog from the command line, for anyone who starts there but prefers the fields.

### Step 16

The File Explorer entry works on any file, but extCheck checks only the four kinds it knows; for anything else, it says so and writes no report.

### Step 17

Patterns work on the command line as in the dialog: star dot docx checks every Word document in the current folder.

Screen reader:

- extCheck star dot docx

Confirm the wording with a live run: buildTutorials -live.

### Step 18

Quotation marks keep a path with spaces in one piece, such as a folder called Course Materials.

Screen reader:

- extCheck quote C colon backslash Course Materials quote

Confirm the wording with a live run: buildTutorials -live.

### Step 19

The session log is kept in your own application data, so a run from a schedule leaves a record even though nobody saw it.

### Step 20

And because the dialog and the command line use the same options, anything learned in one works in the other.

### Step 21

A batch file can also check the exit code, and say when a check needs attention.

Screen reader:

- if errorlevel 1 echo Some files could not be checked

Confirm the wording with a live run: buildTutorials -live.

### Step 22: F1

For anyone who works mostly in a command prompt, the dialog's Help, F1, lists every option too, side by side with its field.

Screen reader:

- extCheck Help dialog

Confirm the wording with a live run: buildTutorials -live.

### Step 23

From the menu, from a prompt, or from a schedule: the same checks, wherever the work is. The last walk gathers it all together.

**Something to try:** Write a one-line batch file that checks a folder into a Reports folder, and run it.

## 9 - Conclusion

What the walks covered, a glossary in two voices, and every way to get help with extCheck.

**Before you start:** Nothing to set up for this walk; it only listens.

### Step 1

That is extCheck. Here is what each walk gave you.

### Step 2

Walk one: the dialog, and a letter for everything in it. Walk two: installing, and the three ways to start. Walks three to six: a Word document, a presentation, a workbook and a Markdown file, each checked and fixed.

### Step 3

Walk seven: a whole folder, with the reports kept together. Walk eight: File Explorer and the command line.

Screen reader:

- 12 files checked, 48 issues

Confirm the wording with a live run: buildTutorials -live.

### Step 4

Now a short glossary. I say the term; the other voice says what it means.

### Step 5

Alternative text.

Screen reader:

- The description of a picture or chart that a screen reader speaks in its place.

### Step 6

Heading level.

Screen reader:

- A heading's place in a document's outline: one for the title, two for its sections, three for theirs.

### Step 7

Reading order.

Screen reader:

- The order a screen reader reads the things on a slide: back to front, in the order they were added.

### Step 8

Header row.

Screen reader:

- The first row of a table, naming each column, so each cell can be said with its column's name.

### Step 9

Rule registry.

Screen reader:

- The list of every rule extCheck checks, with its source, its category and its fix.

### Step 10

MSAC and AXE.

Screen reader:

- The two families of rules: the categories of Microsoft's own Accessibility Checker, and rules adapted from the axe-core engine.

### Step 11

Report.

Screen reader:

- The spreadsheet file extCheck writes for each document, one problem to a row.

### Step 12: F1

Help is never more than a key away. In the dialog, F1 shows Help: every field with its tip, and every key.

Screen reader:

- extCheck Help dialog

Confirm the wording with a live run: buildTutorials -live.

### Step 13: Shift+F1

Shift+F1 says the tip for the field you are on. F7 lists the dialog's controls and moves to the one you choose.

Screen reader:

- Source files: one file, a folder, a pattern, or several, separated by spaces

Confirm the wording with a live run: buildTutorials -live.

### Step 14

Guide, Alt+G, opens the full guide in your browser and then returns to the dialog as you left it. The guide, these walks as text, and the rule registry are all in the help folder of the installation.

### Step 15

F11 checks for a newer version, and the extCheck page on GitHub has the latest release and a place to report a problem.

### Step 16

Accessibility is a habit more than a task. A check before each document leaves your desk takes a minute, and saves every reader the trouble.

### Step 17

Thank you for listening. Happy checking.

**Something to try:** Pick the document you share most often, check it, and fix the first three rows of its report.

<!-- walkthrough ends -->

## 1. Where to Go Next

Press F1 for the guide, Alt+Shift+H for the hotkey list, and Alt+F10 for every
command in one window.
