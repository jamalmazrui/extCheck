// extCheck.cs  — Unified accessibility checker for .docx .xlsx .pptx .md files
// Compile: csc extCheck.cs /platform:x64    (see buildExtCheck.cmd)
// Usage:   extCheck.exe [-h] [-g] [-rules] [-o <dir>] [--view-output]
//                       [-l] [-u] [-f] <filespec> [<filespec> ...]
//          <filespec> may include wildcards: *.docx  docs\*.md  C:\work\*.pptx
// Output:  <basename>.csv in the output directory (-o) or the current
//          working directory if no -o is given.
// Requires: .NET Framework 4.8.1.
//           Word / Excel / PowerPoint required for .docx / .xlsx / .pptx files.
//           No Office required for .md files.
//
// Bitness: extCheck is built 64-bit. Office COM automation requires
// matching bitness between this process and the installed Office. Modern
// Office (Microsoft 365 / Office 2019+ / Office 2024) is 64-bit by
// default. If a user has 32-bit Office, comHelper.createApp will surface a
// clear error pointing at the bitness mismatch.
//
// Architecture: Option D — static format modules + shared static infrastructure.
//   Shared layer:  issue, results, shared, com, logger, configManager, guiDialog
//   Format modules: docxModule, xlsxModule, pptxModule, mdModule
//   Dispatcher:    program.run() — resolves files, calls the right module per extension

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

// ===========================================================================
// Shared: issue record
// ===========================================================================
class issue {
    public string sRuleId, sSource, sCategory, sLocation, sContext, sMessage, sRemediation;

    public issue(string ruleId, string source, string category, string location,
                 string context, string message, string remediation) {
        sRuleId      = ruleId;
        sSource      = source;
        sCategory    = category;
        sLocation    = location;   // Sheet name, Slide label, line number, or "(Document)"
        sContext     = context;
        sMessage     = message;
        sRemediation = remediation;
    }
}

// ===========================================================================
// Shared: issue list, CSV output, console summary
// ===========================================================================
static class results {
    public static List<issue> lIssues = new List<issue>();
    const string sCsvHeader =
        "RuleID,Source,Category,Location,Context,Message,Remediation";

    public static void clear() {
        lIssues.Clear();
    }

    public static void add(
            string ruleId, string source, string category, string location,
            string context, string message, string remediation) {
        lIssues.Add(new issue(ruleId, source, category, location, context, message, remediation));
    }

    public static void writeCsv(string sPath) {
        var sb = new StringBuilder();
        sb.AppendLine(sCsvHeader);
        foreach (issue o in lIssues) {
            sb.AppendLine(
                esc(o.sRuleId)      + "," +
                esc(o.sSource)      + "," +
                esc(o.sCategory)    + "," +
                esc(o.sLocation)    + "," +
                esc(o.sContext)     + "," +
                esc(o.sMessage)     + "," +
                esc(o.sRemediation));
        }
        File.WriteAllText(sPath, sb.ToString(), Encoding.UTF8);
    }

    static string esc(string s) {
        s = s.Replace("\"", "\"\"");
        if (s.Contains(",") || s.Contains("\"") || s.Contains("\r") || s.Contains("\n")) {
            s = "\"" + s + "\"";
        }
        return s;
    }
}

// ===========================================================================
// Shared: vocabulary and utility functions used by multiple format modules
// ===========================================================================
static class shared {
    public static readonly string[] aVague = {
        "click here", "here", "link", "read more", "more", "this",
        "url", "learn more", "continue", "details", "info", "information", "go"
    };

    public static readonly string[] aBullets = { "\u2022", "\u25CF", "-", "*" };

    public static string trunc(string s) {
        return s.Length > 60 ? s.Substring(0, 60) + "..." : s;
    }

    public static bool isVagueLinkText(string sText) {
        string sLower = sText.ToLower();
        if (sText == "") {
            return true;
        }
        foreach (string v in aVague) {
            if (sLower == v || sLower.StartsWith("http") || sLower.StartsWith("www")) {
                return true;
            }
        }
        return false;
    }

    // Checks used identically across XLSX and PPTX
    public static void chartTitleCheck(string sSheet, string sChartName, dynamic oChart) {
        bool bHasTitle = false;
        try { bHasTitle = (bool)oChart.HasTitle; } catch {}
        if (!bHasTitle) {
            results.add("ChartMissingTitle", "MSAC", "Object has no title or name",
                sSheet, "Chart: " + sChartName,
                "A chart has no title. Without a title, users cannot identify what the chart represents.",
                "Go to Chart Design > Add Chart Element > Chart Title and enter a descriptive title.");
        }
        try {
            bool b = (bool)oChart.Axes(1).HasTitle;
            if (!b) {
                results.add("ChartMissingAxisTitle", "AXE", "Missing table header",
                    sSheet, "Chart: " + sChartName + " (category axis)",
                    "Chart category axis has no title. Users cannot determine what categories are being compared.",
                    "Go to Chart Design > Add Chart Element > Axis Titles and label the category axis.");
            }
        } catch {}
        try {
            bool b = (bool)oChart.Axes(2).HasTitle;
            if (!b) {
                results.add("ChartMissingAxisTitle", "AXE", "Missing table header",
                    sSheet, "Chart: " + sChartName + " (value axis)",
                    "Chart value axis has no title. Users cannot determine what quantity is being measured.",
                    "Go to Chart Design > Add Chart Element > Axis Titles and label the value axis. Include units where applicable.");
            }
        } catch {}
    }

    // Writes extCheck.csv — the master rule registry
    public static void writeRulesCsv(string sPath) {
        var rows = new List<string[]> {
            new[] { "RuleID","MSOfficeCategory","WCAGCriteria","Severity","AppliesTo","Description","Remediation" },
            new[] { "AllCapsText","Text","1.4.8","Warning","DOCX, PPTX, MD","Text is typed entirely in capital letters. TTS engines may read each letter individually or stress words unnaturally.","Use mixed-case text. Apply a character style or CSS rule with text-transform:uppercase rather than typing in capitals." },
            new[] { "AltTextIsFilename","Missing alternative text","1.1.1","Error","MD","Alt text appears to be a filename rather than a meaningful description.","Replace the filename with a description of what the image shows and why it matters in context." },
            new[] { "AltTextRedundantPrefix","Missing alternative text","1.1.1","Warning","MD","Alt text begins with 'image of', 'picture of', or 'photo of'. Screen readers already announce the element is an image.","Remove the redundant prefix and begin the alt text with the actual description." },
            new[] { "AltTextTooShort","Missing alternative text","1.1.1","Warning","MD","Alt text is present but very short (fewer than 3 characters) and unlikely to convey meaningful information.","Expand the alt text to describe the image content and its purpose in context." },
            new[] { "BareUrl","Hyperlink","2.4.4","Warning","MD","A raw URL appears in text without descriptive link text. Screen readers read the URL character by character.","Wrap the URL in a Markdown link: [Descriptive text](https://url)." },
            new[] { "BlankOrGenericColumnHeader","Missing table header","1.3.1","Error","XLSX","An Excel Table column header is blank or uses a generic label like Column1.","Click the header cell and enter a concise, descriptive label." },
            new[] { "BlankRowsUsedForLayout","Use of blank cells for formatting","1.3.2","Warning","XLSX","Multiple consecutive blank rows within data suggest blank rows are used for visual spacing.","Remove blank rows used for spacing. Use row height and cell padding via Format Cells instead." },
            new[] { "BlankTableHeader","Missing table header","1.3.1","Error","MD","A Pandoc pipe table has one or more blank column header cells.","Add a descriptive label to each column header cell." },
            new[] { "ChartMissingAltText","Missing alternative text","1.1.1","Error","XLSX, PPTX","A chart has no alt text. Without alt text a blind user receives no information about the chart's content or key finding.","Select the chart, right-click its border, choose Edit Alt Text, and describe the chart's key finding." },
            new[] { "ChartMissingAxisTitle","Missing table header","1.3.1","Warning","XLSX, PPTX","A chart axis has no title label. Users cannot determine what quantity or categories are represented.","Select the chart and add axis titles via Chart Design > Add Chart Element > Axis Titles. Include units where applicable." },
            new[] { "ChartMissingTitle","Object has no title or name","1.3.1","Error","XLSX, PPTX","A chart has no title. Without a title, users cannot identify what the chart represents.","Select the chart, go to Chart Design > Add Chart Element > Chart Title, and enter a descriptive title." },
            new[] { "CodeBlockMissingLanguage","Code","1.3.3","Warning","MD","A fenced code block has no language specifier. Screen reader users receive no cue about what kind of code they are reading.","Add a language identifier after the opening fence: ```python, ```bash, ```json, etc." },
            new[] { "ColorUsedAlone","Use of color alone","1.4.1","Warning","XLSX","Empty cells with background fill color may be conveying status through color alone.","Add a text label or symbol inside the colored cell that conveys the same meaning as the color." },
            new[] { "ComplexTableHeaders","Complex table","1.3.1","Warning","DOCX","Table has header cells in both the first row and first column. Screen readers may not correctly associate data cells with both row and column headers.","Verify screen reader navigation. If possible, restructure into simpler tables. Test with NVDA or JAWS." },
            new[] { "DefaultSheetName","Object has no title or name","2.4.6","Warning","XLSX","A worksheet tab has a default name (Sheet1, Sheet2, etc.). Screen readers announce the tab name when navigating.","Right-click the tab and choose Rename. Enter a concise, descriptive name." },
            new[] { "DefaultTableName","Object has no title or name","2.4.6","Warning","XLSX","An Excel Table has a default name (Table1, Table2, etc.).","Click inside the table, go to Table Design, and replace the default name with a descriptive identifier. No spaces allowed." },
            new[] { "DuplicateHeadingText","Heading issues","2.4.6","Warning","DOCX, PPTX, MD","Two headings at the same level share identical text, creating ambiguous navigation in screen reader heading lists.","Make each heading uniquely descriptive of its section." },
            new[] { "EmptyHeading","Heading issues","1.3.1","Error","MD","A heading marker is present but contains no text. Empty headings create phantom navigation stops.","Add descriptive text after the heading marker, or remove the marker if no heading is intended." },
            new[] { "EmptyLinkText","Hyperlink","2.4.4","Error","MD","A Markdown link has no visible text. Screen readers read the raw URL or produce no output.","Add descriptive text between the square brackets that describes where the link leads." },
            new[] { "EmptyTableCell","Blank cells used for formatting","1.3.1","Warning","DOCX, XLSX, MD","A table cell is empty. Screen reader users may be uncertain whether the cell is intentionally blank.","Add content, use a dash or N/A for missing data, or restructure the table to eliminate unnecessary cells." },
            new[] { "EntireLineBolded","Text","1.3.1","Warning","MD","An entire line is bold. Bolding long passages reduces the emphasis signal for screen reader users.","Use bold sparingly for key terms. If the intent is a heading, use a heading marker instead." },
            new[] { "ExcessiveBlankLines","Repeated blank characters","1.3.2","Warning","MD","Three or more consecutive blank lines used for visual spacing.","Remove extra blank lines. Use a single blank line to separate paragraphs." },
            new[] { "ExcessiveTrailingSpaces","Repeated blank characters","1.3.2","Warning","MD","A line has more than two trailing spaces. Exactly two trailing spaces are an intentional hard line break in Pandoc Markdown; more than two is almost always unintentional.","Remove the extra trailing spaces." },
            new[] { "FakeInlineBullet","List not used correctly","1.3.1","Warning","MD","A Unicode bullet character is used as an inline list marker rather than a Markdown list item.","Use proper Markdown list items: start each item on its own line with - or * followed by a space." },
            new[] { "FakeListBullet","List not used correctly","1.3.1","Warning","DOCX, PPTX, MD","A paragraph uses a manually typed bullet character instead of the application's built-in list style.","Apply the List Bullet built-in style (Word), a bullet list from the Home tab (PowerPoint), or - / * syntax (Markdown)." },
            new[] { "FakeNumberedList","List not used correctly","1.3.1","Warning","MD","Lines appear to be a manually numbered list outside proper Markdown ordered list syntax.","Use Markdown ordered list syntax: start each item with '1.' followed by a space." },
            new[] { "FastAutoAdvance","Timing","2.2.1","Error","PPTX","A slide auto-advances in less than 3 seconds, which may not give screen reader users enough time to hear the content.","Increase the auto-advance time to at least 5 seconds, or disable auto-advance. Go to Transitions > Advance Slide > After." },
            new[] { "FastTransitionSpeed","Motion","2.3.3","Warning","PPTX","A slide uses Fast transition speed. Rapid visual motion can trigger symptoms in users with vestibular disorders.","Set the transition speed to Medium or Slow in the Transitions tab Duration field." },
            new[] { "FloatingTextBox","Object not inline","1.3.2","Warning","DOCX","A floating text box is read by screen readers in insertion order, which may differ from visual reading order.","Convert to inline body text if possible. Otherwise verify reading order in the Selection Pane." },
            new[] { "HeaderRowNotFrozen","Navigation","1.3.1","Warning","XLSX","A large worksheet has no frozen panes. The header row scrolls out of view, removing column context for screen reader users.","Click cell A2, go to View > Freeze Panes > Freeze Panes." },
            new[] { "HeadingEndsWithPunctuation","Object has no title or name","1.3.1","Warning","MD","A heading ends with a period, semicolon, or colon, causing TTS to insert an unnatural pause.","Remove the trailing punctuation from the heading." },
            new[] { "HeadingTooLong","Object has no title or name","2.4.6","Warning","MD","A heading is over 120 characters, which is difficult to scan in a screen reader heading list.","Shorten the heading to under 80 characters. Move additional content into the paragraph below." },
            new[] { "HiddenColumns","Object not inline","1.3.2","Warning","XLSX","Hidden columns disrupt the continuous data sequence and may be skipped by screen readers.","Unhide columns. Consider moving the data to a separate sheet or documenting hidden columns with a visible note." },
            new[] { "HiddenRows","Object not inline","1.3.2","Warning","XLSX","Hidden rows disrupt the continuous data sequence and may be skipped by screen readers.","Unhide rows. Consider using grouping (Data > Group) or filters instead." },
            new[] { "HyperlinkNotUnderlined","Hyperlink","1.4.1","Warning","DOCX","A hyperlink is not underlined, making it indistinguishable from body text for users with color vision deficiencies.","Ensure the Hyperlink character style includes underlining: Home > Styles > right-click Hyperlink > Modify." },
            new[] { "IndentedCodeBlock","Code","1.3.3","Warning","MD","An indented code block (4-space indent) cannot carry a language specifier.","Convert to a fenced code block using ``` with a language identifier." },
            new[] { "LongSectionWithoutHeading","Heading issues","2.4.6","Warning","DOCX, MD","More than 20 paragraphs appear without any heading. Screen reader users rely on headings to navigate.","Add appropriately leveled headings to mark major sections." },
            new[] { "MergedCells","Merged cells","1.3.1","Warning","XLSX","Merged cells disrupt screen reader table navigation and break sort and filter operations.","Remove merged cells: Home > Merge & Center > Unmerge Cells. Use Center Across Selection for visual centering." },
            new[] { "MissingAltText","Missing alternative text","1.1.1","Error","DOCX, XLSX, PPTX, MD","An image, shape, chart, or other non-text object has no alternative text.","Right-click the object, choose Edit Alt Text, and write a concise description. If purely decorative, mark as decorative." },
            new[] { "MissingAuthor","Missing document properties","2.4.2","Warning","DOCX, XLSX, PPTX, MD","The document Author property is not set.","Go to File > Info > Properties and enter the author name. In Markdown, add author: to the YAML front matter." },
            new[] { "MissingDescription","Missing document properties","2.4.2","Warning","MD","No description field is present in the Pandoc YAML front matter.","Add description: with a one-to-two sentence summary to the YAML front matter." },
            new[] { "MissingDocumentLanguage","Missing document language","3.1.1","Error","DOCX","The document's proofing language is not set. Screen readers use this to select the correct TTS voice.","Go to Review > Language > Set Proofing Language and select the appropriate language." },
            new[] { "MissingDocumentTitle","Missing document properties","2.4.2","Error","DOCX","The document Title property is not set. Screen readers announce the title when the document is opened.","Go to File > Info > Properties and enter a descriptive title in the Title field." },
            new[] { "MissingLanguage","Missing document language","3.1.1","Error","MD","No lang field in the YAML front matter. Pandoc uses this to set the HTML lang attribute and PDF language metadata.","Add lang: en-US (or appropriate BCP 47 tag) to the YAML front matter." },
            new[] { "MissingPresentationAuthor","Missing document properties","2.4.2","Warning","PPTX","The presentation Author property is not set.","Go to File > Info > Properties and enter the author name." },
            new[] { "MissingPresentationTitle","Missing document properties","2.4.2","Error","PPTX","The presentation Title property is not set. Screen readers announce the title when the file is opened.","Go to File > Info > Properties and enter a descriptive title." },
            new[] { "MissingTitle","Missing document properties","2.4.2","Error","MD","No title field in the YAML front matter. Screen readers announce the document title when it is opened.","Add title: \"My Document Title\" to the YAML front matter." },
            new[] { "MissingWorkbookAuthor","Missing document properties","2.4.2","Warning","XLSX","The workbook Author property is not set.","Go to File > Info > Properties and enter the author name." },
            new[] { "MissingWorkbookTitle","Missing document properties","2.4.2","Error","XLSX","The workbook Title property is not set. Screen readers announce the title when the file is opened.","Go to File > Info > Properties and enter a descriptive title." },
            new[] { "MovingContent","Motion","2.2.2","Error","MD","A raw HTML marquee or blink element causes moving or blinking content, violating WCAG 2.2.2.","Remove the element entirely. Moving and blinking text has no accessible equivalent." },
            new[] { "NoHeadings","Heading issues","2.4.6","Warning","MD","A document of substantial length has no headings. Screen reader users navigate long documents by heading list.","Add ATX-style headings (# H1, ## H2) to mark major sections." },
            new[] { "NoSpeakerNotes","Missing alternative text","1.1.1","Warning","PPTX","No slides have speaker notes. For distributed presentations, notes provide essential context for screen reader users.","Add speaker notes via View > Notes. Describe visual content and anything conveyed only through visuals." },
            new[] { "NoYamlFrontMatter","Missing document properties","2.4.2","Error","MD","The Markdown file has no YAML front matter block. Without front matter, Pandoc cannot set document title, language, or author.","Add a YAML front matter block starting and ending with --- at the very top of the file. Include at minimum: title and lang." },
            new[] { "NonDescriptiveHyperlinkText","Hyperlink","2.4.4","Error","DOCX, XLSX","Hyperlink display text is non-descriptive or a raw URL. Screen reader users navigating by link list cannot determine the destination.","Edit the hyperlink display text to describe the destination. The text should make sense when read in isolation." },
            new[] { "NonDescriptiveLinkText","Hyperlink","2.4.4","Error","MD","Markdown link text is non-descriptive (click here, here, read more, etc.).","Replace with text that describes the link destination." },
            new[] { "RapidAutoAnimation","Motion","2.2.2","Warning","PPTX","Multiple animations trigger automatically with very short delays, which can overwhelm screen reader users.","Increase delays to at least 1 second per element, or switch to On Click triggers." },
            new[] { "RawBrTag","Repeated blank characters","1.3.2","Warning","MD","A raw HTML br tag is used for a line break, reducing portability to non-HTML Pandoc output formats.","End the line with two trailing spaces or a backslash for a portable hard line break." },
            new[] { "RawHtmlGenericContainer","Use of color alone","1.3.1","Warning","MD","A raw HTML div or span is silently ignored by Pandoc when converting to non-HTML formats.","Use Pandoc's native fenced div (::: {.class}) or bracketed span ([text]{.class}) syntax instead." },
            new[] { "RawHtmlPresentational","Use of color alone","1.3.3","Warning","MD","A raw HTML presentational element (font, center) conveys no semantics and is not supported in non-HTML Pandoc output formats.","Replace with Markdown formatting or a Pandoc-native span/div." },
            new[] { "RawHtmlTable","Missing table header","1.3.1","Warning","MD","A raw HTML table is used. HTML tables require explicit th scope headers and a caption to be accessible.","Use Pandoc pipe table syntax, or ensure the HTML table includes th scope headers and a caption." },
            new[] { "RepeatedBlankCharacters","Repeated blank characters","1.3.2","Warning","DOCX","A paragraph contains consecutive spaces likely used for visual layout.","Remove repeated spaces. Use paragraph indentation or tab stops for layout instead." },
            new[] { "RepeatingAnimation","Motion","2.2.2","Warning","PPTX","A repeating or looping animation violates WCAG 2.2.2 because the user cannot stop it without leaving the slide.","Set the animation repeat count to 1 in the Effect Options dialog." },
            new[] { "SheetNameTooLong","Object has no title or name","2.4.6","Warning","XLSX","A worksheet tab name exceeds 31 characters, which is Excel's maximum.","Right-click the tab, choose Rename, and shorten the name to 31 characters or fewer." },
            new[] { "SkippedHeadingLevel","Heading issues","1.3.1","Error","DOCX, PPTX, MD","The document skips one or more heading levels (e.g., H1 followed directly by H3).","Change the heading style to restore sequential order. Every level must appear before a deeper level is used." },
            new[] { "SlideMissingTitle","Object has no title or name","1.3.1","Error","PPTX","A slide has no title or the title placeholder is empty. Screen readers announce slide titles as navigation landmarks.","Click the title placeholder and enter a concise, descriptive title." },
            new[] { "SmallFontSize","Sufficient contrast","1.4.3","Warning","PPTX","Text on a slide is smaller than 18pt, which may be too small for users with low vision.","Increase font size to at least 18pt for body text and 24pt or larger for key content." },
            new[] { "TableColumnCountMismatch","Complex table","1.3.1","Error","MD","A Pandoc pipe table row has a different number of cells than the header row.","Ensure every row has the same number of pipe-delimited cells as the header row." },
            new[] { "TableMissingHeaderRow","Missing table header","1.3.1","Error","DOCX","A Word table does not have a designated header row.","Click in the first row, go to Table Design > Table Style Options and check Header Row." },
            new[] { "TableMissingHeaders","Missing table header","1.3.1","Error","XLSX","An Excel Table has its header row turned off.","Click inside the table, go to Table Design, and check the Header Row checkbox." },
            new[] { "TitleNotFirstInReadingOrder","Object not inline","1.3.2","Warning","PPTX","The title placeholder is not at the back of the z-order. PowerPoint screen readers traverse shapes from back to front.","Open the Selection Pane (Home > Arrange > Selection Pane) and drag the title placeholder to the bottom of the list." },
            new[] { "UnlabeledFormField","Object has no title or name","1.3.1","Error","DOCX","A Word form field has no help text. Screen readers announce only the field type with no indication of what is requested.","Open the field's properties and enter a descriptive label in the Status Bar Help Text or Help Key field." },
            new[] { "UrlAsLinkText","Hyperlink","2.4.4","Error","MD","A Markdown link uses a raw URL as its display text. Screen readers read the URL character by character.","Replace the URL display text with a description of the link destination." },
            new[] { "VisualContentWithoutNotes","Missing alternative text","1.1.1","Error","PPTX","A slide has images or charts without alt text and no speaker notes, making it entirely inaccessible.","Add alt text to each image and chart, and add speaker notes describing the key visual content." },
        };

        var sb = new StringBuilder();
        foreach (string[] row in rows) {
            sb.AppendLine(csvRow(row));
        }
        File.WriteAllText(sPath, sb.ToString(), Encoding.UTF8);
        Console.WriteLine("Rules CSV written: " + sPath);
        Console.WriteLine("Rules: " + (rows.Count - 1));
    }

    static string csvRow(string[] aFields) {
        var sb = new StringBuilder();
        for (int i = 0; i < aFields.Length; i++) {
            if (i > 0) {
                sb.Append(",");
            }
            sb.Append(esc(aFields[i]));
        }
        return sb.ToString();
    }

    static string esc(string s) {
        s = s.Replace("\"", "\"\"");
        if (s.Contains(",") || s.Contains("\"") || s.Contains("\r") || s.Contains("\n")) {
            s = "\"" + s + "\"";
        }
        return s;
    }
}

// ===========================================================================
// Shared: COM lifecycle helpers
// ===========================================================================
static class comHelper {
    public static dynamic createApp(string sProgId) {
        Type t = Type.GetTypeFromProgID(sProgId);
        if (t == null) {
            // ProgID not registered. Two common causes:
            //  (1) Office (or the relevant app) is not installed.
            //  (2) Bitness mismatch: this process is 64-bit and the
            //      installed Office is 32-bit (or vice versa). Modern
            //      Office is 64-bit by default but legacy 32-bit
            //      installations still exist.
            string sBits = (IntPtr.Size == 8) ? "64-bit" : "32-bit";
            string sOther = (IntPtr.Size == 8) ? "32-bit" : "64-bit";
            throw new Exception(
                sProgId + " is not registered for this process.\r\n" +
                "  This extCheck.exe is " + sBits + ".\r\n" +
                "  Possible causes:\r\n" +
                "    * The required Office application is not installed.\r\n" +
                "    * The installed Office is " + sOther +
                " and cannot be automated by a " + sBits + " process.\r\n" +
                "  Fix: install the matching Office bitness, or use the " +
                sOther + " build of extCheck.");
        }
        dynamic oApp = Activator.CreateInstance(t);
        silenceAlerts(sProgId, oApp);
        return oApp;
    }

    /// <summary>
    /// Silence interactive alerts on a freshly-created Office
    /// Application object so unattended automation does not stall on a
    /// dialog. Mirrors the equivalent helper in 2htm. Each setter is
    /// wrapped in its own try/catch because not all versions of Office
    /// expose every property and some throw harmlessly.
    /// </summary>
    private static void silenceAlerts(string sProgId, dynamic oApp) {
        const int iMsoAutomationSecurityForceDisable = 3;
        const int iWdAlertsNone = 0;
        const int iPpAlertsNone = 1;
        string sLowered = (sProgId ?? "").ToLowerInvariant();
        try { oApp.AutomationSecurity = iMsoAutomationSecurityForceDisable; } catch { }
        if (sLowered.StartsWith("word.")) {
            try { oApp.DisplayAlerts = iWdAlertsNone; } catch { }
            try { oApp.Visible = false; } catch { }
            try { oApp.Options.ConfirmConversions = false; } catch { }
            try { oApp.Options.DoNotPromptForConvert = true; } catch { }
            try { oApp.Options.SaveNormalPrompt = false; } catch { }
            try { oApp.Options.WarnBeforeSavingPrintingSendingMarkup = false; } catch { }
            try { oApp.Options.UpdateLinksAtOpen = false; } catch { }
            try { oApp.Options.CheckGrammarAsYouType = false; } catch { }
            try { oApp.Options.CheckSpellingAsYouType = false; } catch { }
        }
        else if (sLowered.StartsWith("excel.")) {
            try { oApp.DisplayAlerts = false; } catch { }
            try { oApp.Visible = false; } catch { }
            try { oApp.AskToUpdateLinks = false; } catch { }
            try { oApp.AlertBeforeOverwriting = false; } catch { }
            try { oApp.ScreenUpdating = false; } catch { }
            try { oApp.EnableEvents = false; } catch { }
        }
        else if (sLowered.StartsWith("powerpoint.")) {
            try { oApp.DisplayAlerts = iPpAlertsNone; } catch { }
            // PowerPoint cannot run with Visible = false.
        }
    }

    public static void safeClose(dynamic o) { try { o.Close(false); } catch {} }
    public static void safeQuit(dynamic o)  { try { o.Quit(); }       catch {} }
}

// ===========================================================================
// Format module: DOCX
// ===========================================================================
static class docxModule {
    const int iMaxParasBetweenHeadings = 20;
    const int iWdLanguageNone          = 1024;
    const int iWdShapeTextBox          = 17;
    const string sPfx                  = "Heading ";

    static dynamic oWord = null;
    static dynamic oDoc  = null;

    public static bool open(string sPath) {
        try {
            oWord = comHelper.createApp("Word.Application");
            oWord.Visible      = false;
            oWord.DisplayAlerts = 0;
            oDoc  = oWord.Documents.Open(sPath, false, true);
            return true;
        } catch (Exception ex) {
            Console.WriteLine("  ERROR: " + ex.Message);
            quit();
            return false;
        }
    }

    public static void quit() {
        try { if (oDoc  != null) oDoc.Close(false); } catch {}
        try { if (oWord != null) oWord.Quit(false); } catch {}
        object oRawDoc  = oDoc;  oDoc  = null;
        object oRawWord = oWord; oWord = null;
        try { if (oRawDoc  != null) Marshal.ReleaseComObject(oRawDoc);  } catch {}
        try { if (oRawWord != null) Marshal.ReleaseComObject(oRawWord); } catch {}
    }

    public static void checkAll(string sFilePath) {
        documentTitle(sFilePath);
        documentLanguage(sFilePath);
        author(sFilePath);
        images();
        floatingTextBoxes();
        hyperlinks();
        paragraphs();
        tables();
        formFields();
    }

    static void documentTitle(string p) {
        string s = "";
        try { s = (oDoc.BuiltInDocumentProperties["Title"].Value ?? "").ToString().Trim(); } catch {}
        if (s == "") {
            results.add("MissingDocumentTitle", "MSAC", "Missing document properties", "(Document)", p,
                "The document Title property is not set. Screen readers announce the title when the document is opened.",
                "Go to File > Info > Properties and enter a descriptive title. This is separate from the filename.");
        }
    }

    static void documentLanguage(string p) {
        int i = 0;
        try { i = (int)oDoc.Paragraphs[1].Range.LanguageID; } catch {}
        if (i == iWdLanguageNone || i == 0) {
            results.add("MissingDocumentLanguage", "MSAC", "Missing document language", "(Document)", p,
                "The document's proofing language is not set. Screen readers use this to select the correct TTS voice.",
                "Go to Review > Language > Set Proofing Language and select the appropriate language.");
        }
    }

    static void author(string p) {
        string s = "";
        try { s = (oDoc.BuiltInDocumentProperties["Author"].Value ?? "").ToString().Trim(); } catch {}
        if (s == "") {
            results.add("MissingAuthor", "AXE", "Missing document properties", "(Document)", p,
                "The document Author property is not set.",
                "Go to File > Info > Properties and enter the author name.");
        }
    }

    static void images() {
        foreach (dynamic o in oDoc.InlineShapes) {
            string sAlt = "";
            try { sAlt = (o.AlternativeText ?? "").ToString().Trim(); } catch {}
            if (sAlt == "") {
                try { sAlt = (o.Title ?? "").ToString().Trim(); } catch {}
            }
            if (sAlt == "") {
                results.add("MissingAltText", "MSAC", "Missing alternative text", "(Document)", "Inline image",
                    "An inline image has no alternative text.",
                    "Right-click the image, choose Edit Alt Text, and write a concise description. If purely decorative, check the Decorative checkbox.");
            }
        }
        foreach (dynamic o in oDoc.Shapes) {
            int iType = 0;
            string sName = "";
            try { iType = (int)o.Type; } catch {}
            try { sName = o.Name.ToString(); } catch {}
            if (iType == iWdShapeTextBox) {
                continue;
            }
            string sAlt = "";
            try { sAlt = (o.AlternativeText ?? "").ToString().Trim(); } catch {}
            if (sAlt == "") {
                results.add("MissingAltText", "MSAC", "Missing alternative text", "(Document)", "Floating shape: " + sName,
                    "A floating shape has no alternative text.",
                    "Right-click the shape, choose Edit Alt Text, and write a concise description.");
            }
        }
    }

    static void floatingTextBoxes() {
        foreach (dynamic o in oDoc.Shapes) {
            int iType = 0;
            string sName = "";
            try { iType = (int)o.Type; } catch {}
            try { sName = o.Name.ToString(); } catch {}
            if (iType == iWdShapeTextBox) {
                results.add("FloatingTextBox", "MSAC", "Object not inline", "(Document)", "Shape: " + sName,
                    "A floating text box is read by screen readers in insertion order, which may differ from visual reading order.",
                    "Convert to inline body text if possible. Otherwise verify reading order in the Selection Pane.");
            }
        }
    }

    static void hyperlinks() {
        foreach (dynamic oLink in oDoc.Hyperlinks) {
            string sText = "";
            try { sText = (oLink.TextToDisplay ?? "").ToString().Trim(); } catch {}
            if (shared.isVagueLinkText(sText)) {
                results.add("NonDescriptiveHyperlinkText", "AXE", "Hyperlink", "(Document)", "\"" + sText + "\"",
                    "Hyperlink display text is non-descriptive or a raw URL.",
                    "Edit the hyperlink text to describe the destination. It should make sense when read in isolation.");
            }
            int iU = 0;
            try { iU = (int)oLink.Range.Font.Underline; } catch {}
            if (iU == 0) {
                results.add("HyperlinkNotUnderlined", "AXE", "Hyperlink", "(Document)", "\"" + sText + "\"",
                    "Hyperlink is not underlined, making it indistinguishable from body text for color-blind users (WCAG 1.4.1).",
                    "Ensure the Hyperlink character style includes underlining: Home > Styles > right-click Hyperlink > Modify.");
            }
        }
    }

    static void paragraphs() {
        int iLastLv     = 0;
        int iParasSince = 0;
        bool bLongFired = false;
        var lHdgs = new List<string>();

        foreach (dynamic oPara in oDoc.Paragraphs) {
            string sStyle = "";
            string sText  = "";
            try { sStyle = oPara.Style.NameLocal.ToString(); } catch {}
            try { sText  = oPara.Range.Text.ToString().Trim(); } catch {}

            if (sStyle.StartsWith(sPfx)) {
                int iLv = 0;
                if (sStyle.Length > sPfx.Length && int.TryParse(sStyle.Substring(sPfx.Length, 1), out iLv)) {
                    if (iLastLv > 0 && iLv > iLastLv + 1) {
                        results.add("SkippedHeadingLevel", "MSAC", "Heading issues", "(Document)", shared.trunc(sText),
                            "Heading jumps from H" + iLastLv + " to H" + iLv + ", skipping a level.",
                            "Change this heading style to H" + (iLastLv + 1) + " so levels are sequential.");
                    }
                    string sKey = iLv + "|" + sText.ToLower().Trim();
                    if (lHdgs.Contains(sKey)) {
                        results.add("DuplicateHeadingText", "AXE", "Heading issues", "(Document)", shared.trunc(sText),
                            "Two H" + iLv + " headings share identical text.",
                            "Make each heading uniquely descriptive of its section.");
                    } else {
                        lHdgs.Add(sKey);
                    }
                    iLastLv     = iLv;
                    iParasSince = 0;
                    bLongFired  = false;
                }
            } else {
                if (sText != "") {
                    iParasSince++;
                }
                if (iParasSince > iMaxParasBetweenHeadings && !bLongFired && iLastLv == 0) {
                    results.add("LongSectionWithoutHeading", "MSAC", "Heading issues", "(Document)", "Start of document",
                        "More than " + iMaxParasBetweenHeadings + " paragraphs appear before any heading.",
                        "Add Heading styles to mark major sections. Apply from the Home tab Styles gallery.");
                    bLongFired = true;
                }
            }

            if (sText.Contains("  ")) {
                results.add("RepeatedBlankCharacters", "MSAC", "Repeated blank characters", "(Document)", shared.trunc(sText),
                    "Paragraph contains consecutive spaces likely used for visual layout.",
                    "Remove repeated spaces. Use paragraph indentation or tab stops for layout instead.");
            }

            if (sText.Length > 4 && sText == sText.ToUpper() && sText != sText.ToLower()) {
                results.add("AllCapsText", "MSAC", "Text", "(Document)", shared.trunc(sText),
                    "Paragraph is typed entirely in capitals. TTS engines may read each letter individually.",
                    "Use mixed-case text. Apply AllCaps character formatting if the visual style is required.");
            }

            if (sText.Length > 1) {
                string sFirst = sText.Substring(0, 1);
                foreach (string b in shared.aBullets) {
                    if (sFirst == b && !sStyle.ToLower().Contains("list")) {
                        results.add("FakeListBullet", "MSAC", "List not used correctly", "(Document)", shared.trunc(sText),
                            "Paragraph uses a manually typed bullet instead of a Word list style.",
                            "Apply the 'List Bullet' built-in style from the Home tab Styles gallery.");
                    }
                }
            }
        }
    }

    static void tables() {
        int iNum = 0;
        foreach (dynamic oTable in oDoc.Tables) {
            iNum++;
            bool bHdr = false;
            try { bHdr = (bool)oTable.Rows[1].HeadingFormat; } catch {}
            if (!bHdr) {
                results.add("TableMissingHeaderRow", "MSAC", "Missing table header", "(Document)", "Table " + iNum,
                    "Table does not have a designated header row. Screen readers cannot identify column headers.",
                    "Click in the first row, go to Table Design > Table Style Options and check Header Row.");
            }

            bool bFstCol = false;
            try {
                int r = (int)oTable.Rows.Count;
                int c = (int)oTable.Columns.Count;
                if (r > 1 && c > 1) {
                    string cs = oTable.Cell(2, 1).Range.Paragraphs[1].Style.NameLocal.ToString();
                    if (cs.ToLower().Contains("header")) {
                        bFstCol = true;
                    }
                }
            } catch {}
            if (bHdr && bFstCol) {
                results.add("ComplexTableHeaders", "MSAC", "Complex table", "(Document)", "Table " + iNum,
                    "Table has header cells in both the first row and first column.",
                    "Verify screen reader navigation. If possible, restructure into simpler tables. Test with NVDA or JAWS.");
            }

            int iRow = 0;
            foreach (dynamic oRow in oTable.Rows) {
                iRow++;
                int iCol = 0;
                foreach (dynamic oCell in oRow.Cells) {
                    iCol++;
                    string sC = "";
                    try {
                        sC = oCell.Range.Text.ToString();
                        if (sC.Length >= 2) {
                            sC = sC.Substring(0, sC.Length - 2);
                        }
                        sC = sC.Trim();
                    } catch {}
                    if (sC == "" && !(iRow == 1 && bHdr)) {
                        results.add("EmptyTableCell", "MSAC", "Blank cells used for formatting", "(Document)",
                            "Table " + iNum + " Row " + iRow + " Col " + iCol,
                            "Table cell is empty.",
                            "Add content, use N/A for missing data, or restructure to eliminate unnecessary cells.");
                    }
                }
            }
        }
    }

    static void formFields() {
        foreach (dynamic oField in oDoc.FormFields) {
            string sName = "";
            string sHelp = "";
            try { sName = oField.Name.ToString(); } catch {}
            try { sHelp = (oField.StatusBarText ?? "").ToString().Trim(); } catch {}
            if (sHelp == "") {
                try { sHelp = (oField.HelpText ?? "").ToString().Trim(); } catch {}
            }
            if (sHelp == "") {
                results.add("UnlabeledFormField", "AXE", "Object has no title or name", "(Document)", "Field: " + sName,
                    "Form field has no help text. Screen readers announce only the field type with no indication of what is requested.",
                    "Open the field's properties and enter a descriptive label in the Status Bar Help Text field.");
            }
        }
    }
}

// ===========================================================================
// Format module: XLSX
// ===========================================================================
static class xlsxModule {
    const int iMaxConsecutiveEmptyRows = 5;
    const int iMaxSheetNameLength      = 31;

    static dynamic oExcel = null;
    static dynamic oWb    = null;

    public static bool open(string sPath) {
        try {
            oExcel = comHelper.createApp("Excel.Application");
            oExcel.Visible       = false;
            oExcel.DisplayAlerts = false;
            oWb    = oExcel.Workbooks.Open(sPath, false, true);
            return true;
        } catch (Exception ex) {
            Console.WriteLine("  ERROR: " + ex.Message);
            quit();
            return false;
        }
    }

    public static void quit() {
        try { if (oWb    != null) oWb.Close(false); } catch {}
        try { if (oExcel != null) oExcel.Quit();    } catch {}
        object oRawWb    = oWb;    oWb    = null;
        object oRawExcel = oExcel; oExcel = null;
        try { if (oRawWb    != null) Marshal.ReleaseComObject(oRawWb);    } catch {}
        try { if (oRawExcel != null) Marshal.ReleaseComObject(oRawExcel); } catch {}
    }

    public static void checkAll(string sFilePath) {
        workbookMeta(sFilePath);
        foreach (dynamic oSheet in oWb.Sheets) {
            string sName = "(unknown)";
            string sType = "";
            try { sName = oSheet.Name.ToString(); } catch {}
            try { sType = oSheet.Type.ToString(); } catch {}
            if (sType == "3") {
                continue; // xlChart
            }
            sheetImages(oSheet, sName);
            sheetCharts(oSheet, sName);
            sheetHyperlinks(oSheet, sName);
            sheetMergedCells(oSheet, sName);
            sheetTables(oSheet, sName);
            sheetColorOnly(oSheet, sName);
            sheetEmptyRegions(oSheet, sName);
            sheetFrozenPanes(oSheet, sName);
            sheetHiddenContent(oSheet, sName);
        }
    }

    static void workbookMeta(string sFilePath) {
        string sTitle = "";
        try { sTitle = (oWb.BuiltinDocumentProperties["Title"].Value ?? "").ToString().Trim(); } catch {}
        if (sTitle == "") {
            results.add("MissingWorkbookTitle", "MSAC", "Missing document properties", "(Workbook)", sFilePath,
                "The workbook Title property is not set.",
                "Go to File > Info > Properties and enter a descriptive title.");
        }
        string sAuthor = "";
        try { sAuthor = (oWb.BuiltinDocumentProperties["Author"].Value ?? "").ToString().Trim(); } catch {}
        if (sAuthor == "") {
            results.add("MissingWorkbookAuthor", "AXE", "Missing document properties", "(Workbook)", sFilePath,
                "The workbook Author property is not set.",
                "Go to File > Info > Properties and enter the author name.");
        }
        try {
            foreach (dynamic oSheet in oWb.Sheets) {
                string sName = oSheet.Name.ToString();
                bool bDefault = sName.StartsWith("Sheet") && sName.Length <= 7;
                if (bDefault) {
                    bool bNum = true;
                    foreach (char c in sName.Substring(5)) {
                        if (!char.IsDigit(c)) {
                            bNum = false;
                        }
                    }
                    if (bNum) {
                        results.add("DefaultSheetName", "MSAC", "Object has no title or name", sName, "Tab: " + sName,
                            "Sheet tab has a default name (Sheet1, Sheet2, etc.).",
                            "Right-click the tab and choose Rename. Enter a concise, descriptive name.");
                    }
                }
                if (sName.Length > iMaxSheetNameLength) {
                    results.add("SheetNameTooLong", "MSAC", "Object has no title or name", sName, "Tab: " + sName,
                        "Sheet name exceeds " + iMaxSheetNameLength + " characters, which is Excel's maximum.",
                        "Right-click the tab, choose Rename, and shorten the name to " + iMaxSheetNameLength + " characters or fewer.");
                }
            }
        } catch {}
    }

    static void sheetImages(dynamic oSheet, string sSheet) {
        try {
            foreach (dynamic oShape in oSheet.Shapes) {
                string sAlt  = "";
                string sName = "";
                int iType    = 0;
                try { sName = oShape.Name.ToString(); } catch {}
                try { sAlt  = (oShape.AlternativeText ?? "").ToString().Trim(); } catch {}
                try { iType = (int)oShape.Type; } catch {}
                if (iType == 9 || iType == 10) {
                    continue;
                }
                if (sAlt == "") {
                    results.add("MissingAltText", "MSAC", "Missing alternative text", sSheet, "Shape: " + sName,
                        "A shape or image has no alternative text.",
                        "Right-click the shape, choose Edit Alt Text, and write a concise description.");
                }
            }
        } catch {}
    }

    static void sheetCharts(dynamic oSheet, string sSheet) {
        try {
            foreach (dynamic oCO in oSheet.ChartObjects()) {
                string sName = "";
                try { sName = oCO.Name.ToString(); } catch {}
                string sAlt = "";
                try { sAlt = (oCO.ShapeRange.AlternativeText ?? "").ToString().Trim(); } catch {}
                if (sAlt == "") {
                    results.add("ChartMissingAltText", "MSAC", "Missing alternative text", sSheet, "Chart: " + sName,
                        "A chart has no alt text.",
                        "Select the chart, right-click its border, choose Edit Alt Text, and describe the chart's key finding.");
                }
                try { shared.chartTitleCheck(sSheet, sName, oCO.Chart); } catch {}
            }
        } catch {}
    }

    static void sheetHyperlinks(dynamic oSheet, string sSheet) {
        try {
            foreach (dynamic oLink in oSheet.Hyperlinks) {
                string sText = "";
                string sAddr = "";
                try { sText = (oLink.TextToDisplay ?? "").ToString().Trim(); } catch {}
                try { sAddr = (oLink.Address ?? "").ToString().Trim(); } catch {}
                if (shared.isVagueLinkText(sText) || sText == sAddr) {
                    results.add("NonDescriptiveHyperlinkText", "AXE", "Hyperlink", sSheet, "Link: \"" + sText + "\"",
                        "Hyperlink display text is non-descriptive or a raw URL.",
                        "Select the cell, press Ctrl+K, and update the Text to Display field with a meaningful description.");
                }
            }
        } catch {}
    }

    static void sheetMergedCells(dynamic oSheet, string sSheet) {
        int iCount = 0;
        try {
            var lSeen = new List<string>();
            foreach (dynamic oCell in oSheet.UsedRange.Cells) {
                bool bMerged = false;
                try { bMerged = (bool)oCell.MergeCells; } catch {}
                if (bMerged) {
                    string sAddr = "";
                    try { sAddr = oCell.MergeArea.Address.ToString(); } catch {}
                    if (!lSeen.Contains(sAddr)) {
                        lSeen.Add(sAddr);
                        iCount++;
                    }
                }
            }
        } catch {}
        if (iCount > 0) {
            results.add("MergedCells", "MSAC", "Merged cells", sSheet, iCount + " merged cell region(s)",
                "Merged cells disrupt screen reader table navigation and break sort and filter operations.",
                "Remove merged cells: Home > Merge & Center > Unmerge Cells. Use Center Across Selection instead.");
        }
    }

    static void sheetTables(dynamic oSheet, string sSheet) {
        try {
            foreach (dynamic oTable in oSheet.ListObjects) {
                string sTName = "";
                try { sTName = oTable.Name.ToString(); } catch {}
                bool bShowHdr = false;
                try { bShowHdr = (bool)oTable.ShowHeaders; } catch {}
                if (!bShowHdr) {
                    results.add("TableMissingHeaders", "MSAC", "Missing table header", sSheet, "Table: " + sTName,
                        "An Excel Table has its header row turned off.",
                        "Click inside the table, go to Table Design, and check the Header Row checkbox.");
                }
                if (sTName.StartsWith("Table") && sTName.Length <= 10) {
                    bool bDef = true;
                    foreach (char c in sTName.Substring(5)) {
                        if (!char.IsDigit(c)) {
                            bDef = false;
                        }
                    }
                    if (bDef) {
                        results.add("DefaultTableName", "MSAC", "Object has no title or name", sSheet, "Table: " + sTName,
                            "Excel Table has a default name (Table1, Table2, etc.).",
                            "Click inside the table, go to Table Design, and replace the default name with a descriptive identifier.");
                    }
                }
                try {
                    int iCol = 0;
                    foreach (dynamic oCell in oTable.HeaderRowRange.Cells) {
                        iCol++;
                        string sV = "";
                        try { sV = (oCell.Value2 ?? "").ToString().Trim(); } catch {}
                        if (sV == "" || sV.ToLower().StartsWith("column")) {
                            results.add("BlankOrGenericColumnHeader", "MSAC", "Missing table header", sSheet,
                                "Table: " + sTName + " Col " + iCol,
                                "Table column header is blank or uses a generic label like 'Column1'.",
                                "Click the header cell and enter a concise, descriptive label.");
                        }
                    }
                } catch {}
            }
        } catch {}
    }

    static void sheetColorOnly(dynamic oSheet, string sSheet) {
        int iCount = 0;
        try {
            foreach (dynamic oCell in oSheet.UsedRange.Cells) {
                string sV = "";
                try { sV = (oCell.Value2 ?? "").ToString().Trim(); } catch {}
                if (sV != "") {
                    continue;
                }
                long lColor = -1;
                try { lColor = (long)oCell.Interior.Color; } catch {}
                if (lColor >= 0 && lColor != 16777215) {
                    iCount++;
                }
            }
        } catch {}
        if (iCount > 0) {
            results.add("ColorUsedAlone", "AXE", "Use of color alone", sSheet, iCount + " empty cell(s) with background color",
                "Empty cells with background fill color may be conveying status through color alone (WCAG 1.4.1).",
                "Add a text label inside the colored cell (e.g., 'Overdue', 'Complete'). Do not rely on color alone.");
        }
    }

    static void sheetEmptyRegions(dynamic oSheet, string sSheet) {
        int iConsec = 0;
        int iMax    = 0;
        try {
            int iRows = (int)oSheet.UsedRange.Rows.Count;
            int iCols = (int)oSheet.UsedRange.Columns.Count;
            for (int r = 1; r <= iRows; r++) {
                bool bEmpty = true;
                for (int c = 1; c <= Math.Min(iCols, 20); c++) {
                    string sV = "";
                    try { sV = (oSheet.UsedRange.Cells[r, c].Value2 ?? "").ToString().Trim(); } catch {}
                    if (sV != "") {
                        bEmpty = false;
                        break;
                    }
                }
                if (bEmpty) {
                    iConsec++;
                    if (iConsec > iMax) {
                        iMax = iConsec;
                    }
                } else {
                    iConsec = 0;
                }
            }
        } catch {}
        if (iMax >= iMaxConsecutiveEmptyRows) {
            results.add("BlankRowsUsedForLayout", "MSAC", "Use of blank cells for formatting", sSheet,
                iMax + " consecutive empty rows",
                "Multiple consecutive empty rows within data suggest they are used for visual spacing.",
                "Remove blank rows used for spacing. Use row height and cell padding via Format Cells instead.");
        }
    }

    static void sheetFrozenPanes(dynamic oSheet, string sSheet) {
        bool bFrozen = false;
        int iRows    = 0;
        try { bFrozen = (bool)oSheet.Application.ActiveWindow.FreezePanes; } catch {}
        try { iRows   = (int)oSheet.UsedRange.Rows.Count; } catch {}
        if (!bFrozen && iRows > 20) {
            results.add("HeaderRowNotFrozen", "AXE", "Navigation", sSheet, "Sheet has " + iRows + " rows but no frozen panes",
                "On a large sheet the header row scrolls out of view, removing column context for screen reader users.",
                "Click cell A2, go to View > Freeze Panes > Freeze Panes.");
        }
    }

    static void sheetHiddenContent(dynamic oSheet, string sSheet) {
        int iHidRows = 0;
        int iHidCols = 0;
        try {
            foreach (dynamic r in oSheet.UsedRange.Rows) {
                bool b = false;
                try { b = (bool)r.Hidden; } catch {}
                if (b) {
                    iHidRows++;
                }
            }
        } catch {}
        try {
            foreach (dynamic c in oSheet.UsedRange.Columns) {
                bool b = false;
                try { b = (bool)c.Hidden; } catch {}
                if (b) {
                    iHidCols++;
                }
            }
        } catch {}
        if (iHidRows > 0) {
            results.add("HiddenRows", "MSAC", "Object not inline", sSheet, iHidRows + " hidden row(s)",
                "Hidden rows disrupt the continuous data sequence and may be skipped by screen readers.",
                "Unhide rows. Consider using grouping (Data > Group) or filters instead.");
        }
        if (iHidCols > 0) {
            results.add("HiddenColumns", "MSAC", "Object not inline", sSheet, iHidCols + " hidden column(s)",
                "Hidden columns disrupt the continuous data sequence and may be skipped by screen readers.",
                "Unhide columns. Consider moving the data to a separate sheet or documenting hidden columns with a visible note.");
        }
    }
}

// ===========================================================================
// Format module: PPTX
// ===========================================================================
static class pptxModule {
    const int iPpPlaceholderTitle       = 1;
    const int iPpPlaceholderCenterTitle = 3;
    const int iMsoShapeTypePicture      = 13;
    const int iMsoShapeTypeLinkedPic    = 11;
    const int iMsoShapeTypeOLE          = 7;
    const int iMsoShapeTypeLine         = 9;
    const int iMsoShapeTypeConnector    = 10;

    static dynamic oPpt          = null;
    static dynamic oPresentation = null;

    public static bool open(string sPath) {
        try {
            oPpt             = comHelper.createApp("PowerPoint.Application");
            oPpt.Visible     = true; // Required — PowerPoint does not support headless mode
            oPpt.WindowState = 2;    // ppWindowMinimized
            oPresentation    = oPpt.Presentations.Open(sPath, true, false, false);
            return true;
        } catch (Exception ex) {
            Console.WriteLine("  ERROR: " + ex.Message);
            quit();
            return false;
        }
    }

    public static void quit() {
        try { if (oPresentation != null) oPresentation.Close(); } catch {}
        try { if (oPpt          != null) oPpt.Quit();           } catch {}
        object oRawPres = oPresentation; oPresentation = null;
        object oRawPpt  = oPpt;          oPpt          = null;
        try { if (oRawPres != null) Marshal.ReleaseComObject(oRawPres); } catch {}
        try { if (oRawPpt  != null) Marshal.ReleaseComObject(oRawPpt);  } catch {}
    }

    public static void checkAll(string sFilePath) {
        presentationMeta(sFilePath);
        speakerNotes();
        int iNum = 0;
        try {
            foreach (dynamic oSlide in oPresentation.Slides) {
                iNum++;
                string sLabel = slideLabel(oSlide, iNum);
                slideTitle(oSlide, sLabel);
                slideShapes(oSlide, sLabel);
                slideTransition(oSlide, sLabel);
                slideAnimations(oSlide, sLabel);
                slideNotes(oSlide, sLabel);
            }
        } catch {}
        readingOrder();
    }

    static void presentationMeta(string sFilePath) {
        string sTitle = "";
        try { sTitle = (oPresentation.BuiltInDocumentProperties["Title"].Value ?? "").ToString().Trim(); } catch {}
        if (sTitle == "") {
            results.add("MissingPresentationTitle", "MSAC", "Missing document properties", "(Presentation)", sFilePath,
                "The presentation Title property is not set.",
                "Go to File > Info > Properties and enter a descriptive title.");
        }
        string sAuthor = "";
        try { sAuthor = (oPresentation.BuiltInDocumentProperties["Author"].Value ?? "").ToString().Trim(); } catch {}
        if (sAuthor == "") {
            results.add("MissingPresentationAuthor", "AXE", "Missing document properties", "(Presentation)", sFilePath,
                "The presentation Author property is not set.",
                "Go to File > Info > Properties and enter the author name.");
        }
    }

    static void speakerNotes() {
        int iTotal     = 0;
        int iWithNotes = 0;
        try {
            iTotal = (int)oPresentation.Slides.Count;
            foreach (dynamic oSlide in oPresentation.Slides) {
                string sN = "";
                try { sN = oSlide.NotesPage.Shapes[2].TextFrame.TextRange.Text.ToString().Trim(); } catch {}
                if (sN != "") {
                    iWithNotes++;
                }
            }
        } catch {}
        if (iTotal > 0 && iWithNotes == 0) {
            results.add("NoSpeakerNotes", "AXE", "Missing alternative text", "(Presentation)", "All " + iTotal + " slide(s)",
                "No slides have speaker notes. For distributed presentations, notes provide essential context for screen reader users.",
                "Add speaker notes via View > Notes. Describe visual content and anything conveyed only through visuals.");
        }
    }

    static void slideTitle(dynamic oSlide, string sLabel) {
        bool bHasTitle = false;
        try {
            foreach (dynamic oShape in oSlide.Shapes) {
                bool bIsP  = false;
                int iPType = 0;
                try { bIsP  = (bool)oShape.IsPlaceholder; } catch {}
                if (!bIsP) {
                    continue;
                }
                try { iPType = (int)oShape.PlaceholderFormat.Type; } catch {}
                if (iPType == iPpPlaceholderTitle || iPType == iPpPlaceholderCenterTitle) {
                    string sT = "";
                    try { sT = oShape.TextFrame.TextRange.Text.ToString().Trim(); } catch {}
                    if (sT != "") {
                        bHasTitle = true;
                        break;
                    }
                }
            }
        } catch {}
        if (!bHasTitle) {
            results.add("SlideMissingTitle", "MSAC", "Object has no title or name", sLabel, "Title placeholder",
                "Slide has no title or the title placeholder is empty. Screen readers announce slide titles as navigation landmarks.",
                "Click the title placeholder and enter a concise, descriptive title.");
        }
    }

    static void slideShapes(dynamic oSlide, string sLabel) {
        try {
            foreach (dynamic oShape in oSlide.Shapes) {
                string sName  = "";
                int iType     = 0;
                bool bIsP     = false;
                int iPType    = 0;
                try { sName   = oShape.Name.ToString(); } catch {}
                try { iType   = (int)oShape.Type; } catch {}
                try { bIsP    = (bool)oShape.IsPlaceholder; } catch {}
                if (bIsP) {
                    try { iPType = (int)oShape.PlaceholderFormat.Type; } catch {}
                }
                bool bIsTitle = (iPType == iPpPlaceholderTitle || iPType == iPpPlaceholderCenterTitle);

                if (iType != iMsoShapeTypeLine && iType != iMsoShapeTypeConnector) {
                    bool bNeedsAlt = (iType == iMsoShapeTypePicture || iType == iMsoShapeTypeLinkedPic || iType == iMsoShapeTypeOLE);
                    if (!bNeedsAlt && !bIsP) {
                        bNeedsAlt = true;
                    }
                    if (bNeedsAlt && !bIsTitle) {
                        string sAlt = "";
                        try { sAlt = (oShape.AlternativeText ?? "").ToString().Trim(); } catch {}
                        if (sAlt == "") {
                            results.add("MissingAltText", "MSAC", "Missing alternative text", sLabel, "Shape: " + sName,
                                "A shape or image has no alternative text.",
                                "Right-click the shape, choose Edit Alt Text, and write a concise description. If purely decorative, check the Decorative checkbox.");
                        }
                    }
                }

                bool bHasChart = false;
                try { bHasChart = (bool)oShape.HasChart; } catch {}
                if (bHasChart) {
                    try { shared.chartTitleCheck(sLabel, sName, oShape.Chart); } catch {}
                }

                bool bHasText = false;
                try { bHasText = (bool)oShape.HasTextFrame; } catch {}
                if (!bHasText) {
                    continue;
                }
                try {
                    foreach (dynamic oPara in oShape.TextFrame.TextRange.Paragraphs()) {
                        string sText = "";
                        try { sText = oPara.Text.ToString().Trim(); } catch {}
                        if (sText == "") {
                            continue;
                        }
                        float nSize = 0;
                        try { nSize = (float)oPara.Font.Size; } catch {}
                        if (nSize > 0 && nSize < 18) {
                            results.add("SmallFontSize", "MSAC", "Sufficient contrast", sLabel,
                                shared.trunc(sText) + " (" + nSize + "pt)",
                                "Text is " + nSize + "pt, which may be too small for users with low vision.",
                                "Increase font size to at least 18pt for body text and 24pt or larger for key content.");
                        }
                        if (sText.Length > 4 && sText == sText.ToUpper() && sText != sText.ToLower()) {
                            results.add("AllCapsText", "MSAC", "Text", sLabel, shared.trunc(sText),
                                "Text paragraph is typed entirely in capitals. TTS may read each letter individually.",
                                "Use mixed-case text. Apply an AllCaps text effect if the style requires it.");
                        }
                        if (sText.Length > 1) {
                            string sFirst    = sText.Substring(0, 1);
                            int iBulletType  = 0;
                            try { iBulletType = (int)oPara.ParagraphFormat.Bullet.Type; } catch {}
                            foreach (string b in shared.aBullets) {
                                if (sFirst == b && iBulletType == 0) {
                                    results.add("FakeListBullet", "MSAC", "List not used correctly", sLabel, shared.trunc(sText),
                                        "Text uses a manually typed bullet instead of a PowerPoint list style.",
                                        "Apply a bullet list style from the Home tab Paragraph group.");
                                }
                            }
                        }
                    }
                } catch {}
            }
        } catch {}
    }

    static void slideTransition(dynamic oSlide, string sLabel) {
        try {
            dynamic oT  = oSlide.SlideShowTransition;
            bool bAuto  = false;
            float nTime = 0;
            try { bAuto = (bool)oT.AdvanceOnTime; } catch {}
            try { nTime = (float)oT.AdvanceTime; } catch {}
            if (bAuto && nTime > 0 && nTime < 3) {
                results.add("FastAutoAdvance", "AXE", "Timing", sLabel, "Auto-advance: " + nTime + "s",
                    "Slide auto-advances in less than 3 seconds, leaving insufficient time for screen reader announcement (WCAG 2.2.1).",
                    "Increase auto-advance time to at least 5 seconds, or disable it. Go to Transitions > Advance Slide > After.");
            }
            int iSpeed = 0;
            try { iSpeed = (int)oT.Speed; } catch {}
            if (iSpeed == 3) {
                results.add("FastTransitionSpeed", "AXE", "Motion", sLabel, "Transition speed: Fast",
                    "Fast transition speed can trigger symptoms in users with vestibular disorders (WCAG 2.3.3).",
                    "Set transition speed to Medium or Slow in the Transitions tab Duration field.");
            }
        } catch {}
    }

    static void slideAnimations(dynamic oSlide, string sLabel) {
        try {
            dynamic oSeq = oSlide.TimeLine.MainSequence;
            int iCount   = 0;
            try { iCount = (int)oSeq.Count; } catch {}
            for (int i = 1; i <= iCount; i++) {
                try {
                    dynamic oEffect = oSeq[i];
                    int iRepeat = 0;
                    try { iRepeat = (int)oEffect.Timing.RepeatCount; } catch {}
                    if (iRepeat > 1 || iRepeat == -1) {
                        results.add("RepeatingAnimation", "AXE", "Motion", sLabel, "Animation " + i + " repeats",
                            "A repeating or looping animation violates WCAG 2.2.2 (Pause, Stop, Hide).",
                            "Set the animation repeat count to 1 in the Effect Options dialog.");
                    }
                    int iTrigger = 0;
                    float nDelay = 0;
                    try { iTrigger = (int)oEffect.Timing.TriggerType; } catch {}
                    try { nDelay   = (float)oEffect.Timing.TriggerDelayTime; } catch {}
                    if ((iTrigger == 2 || iTrigger == 3) && nDelay < 1 && iCount > 3) {
                        results.add("RapidAutoAnimation", "AXE", "Motion", sLabel, "Animation " + i + ": " + nDelay + "s delay",
                            "Multiple animations trigger automatically with very short delays, overwhelming screen reader users.",
                            "Increase delays to at least 1 second per element, or switch to On Click triggers.");
                    }
                } catch {}
            }
        } catch {}
    }

    static void slideNotes(dynamic oSlide, string sLabel) {
        bool bHasUncaptioned = false;
        try {
            foreach (dynamic oShape in oSlide.Shapes) {
                int iType      = 0;
                bool bHasChart = false;
                try { iType    = (int)oShape.Type; } catch {}
                try { bHasChart = (bool)oShape.HasChart; } catch {}
                if (iType == iMsoShapeTypePicture || iType == iMsoShapeTypeLinkedPic || bHasChart) {
                    string sAlt = "";
                    try { sAlt = (oShape.AlternativeText ?? "").ToString().Trim(); } catch {}
                    if (sAlt == "") {
                        bHasUncaptioned = true;
                    }
                }
            }
        } catch {}
        if (!bHasUncaptioned) {
            return;
        }
        string sNotes = "";
        try { sNotes = oSlide.NotesPage.Shapes[2].TextFrame.TextRange.Text.ToString().Trim(); } catch {}
        if (sNotes == "") {
            results.add("VisualContentWithoutNotes", "AXE", "Missing alternative text", sLabel,
                "Images or charts without alt text and no speaker notes",
                "This slide has visual content with no alt text and no speaker notes, making it entirely inaccessible.",
                "Add alt text to each image and chart, and add speaker notes describing the key visual content.");
        }
    }

    static void readingOrder() {
        int iNum = 0;
        try {
            foreach (dynamic oSlide in oPresentation.Slides) {
                iNum++;
                string sLabel = slideLabel(oSlide, iNum);
                int iCount    = 0;
                try { iCount  = (int)oSlide.Shapes.Count; } catch {}
                if (iCount < 2) {
                    continue;
                }
                dynamic oBack = null;
                try { oBack   = oSlide.Shapes[iCount]; } catch {}
                if (oBack == null) {
                    continue;
                }
                bool bBackIsTitle = false;
                try {
                    bool bIsP = (bool)oBack.IsPlaceholder;
                    if (bIsP) {
                        int t = (int)oBack.PlaceholderFormat.Type;
                        bBackIsTitle = (t == iPpPlaceholderTitle || t == iPpPlaceholderCenterTitle);
                    }
                } catch {}
                if (!bBackIsTitle) {
                    results.add("TitleNotFirstInReadingOrder", "AXE", "Object not inline", sLabel, "Title z-order",
                        "The title placeholder is not at the back of the z-order. PowerPoint screen readers traverse shapes back-to-front.",
                        "Open the Selection Pane (Home > Arrange > Selection Pane) and drag the title to the bottom of the list.");
                }
            }
        } catch {}
    }

    static string slideLabel(dynamic oSlide, int iNum) {
        string sName = "";
        try { sName = oSlide.Name.ToString(); } catch {}
        return (sName != "" && sName != "Slide " + iNum)
            ? "Slide " + iNum + " (" + sName + ")"
            : "Slide " + iNum;
    }
}

// ===========================================================================
// Format module: MD (Pandoc Markdown) — no COM, regex-based text analysis
// ===========================================================================
static class mdModule {
    const int iMaxParasBetweenHeadings = 20;
    const int iMaxHeadingLength        = 120;
    const int iMinAltTextLength        = 3;

    static string[] aLines;
    static Dictionary<string, string> dMeta;
    static int iYamlEndLine;
    static Regex reFence = new Regex(@"^(`{3,}|~{3,})");

    public static bool open(string sPath) {
        aLines       = File.ReadAllText(sPath, Encoding.UTF8).Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        dMeta        = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        iYamlEndLine = 0;
        parseYaml();
        return true;
    }

    public static void quit() {}  // no COM to release

    public static void checkAll(string sFilePath) {
        metadata(sFilePath);
        headings();
        images();
        links();
        tables();
        lists();
        emphasis();
        codeBlocks();
        rawHtml();
        lineIssues();
    }

    static void parseYaml() {
        if (aLines.Length < 2 || aLines[0].Trim() != "---") {
            return;
        }
        for (int i = 1; i < aLines.Length; i++) {
            string sT = aLines[i].Trim();
            if (sT == "---" || sT == "...") {
                iYamlEndLine = i + 1;
                break;
            }
            int iC = aLines[i].IndexOf(':');
            if (iC > 0) {
                string sKey = aLines[i].Substring(0, iC).Trim().ToLower();
                string sVal = aLines[i].Substring(iC + 1).Trim().Trim('"').Trim('\'');
                if (!dMeta.ContainsKey(sKey)) {
                    dMeta[sKey] = sVal;
                }
            }
        }
    }

    static bool inYaml(int iLn)      { return iLn <= iYamlEndLine; }
    static bool hasMeta(string sKey) { return dMeta.ContainsKey(sKey) && dMeta[sKey] != ""; }

    static void metadata(string sFilePath) {
        if (iYamlEndLine == 0) {
            add("NoYamlFrontMatter", 1, "MSAC", "Missing document properties", sFilePath,
                "The file has no YAML front matter block. Without front matter, Pandoc cannot set document title, language, or author.",
                "Add a YAML front matter block at the very top starting and ending with ---. Include at minimum: title and lang.");
            return;
        }
        if (!hasMeta("title")) {
            add("MissingTitle", 1, "MSAC", "Missing document properties", sFilePath,
                "No 'title' field found in YAML front matter.",
                "Add 'title: \"My Document Title\"' to the YAML front matter.");
        }
        if (!hasMeta("lang")) {
            add("MissingLanguage", 1, "MSAC", "Missing document language", sFilePath,
                "No 'lang' field in YAML front matter. Screen readers use this to select the correct TTS voice.",
                "Add 'lang: en-US' (or appropriate BCP 47 tag) to the YAML front matter.");
        }
        if (!hasMeta("author")) {
            add("MissingAuthor", 1, "AXE", "Missing document properties", sFilePath,
                "No 'author' field in YAML front matter.",
                "Add 'author: \"Name\"' to the YAML front matter.");
        }
        if (!hasMeta("description")) {
            add("MissingDescription", 1, "AXE", "Missing document properties", sFilePath,
                "No 'description' field in YAML front matter.",
                "Add 'description: \"Brief summary\"' to the YAML front matter.");
        }
    }

    // Convenience overload for MD: location is the line number as a string
    static void add(string ruleId, int iLn, string source, string category, string context, string message, string remediation) {
        results.add(ruleId, source, category, "Line " + iLn, context, message, remediation);
    }

    static void headings() {
        int iLastLv      = 0;
        int iParasSince  = 0;
        bool bLongFired  = false;
        bool bFoundFirst = false;
        var lHdgs = new List<string>();
        bool bInFence = false;
        var reAtx  = new Regex(@"^(#{1,6})\s+(.+)$");
        var reH1ul = new Regex(@"^=+\s*$");
        var reH2ul = new Regex(@"^-+\s*$");

        for (int i = 0; i < aLines.Length; i++) {
            int iLn = i + 1;
            if (inYaml(iLn)) {
                continue;
            }
            string sLine = aLines[i];
            string sT    = sLine.Trim();
            if (reFence.IsMatch(sT)) {
                bInFence = !bInFence;
                continue;
            }
            if (bInFence) {
                continue;
            }

            int iSetextLv = 0;
            if (i > 0 && sT.Length > 0) {
                if (reH1ul.IsMatch(sT) && aLines[i-1].Trim().Length > 0) {
                    iSetextLv = 1;
                } else if (reH2ul.IsMatch(sT) && aLines[i-1].Trim().Length > 0) {
                    iSetextLv = 2;
                }
            }

            var mAtx    = reAtx.Match(sLine);
            bool bIsHdg = mAtx.Success || iSetextLv > 0;
            int iLv     = 0;
            string sHText = "";

            if (mAtx.Success) {
                iLv    = mAtx.Groups[1].Value.Length;
                sHText = Regex.Replace(mAtx.Groups[2].Value.Trim(), @"\s+#+\s*$", "").Trim();
            } else if (iSetextLv > 0) {
                iLv    = iSetextLv;
                sHText = aLines[i-1].Trim();
            }

            if (bIsHdg) {
                bFoundFirst = true;
                if (iLastLv > 0 && iLv > iLastLv + 1) {
                    add("SkippedHeadingLevel", iLn, "MSAC", "Heading issues", shared.trunc(sHText),
                        "Heading jumps from H" + iLastLv + " to H" + iLv + ", skipping a level.",
                        "Change this heading to H" + (iLastLv + 1) + " so levels are sequential.");
                }
                if (sHText == "") {
                    add("EmptyHeading", iLn, "MSAC", "Heading issues", "(empty)",
                        "Heading marker is present but contains no text.",
                        "Add descriptive text after the heading marker, or remove the marker.");
                }
                string sKey = iLv + "|" + sHText.ToLower().Trim();
                if (lHdgs.Contains(sKey)) {
                    add("DuplicateHeadingText", iLn, "AXE", "Heading issues", shared.trunc(sHText),
                        "Two H" + iLv + " headings share identical text, creating ambiguous navigation.",
                        "Make each heading uniquely descriptive of its section.");
                } else {
                    lHdgs.Add(sKey);
                }
                if (sHText.Length > iMaxHeadingLength) {
                    add("HeadingTooLong", iLn, "MSAC", "Object has no title or name", shared.trunc(sHText) + "...",
                        "Heading is " + sHText.Length + " characters, which is unusually long for a navigation landmark.",
                        "Shorten the heading to under 80 characters.");
                }
                if (sHText.Length > 0) {
                    char cLast = sHText[sHText.Length - 1];
                    if (cLast == '.' || cLast == ';' || cLast == ':') {
                        add("HeadingEndsWithPunctuation", iLn, "MSAC", "Object has no title or name", shared.trunc(sHText),
                            "Heading ends with '" + cLast + "', causing TTS to insert an unnatural pause.",
                            "Remove the trailing punctuation from the heading.");
                    }
                }
                iLastLv     = iLv;
                iParasSince = 0;
                bLongFired  = false;
            } else {
                if (sT.Length > 0 && !sT.StartsWith("|") && !sT.StartsWith("    ")) {
                    iParasSince++;
                    if (iParasSince > iMaxParasBetweenHeadings && !bLongFired && iLastLv == 0) {
                        add("LongSectionWithoutHeading", iLn, "MSAC", "Heading issues", "(document start)",
                            "More than " + iMaxParasBetweenHeadings + " content lines appear before any heading.",
                            "Add a top-level heading near the start of the document and headings throughout to mark major sections.");
                        bLongFired = true;
                    }
                }
            }
        }

        if (!bFoundFirst && aLines.Length > 30) {
            results.add("NoHeadings", "MSAC", "Heading issues", "Line 1", "(document)",
                "Document has no headings. Screen reader users navigate long documents primarily by heading list.",
                "Add ATX-style headings (# H1, ## H2) to mark major sections.");
        }
    }

    static void images() {
        var reImg     = new Regex(@"!\[([^\]]*)\]\(([^)]*)\)");
        var reImgRef  = new Regex(@"!\[([^\]]*)\]\[([^\]]*)\]");
        var reHtmlImg = new Regex(@"<img\b[^>]*>", RegexOptions.IgnoreCase);
        var reAlt     = new Regex(@"\balt\s*=\s*(?:""([^""]*)""|'([^']*)'|(\S+))", RegexOptions.IgnoreCase);
        bool bInFence = false;

        for (int i = 0; i < aLines.Length; i++) {
            int iLn = i + 1;
            if (inYaml(iLn)) {
                continue;
            }
            string sLine = aLines[i];
            string sT    = sLine.Trim();
            if (reFence.IsMatch(sT)) {
                bInFence = !bInFence;
                continue;
            }
            if (bInFence) {
                continue;
            }

            foreach (Match m in reImg.Matches(sLine)) {
                string sAlt = m.Groups[1].Value.Trim();
                string sUrl = m.Groups[2].Value.Trim();
                int iTs     = sUrl.IndexOf('"');
                if (iTs > 0) {
                    sUrl = sUrl.Substring(0, iTs).Trim();
                }
                if (sAlt == "") {
                    add("MissingAltText", iLn, "MSAC", "Missing alternative text", shared.trunc(sUrl),
                        "Image has empty alt text.",
                        "Add descriptive alt text inside the square brackets: ![A description](url).");
                } else if (sAlt.Length < iMinAltTextLength) {
                    add("AltTextTooShort", iLn, "MSAC", "Missing alternative text", shared.trunc(sAlt),
                        "Image alt text is very short (\"" + sAlt + "\") and unlikely to convey meaningful information.",
                        "Expand the alt text to describe the image content and its purpose.");
                } else if (Regex.IsMatch(sAlt, @"\.[a-zA-Z]{2,4}$") && !sAlt.Contains(" ")) {
                    add("AltTextIsFilename", iLn, "MSAC", "Missing alternative text", shared.trunc(sAlt),
                        "Image alt text appears to be a filename rather than a meaningful description.",
                        "Replace the filename with a description of what the image shows.");
                } else if (sAlt.ToLower().StartsWith("image of") || sAlt.ToLower().StartsWith("picture of") || sAlt.ToLower().StartsWith("photo of")) {
                    add("AltTextRedundantPrefix", iLn, "MSAC", "Missing alternative text", shared.trunc(sAlt),
                        "Alt text begins with a redundant phrase. Screen readers already announce the element is an image.",
                        "Remove the redundant prefix and begin the alt text with the actual description.");
                }
            }

            foreach (Match m in reImgRef.Matches(sLine)) {
                if (m.Groups[1].Value.Trim() == "") {
                    add("MissingAltText", iLn, "MSAC", "Missing alternative text", "ref: " + m.Groups[2].Value.Trim(),
                        "Reference-style image has empty alt text.",
                        "Add descriptive alt text: ![A description][ref]");
                }
            }

            foreach (Match m in reHtmlImg.Matches(sLine)) {
                var mAlt    = reAlt.Match(m.Value);
                string sAlt = mAlt.Success
                    ? (mAlt.Groups[1].Value + mAlt.Groups[2].Value + mAlt.Groups[3].Value).Trim()
                    : "";
                if (!mAlt.Success || sAlt == "") {
                    add("MissingAltText", iLn, "MSAC", "Missing alternative text", shared.trunc(m.Value),
                        "Raw HTML img tag is missing an alt attribute or has empty alt.",
                        "Add alt=\"A description\" to the img tag. Use alt=\"\" only for decorative images.");
                }
            }
        }
    }

    static void links() {
        var reLink    = new Regex(@"\[([^\]]*)\]\(([^)]*)\)");
        var reLinkRef = new Regex(@"(?<!!)\[([^\]]+)\]\[([^\]]*)\]");
        var reBareUrl = new Regex(@"(?<![(\[<])(https?://\S+)", RegexOptions.IgnoreCase);
        bool bInFence = false;

        for (int i = 0; i < aLines.Length; i++) {
            int iLn = i + 1;
            if (inYaml(iLn)) {
                continue;
            }
            string sLine = aLines[i];
            string sT    = sLine.Trim();
            if (reFence.IsMatch(sT)) {
                bInFence = !bInFence;
                continue;
            }
            if (bInFence) {
                continue;
            }

            foreach (Match m in reLink.Matches(sLine)) {
                if (m.Index > 0 && sLine[m.Index - 1] == '!') {
                    continue;
                }
                string sText = m.Groups[1].Value.Trim();
                string sUrl  = m.Groups[2].Value.Trim();
                int iTs      = sUrl.IndexOf('"');
                if (iTs > 0) {
                    sUrl = sUrl.Substring(0, iTs).Trim();
                }
                checkLinkText(sText, sUrl, iLn);
            }

            foreach (Match m in reLinkRef.Matches(sLine)) {
                checkLinkText(m.Groups[1].Value.Trim(), "[" + m.Groups[2].Value.Trim() + "]", iLn);
            }

            foreach (Match m in reBareUrl.Matches(sLine)) {
                bool bInside = false;
                int idx      = m.Index;
                for (int j = idx - 1; j >= 0 && j >= idx - 3; j--) {
                    if (sLine[j] == '(' || sLine[j] == '[') {
                        bInside = true;
                        break;
                    }
                }
                if (!bInside) {
                    add("BareUrl", iLn, "AXE", "Hyperlink", shared.trunc(m.Value),
                        "A bare URL appears in text without descriptive link text. Screen readers read the URL character by character.",
                        "Wrap the URL in a Markdown link: [Descriptive text](https://url).");
                }
            }
        }
    }

    static void checkLinkText(string sText, string sTarget, int iLn) {
        if (sText == "") {
            add("EmptyLinkText", iLn, "AXE", "Hyperlink", shared.trunc(sTarget),
                "Link has no visible text. Screen readers read the raw URL or produce no output.",
                "Add descriptive text between the square brackets.");
            return;
        }
        foreach (string v in shared.aVague) {
            if (sText.ToLower() == v) {
                add("NonDescriptiveLinkText", iLn, "AXE", "Hyperlink", "\"" + sText + "\"",
                    "Link text \"" + sText + "\" is non-descriptive and does not indicate the destination.",
                    "Replace with text that describes where the link goes.");
                return;
            }
        }
        if (sText.StartsWith("http://") || sText.StartsWith("https://") || sText.StartsWith("www.")) {
            add("UrlAsLinkText", iLn, "AXE", "Hyperlink", shared.trunc(sText),
                "Link text is a raw URL, which screen readers read character by character.",
                "Replace the URL with a description of the link destination.");
        }
    }

    static void tables() {
        var reSep     = new Regex(@"^\|?[\s\-:]+(\|[\s\-:]+)+\|?\s*$");
        var reRow     = new Regex(@"^\|.+\|?\s*$");
        bool bInFence = false;
        int iTableStart = -1;
        int iTableNum   = 0;
        int iColCount   = 0;

        for (int i = 0; i < aLines.Length; i++) {
            int iLn = i + 1;
            if (inYaml(iLn)) {
                continue;
            }
            string sLine = aLines[i];
            string sT    = sLine.Trim();
            if (reFence.IsMatch(sT)) {
                bInFence = !bInFence;
                continue;
            }
            if (bInFence) {
                continue;
            }

            bool bIsRow = reRow.IsMatch(sT);
            bool bIsSep = reSep.IsMatch(sT);

            if (bIsRow && iTableStart < 0 && i + 1 < aLines.Length && reSep.IsMatch(aLines[i+1].Trim())) {
                iTableStart = iLn;
                iTableNum++;
                string[] aCells = sT.Trim('|').Split('|');
                iColCount = aCells.Length;
                int iCol  = 0;
                foreach (string sCell in aCells) {
                    iCol++;
                    if (stripMd(sCell).Trim() == "") {
                        add("BlankTableHeader", iLn, "MSAC", "Missing table header", "Table " + iTableNum + " Col " + iCol,
                            "Table column header cell is blank.",
                            "Add a descriptive label to each column header cell.");
                    }
                }
                continue;
            }

            if (iTableStart > 0 && bIsSep) {
                continue;
            }

            if (iTableStart > 0 && bIsRow) {
                string[] aCells = sT.Trim('|').Split('|');
                int iCol        = 0;
                foreach (string sCell in aCells) {
                    iCol++;
                    if (stripMd(sCell).Trim() == "") {
                        add("EmptyTableCell", iLn, "MSAC", "Blank cells used for formatting", "Table " + iTableNum + " Col " + iCol,
                            "Table data cell is empty.",
                            "Add content, or use a dash or N/A to indicate intentionally missing data.");
                    }
                }
                if (aCells.Length != iColCount) {
                    add("TableColumnCountMismatch", iLn, "AXE", "Complex table", "Table " + iTableNum + " row at line " + iLn,
                        "Row has " + aCells.Length + " cells but the header has " + iColCount + " columns.",
                        "Ensure every row has the same number of pipe-delimited cells as the header row.");
                }
                continue;
            }

            if (iTableStart > 0 && !bIsRow) {
                iTableStart = -1;
                iColCount   = 0;
            }
        }
    }

    static void lists() {
        var reFakeNum       = new Regex(@"^\s{0,3}(\d+)[.)]\s+\S");
        var reFakeMidBullet = new Regex(@"\s[•·‣⁃]\s");
        bool bInFence = false;

        for (int i = 0; i < aLines.Length; i++) {
            int iLn = i + 1;
            if (inYaml(iLn)) {
                continue;
            }
            string sLine = aLines[i];
            string sT    = sLine.Trim();
            if (reFence.IsMatch(sT)) {
                bInFence = !bInFence;
                continue;
            }
            if (bInFence) {
                continue;
            }

            if (reFakeMidBullet.IsMatch(sLine)) {
                add("FakeInlineBullet", iLn, "MSAC", "List not used correctly", shared.trunc(sT),
                    "Line contains a Unicode bullet character used as an inline list marker.",
                    "Use proper Markdown list items: start each item on its own line with - or * followed by a space.");
            }

            if (!Regex.IsMatch(sLine, @"^\s{0,3}[-*+]\s") && reFakeNum.IsMatch(sLine)
                    && i > 0 && reFakeNum.IsMatch(aLines[i-1])) {
                add("FakeNumberedList", iLn, "MSAC", "List not used correctly", shared.trunc(sT),
                    "Lines appear to be a manually numbered list outside proper Markdown ordered list syntax.",
                    "Use Markdown ordered list syntax: start each item with '1.' followed by a space.");
            }
        }
    }

    static void emphasis() {
        var reAllCaps  = new Regex(@"(?<![`\w])[A-Z]{7,}(?![`\w])");
        var reCodeSpan = new Regex(@"`[^`]+`");
        var reFullBold = new Regex(@"^\*{2}.{10,}\*{2}$|^_{2}.{10,}_{2}$");
        bool bInFence  = false;

        for (int i = 0; i < aLines.Length; i++) {
            int iLn = i + 1;
            if (inYaml(iLn)) {
                continue;
            }
            string sLine = aLines[i];
            string sT    = sLine.Trim();
            if (reFence.IsMatch(sT)) {
                bInFence = !bInFence;
                continue;
            }
            if (bInFence) {
                continue;
            }

            string sNoCode = reCodeSpan.Replace(sLine, " ");
            foreach (Match m in reAllCaps.Matches(sNoCode)) {
                add("AllCapsText", iLn, "MSAC", "Text", "\"" + m.Value + "\"",
                    "Long all-caps sequence \"" + m.Value + "\" may cause TTS to read each letter individually.",
                    "Use mixed case. Reserve all-caps for true acronyms of six characters or fewer.");
            }

            if (reFullBold.IsMatch(sT) && sT.Length > 20) {
                add("EntireLineBolded", iLn, "AXE", "Text", shared.trunc(sT),
                    "The entire line is bold. Bolding long passages reduces the emphasis signal for screen reader users.",
                    "Use bold sparingly for key terms. If the intent is a heading, use a heading marker instead.");
            }
        }
    }

    static void codeBlocks() {
        var reFenceOpen = new Regex(@"^(`{3,}|~{3,})(\w*)");
        bool bInFence   = false;

        for (int i = 0; i < aLines.Length; i++) {
            int iLn = i + 1;
            if (inYaml(iLn)) {
                continue;
            }
            string sT = aLines[i].Trim();
            var mF    = reFenceOpen.Match(sT);
            if (mF.Success) {
                if (!bInFence) {
                    bInFence = true;
                    if (mF.Groups[2].Value.Trim() == "") {
                        add("CodeBlockMissingLanguage", iLn, "AXE", "Code", "Fenced block at line " + iLn,
                            "Fenced code block has no language specifier. Screen reader users receive no cue about what kind of code they are reading.",
                            "Add a language identifier after the opening fence: ```python, ```bash, ```json, etc.");
                    }
                } else {
                    bInFence = false;
                }
                continue;
            }

            if (!bInFence && aLines[i].StartsWith("    ") && aLines[i].Trim().Length > 0) {
                if (i > 0 && !aLines[i-1].StartsWith("    ") && aLines[i-1].Trim().Length > 0) {
                    add("IndentedCodeBlock", iLn, "AXE", "Code", shared.trunc(aLines[i].Trim()),
                        "Indented code block (4-space indent) cannot carry a language specifier.",
                        "Convert to a fenced code block using ``` with a language identifier.");
                }
            }
        }
    }

    static void rawHtml() {
        var reTag = new Regex(@"<([a-zA-Z][a-zA-Z0-9]*)\b[^>]*>", RegexOptions.IgnoreCase);
        bool bInFence = false;

        for (int i = 0; i < aLines.Length; i++) {
            int iLn = i + 1;
            if (inYaml(iLn)) {
                continue;
            }
            string sLine = aLines[i];
            if (reFence.IsMatch(sLine.Trim())) {
                bInFence = !bInFence;
                continue;
            }
            if (bInFence) {
                continue;
            }

            foreach (Match m in reTag.Matches(sLine)) {
                string sTag = m.Groups[1].Value.ToLower();
                if (sTag == "table") {
                    add("RawHtmlTable", iLn, "MSAC", "Missing table header", shared.trunc(m.Value),
                        "Raw HTML table used. HTML tables require explicit th scope headers and a caption to be accessible.",
                        "Use Pandoc pipe table syntax, or ensure the HTML table includes th scope headers and a caption.");
                } else if (sTag == "font" || sTag == "center") {
                    add("RawHtmlPresentational", iLn, "MSAC", "Use of color alone", shared.trunc(m.Value),
                        "Raw HTML presentational element <" + sTag + "> conveys no semantics and is not supported in non-HTML Pandoc output formats.",
                        "Replace with Markdown formatting or a Pandoc-native span/div.");
                } else if (sTag == "marquee" || sTag == "blink") {
                    add("MovingContent", iLn, "AXE", "Motion", shared.trunc(m.Value),
                        "Raw HTML <" + sTag + "> causes moving or blinking content, violating WCAG 2.2.2.",
                        "Remove the element entirely. Moving and blinking text has no accessible equivalent.");
                } else if (sTag == "div" || sTag == "span") {
                    add("RawHtmlGenericContainer", iLn, "AXE", "Use of color alone", shared.trunc(m.Value),
                        "Raw HTML <" + sTag + "> container is silently ignored by Pandoc when converting to non-HTML formats.",
                        "Use Pandoc's native fenced div (::: {.class}) or bracketed span ([text]{.class}) syntax instead.");
                }
            }
        }
    }

    static void lineIssues() {
        bool bInFence     = false;
        var reTrail       = new Regex(@"  +$");
        int iConsecBlanks = 0;

        for (int i = 0; i < aLines.Length; i++) {
            int iLn = i + 1;
            if (inYaml(iLn)) {
                continue;
            }
            string sLine = aLines[i];
            string sT    = sLine.Trim();
            if (reFence.IsMatch(sT)) {
                bInFence      = !bInFence;
                iConsecBlanks = 0;
                continue;
            }
            if (bInFence) {
                iConsecBlanks = 0;
                continue;
            }

            if (sT == "") {
                iConsecBlanks++;
                if (iConsecBlanks == 3) {
                    add("ExcessiveBlankLines", iLn, "MSAC", "Repeated blank characters", "Line " + iLn,
                        "Three or more consecutive blank lines used for visual spacing.",
                        "Remove extra blank lines. Use a single blank line to separate paragraphs.");
                }
            } else {
                iConsecBlanks = 0;
            }

            if (reTrail.IsMatch(sLine)) {
                string sTr = sLine.Substring(sLine.TrimEnd().Length);
                if (sTr.Length > 2) {
                    add("ExcessiveTrailingSpaces", iLn, "MSAC", "Repeated blank characters",
                        "Line " + iLn + " (" + sTr.Length + " trailing spaces)",
                        "Line has " + sTr.Length + " trailing spaces. Exactly two trailing spaces are an intentional hard line break; more than two is almost always unintentional.",
                        "Remove the extra trailing spaces.");
                }
            }

            if (Regex.IsMatch(sLine, @"<br\s*/?>", RegexOptions.IgnoreCase)) {
                add("RawBrTag", iLn, "MSAC", "Repeated blank characters", shared.trunc(sT),
                    "Raw HTML br tag used for a line break, reducing portability to non-HTML Pandoc output formats.",
                    "End the line with two trailing spaces or a backslash for a portable hard line break.");
            }
        }
    }

    static string stripMd(string s) {
        s = Regex.Replace(s, @"`[^`]+`", "");
        s = Regex.Replace(s, @"\*{1,2}([^*]+)\*{1,2}", "$1");
        s = Regex.Replace(s, @"_{1,2}([^_]+)_{1,2}", "$1");
        s = Regex.Replace(s, @"\[([^\]]+)\]\([^)]+\)", "$1");
        return s;
    }
}

// ===========================================================================
// Entry point class (required by C# 7.3)
// ===========================================================================
// ===========================================================================
// Entry point. [STAThread] is REQUIRED for two reasons:
//
// 1) Office COM automation. Without STA, Word/Excel/PowerPoint COM servers
//    (in our case, in-process via dynamic) can disconnect mid-operation
//    with HRESULT 0x80010108 (RPC_E_DISCONNECTED) or 0x80010114
//    (OLE_E_OBJNOTCONNECTED). PowerPoint shape iteration and Excel
//    UsedRange.Value2 are particularly sensitive. Setting [STAThread]
//    initializes the main thread as a single-threaded apartment, which
//    is the apartment Office COM servers expect.
//
// 2) WinForms common dialogs. OpenFileDialog and FolderBrowserDialog
//    require an STA thread; otherwise ShowDialog throws or hangs.
//
// IMPORTANT for 64-bit builds: this exe is built /platform:x64. To
// automate Microsoft Office via COM, the bitness of this process must
// match the bitness of the installed Office. Modern Office (Microsoft
// 365, Office 2019+, Office 2024) is 64-bit by default; if a user has
// 32-bit Office on their machine, comHelper.createApp will fail with a
// Type.GetTypeFromProgID == null result and the error message will
// guide the user to a 32-bit rebuild.
// ===========================================================================
class extCheck {
    [STAThread]
    static int Main(string[] aArgs) { return program.run(aArgs); }
}

// ===========================================================================
// Diagnostic logger. Off by default; enabled with --log / -l or by
// checking "Log session" in the GUI dialog.
//
// When enabled, writes a UTF-8 file named extCheck.log (with BOM) to
// the output directory if one was specified (-o or the GUI Output
// directory field), or to the current directory otherwise. Each line
// is prefixed with an ISO-8601 timestamp and severity level. The log
// stream is flushed after every write so that if the process crashes,
// the log captures everything up to the crash point. Each session
// starts with a fresh file -- any prior extCheck.log in the same
// location is overwritten, so the file always reflects only the
// current run.
//
// All methods no-op silently when the log is not open, so call sites
// can log unconditionally without guarding each call.
// ===========================================================================
public static class logger
{
    private static StreamWriter writer = null;

    public static void open(string sDir = "")
    {
        if (writer != null) return;
        string sLogDir;
        try {
            if (!string.IsNullOrWhiteSpace(sDir) && Directory.Exists(sDir)) {
                sLogDir = Path.GetFullPath(sDir);
            } else {
                sLogDir = Directory.GetCurrentDirectory();
            }
        } catch {
            sLogDir = Directory.GetCurrentDirectory();
        }
        string sPath = Path.Combine(sLogDir, program.sLogFileName);
        try {
            writer = new StreamWriter(sPath, append: false, encoding: new UTF8Encoding(true));
            writer.AutoFlush = true;
        } catch (Exception ex) {
            Console.Error.WriteLine("[WARN] Could not open log file '" + sPath + "': " +
                ex.Message + ". Continuing without a log.");
            writer = null;
        }
    }

    public static void close()
    {
        if (writer == null) return;
        try {
            writer.WriteLine(stamp("INFO") + " Log closed.");
            writer.Flush();
            writer.Close();
        } catch { }
        writer = null;
    }

    public static void info(string sMsg)  { write("INFO", sMsg); }
    public static void warn(string sMsg)  { write("WARN", sMsg); }
    public static void error(string sMsg) { write("ERROR", sMsg); }
    public static void debug(string sMsg) { write("DEBUG", sMsg); }

    // Write the run header to the top of the log: program name and
    // version, the friendly run-start timestamp, and the resolved
    // parameter list. Emits raw lines (no per-line timestamp/level
    // prefix) so the header reads as a clean banner. The processing
    // notifications that follow use the standard format via
    // info/warn/error/debug.
    public static void header(string sName, string sVersion,
        List<KeyValuePair<string, string>> dParams)
    {
        if (writer == null) return;
        try {
            writer.WriteLine("=== " + sName + " " + sVersion + " ===");
            writer.WriteLine("Run on " + friendlyTime(DateTime.Now));
            if (dParams != null && dParams.Count > 0) {
                writer.WriteLine("Parameters:");
                int iPad = 0;
                foreach (var oKv in dParams)
                    if (oKv.Key.Length > iPad) iPad = oKv.Key.Length;
                foreach (var oKv in dParams)
                    writer.WriteLine("  " + oKv.Key.PadRight(iPad) + " : " + oKv.Value);
            }
            writer.WriteLine("===");
        } catch { }
    }

    public static string friendlyTime(DateTime dt)
    {
        return dt.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture) +
            " at " +
            dt.ToString("h:mm tt", CultureInfo.InvariantCulture);
    }

    private static void write(string sLevel, string sMsg)
    {
        if (writer == null) return;
        try { writer.WriteLine(stamp(sLevel) + " " + sMsg); } catch { }
    }

    private static string stamp(string sLevel)
    {
        return DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss.fff") + " [" + sLevel + "]";
    }
}

// ===========================================================================
// Configuration manager. Reads and writes a small INI file at
// %LOCALAPPDATA%\extCheck\extCheck.ini. Opt-in via -u / --use-configuration
// or the GUI Use configuration checkbox. Without that flag, extCheck
// leaves no filesystem footprint of its own. The configuration stores
// the source files string, the output directory, and the option
// checkboxes (force, view_output, log_session). The Show rules
// checkbox is intentionally NOT persisted — it's a one-shot operation.
// ===========================================================================
public static class configManager
{
    public static string getConfigDir()
    {
        string sAppData = Environment.GetFolderPath(
            Environment.SpecialFolder.LocalApplicationData);
        return Path.Combine(sAppData, program.sConfigDirName);
    }

    public static string getConfigPath()
    {
        return Path.Combine(getConfigDir(), program.sConfigFileName);
    }

    public static bool configExists()
    {
        try { return File.Exists(getConfigPath()); }
        catch { return false; }
    }

    public static void eraseAll()
    {
        string sDir = getConfigDir();
        string sPath = getConfigPath();
        try {
            if (File.Exists(sPath)) {
                File.Delete(sPath);
                logger.info("Deleted configuration file: " + sPath);
            }
        } catch (Exception ex) {
            logger.info("Could not delete configuration file " +
                sPath + ": " + ex.Message);
        }
        try {
            if (Directory.Exists(sDir)) {
                bool bEmpty = Directory.EnumerateFileSystemEntries(sDir)
                    .GetEnumerator().MoveNext() == false;
                if (bEmpty) {
                    Directory.Delete(sDir);
                    logger.info("Removed empty configuration directory: " +
                        sDir);
                }
            }
        } catch (Exception ex) {
            logger.info("Could not remove configuration directory " +
                sDir + ": " + ex.Message);
        }
    }

    public static void loadInto(List<string> lsFileArgs)
    {
        string sPath = getConfigPath();
        if (!File.Exists(sPath)) return;

        Dictionary<string, string> dVals;
        try {
            dVals = parseFile(sPath);
        } catch (Exception ex) {
            string sMsg = "Could not read configuration from:\r\n" +
                sPath + "\r\n\r\n" + ex.Message;
            Console.Error.WriteLine("[WARN] " + sMsg);
            if (program.bGuiMode) {
                try {
                    MessageBox.Show(sMsg,
                        "extCheck — Configuration not loaded",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                } catch { }
            }
            return;
        }

        if (!program.bSourceFromCli) {
            string sSaved;
            if (dVals.TryGetValue("source_files", out sSaved) &&
                !string.IsNullOrWhiteSpace(sSaved)) {
                foreach (var sArg in program.splitSourceField(sSaved))
                    lsFileArgs.Add(sArg);
            }
        }

        if (!program.bOutputDirFromCli)
            program.sOutputDir = getOrEmpty(dVals, "output_directory");
        if (!program.bForceFromCli)
            program.bForce = getBool(dVals, "force_replacements");
        if (!program.bViewOutputFromCli)
            program.bViewOutput = getBool(dVals, "view_output");
        if (!program.bLogFromCli)
            program.bLog = getBool(dVals, "log_session");
    }

    public static void save(string sSource, string sOutputDir,
        bool bForce, bool bView, bool bLog)
    {
        string sDir = getConfigDir();
        string sPath = getConfigPath();
        try {
            if (!Directory.Exists(sDir)) Directory.CreateDirectory(sDir);
            var sb = new StringBuilder();
            sb.AppendLine("; extCheck configuration");
            sb.AppendLine("; auto-written on OK-click when Use configuration was checked.");
            sb.AppendLine("; Delete this file to reset, or click Default settings in");
            sb.AppendLine("; the GUI, which also deletes the file and the extCheck folder.");
            sb.AppendLine("source_files=" + (sSource ?? ""));
            sb.AppendLine("output_directory=" + (sOutputDir ?? ""));
            sb.AppendLine("force_replacements=" + (bForce ? "1" : "0"));
            sb.AppendLine("view_output=" + (bView ? "1" : "0"));
            sb.AppendLine("log_session=" + (bLog ? "1" : "0"));
            File.WriteAllText(sPath, sb.ToString(), new UTF8Encoding(true));
            logger.info("Saved configuration to " + sPath);
        } catch (Exception ex) {
            string sMsg = "Could not save configuration to:\r\n" +
                sPath + "\r\n\r\n" + ex.Message;
            Console.Error.WriteLine("[WARN] " + sMsg);
            logger.info("Could not save configuration: " + ex.Message);
            if (program.bGuiMode) {
                try {
                    MessageBox.Show(sMsg,
                        "extCheck — Configuration not saved",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                } catch { }
            }
        }
    }

    private static Dictionary<string, string> parseFile(string sPath)
    {
        var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sLineRaw in File.ReadAllLines(sPath)) {
            string sLine = sLineRaw.Trim();
            if (sLine.Length == 0) continue;
            if (sLine.StartsWith(";") || sLine.StartsWith("#")) continue;
            if (sLine.StartsWith("[") && sLine.EndsWith("]")) continue;
            int iEq = sLine.IndexOf('=');
            if (iEq <= 0) continue;
            string sKey = sLine.Substring(0, iEq).Trim();
            string sVal = sLine.Substring(iEq + 1).Trim();
            d[sKey] = sVal;
        }
        return d;
    }

    private static bool getBool(Dictionary<string, string> d, string sKey)
    {
        string s;
        if (!d.TryGetValue(sKey, out s)) return false;
        s = (s ?? "").Trim();
        return s == "1" || s.Equals("true", StringComparison.OrdinalIgnoreCase) ||
            s.Equals("yes", StringComparison.OrdinalIgnoreCase);
    }

    private static string getOrEmpty(Dictionary<string, string> d, string sKey)
    {
        string s;
        if (!d.TryGetValue(sKey, out s)) return "";
        return s ?? "";
    }
}

// ===========================================================================
// GUI parameter dialog. Modal WinForms dialog that gathers source files,
// output directory, and option checkboxes (Force, View output,
// Log session, Use configuration). A separate Help button shows a
// MessageBox with an option to open the README in the browser; F1
// triggers the same handler. A Default settings button resets the
// fields to factory defaults and deletes the saved configuration.
//
// Returns true if OK was clicked, false if Cancel was clicked or the
// user closed the dialog. The ref parameters carry initial values in
// and the user's choices out.
// ===========================================================================
public static class guiDialog
{
    public static bool show(ref string sSource, ref string sOutputDir,
        ref bool bForce, ref bool bView, ref bool bLog, ref bool bUseCfg)
    {
        var frm = new Form();
        frm.Text = program.sProgramName;
        frm.FormBorderStyle = FormBorderStyle.FixedDialog;
        frm.StartPosition = FormStartPosition.CenterScreen;
        frm.MaximizeBox = false;
        frm.MinimizeBox = false;
        frm.ShowInTaskbar = false;

        // Layout constants. Declared before any use of iLayoutFormWidth
        // (ClientSize) so the literal 560 doesn't appear inline.
        const int iLayoutLeft = 12;
        const int iLayoutRight = 12;
        const int iLayoutTop = 12;
        const int iLayoutGap = 7;
        const int iLayoutRowGap = 11;
        const int iLayoutLabelWidth = 110;
        const int iLayoutButtonWidth = 130;
        const int iLayoutButtonHeight = 26;
        const int iLayoutTextHeight = 23;
        const int iLayoutFormWidth = 560;

        frm.ClientSize = new System.Drawing.Size(iLayoutFormWidth, 200);
        frm.Font = System.Drawing.SystemFonts.MessageBoxFont;

        // Route F1 to the Help button action.
        frm.KeyPreview = true;
        frm.KeyDown += (s, e) => {
            if (e.KeyCode == Keys.F1) {
                e.Handled = true;
                e.SuppressKeyPress = true;
                showHelpMessage();
            }
        };

        int iFormW = frm.ClientSize.Width;
        int iTextX = iLayoutLeft + iLayoutLabelWidth + iLayoutGap;
        int iTextW = iFormW - iTextX - iLayoutGap - iLayoutButtonWidth - iLayoutRight;
        int iBtnX = iFormW - iLayoutRight - iLayoutButtonWidth;

        // --- Row 1: Source files ---
        int y = iLayoutTop;
        var lblSource = new Label();
        lblSource.Text = "&Source files:";
        lblSource.AutoSize = false;
        lblSource.Location = new System.Drawing.Point(iLayoutLeft, y + 3);
        lblSource.Size = new System.Drawing.Size(iLayoutLabelWidth, iLayoutTextHeight);
        lblSource.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
        frm.Controls.Add(lblSource);

        var txtSource = new TextBox();
        txtSource.Text = string.IsNullOrWhiteSpace(sSource)
            ? defaultSourceForGui()
            : sSource;
        txtSource.Location = new System.Drawing.Point(iTextX, y);
        txtSource.Size = new System.Drawing.Size(iTextW, iLayoutTextHeight);
        txtSource.TabIndex = 0;
        frm.Controls.Add(txtSource);

        var btnBrowseSource = new Button();
        btnBrowseSource.Text = "&Browse source...";
        btnBrowseSource.Location = new System.Drawing.Point(iBtnX, y - 1);
        btnBrowseSource.Size = new System.Drawing.Size(iLayoutButtonWidth, iLayoutButtonHeight);
        btnBrowseSource.TabIndex = 1;
        btnBrowseSource.UseVisualStyleBackColor = true;
        frm.Controls.Add(btnBrowseSource);

        // --- Row 2: Output directory ---
        y += iLayoutTextHeight + iLayoutRowGap;
        var lblOut = new Label();
        lblOut.Text = "&Output directory:";
        lblOut.AutoSize = false;
        lblOut.Location = new System.Drawing.Point(iLayoutLeft, y + 3);
        lblOut.Size = new System.Drawing.Size(iLayoutLabelWidth, iLayoutTextHeight);
        lblOut.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
        frm.Controls.Add(lblOut);

        var txtOut = new TextBox();
        txtOut.Text = sOutputDir ?? "";
        txtOut.Location = new System.Drawing.Point(iTextX, y);
        txtOut.Size = new System.Drawing.Size(iTextW, iLayoutTextHeight);
        txtOut.TabIndex = 2;
        frm.Controls.Add(txtOut);

        var btnChooseOut = new Button();
        btnChooseOut.Text = "&Choose output...";
        btnChooseOut.Location = new System.Drawing.Point(iBtnX, y - 1);
        btnChooseOut.Size = new System.Drawing.Size(iLayoutButtonWidth, iLayoutButtonHeight);
        btnChooseOut.TabIndex = 3;
        btnChooseOut.UseVisualStyleBackColor = true;
        frm.Controls.Add(btnChooseOut);

        // Wire up Browse source: opens an OpenFileDialog. The user
        // can pick a single file; multi-select would be possible but
        // copying the multi-file result back into a textbox quoted
        // properly is more complex than necessary, and the common
        // case is wildcards which the user types directly.
        btnBrowseSource.Click += (s, e) => {
            using (var dialog = new OpenFileDialog()) {
                dialog.Title = "Choose a file to check";
                dialog.Filter =
                    "Supported files (*.docx;*.xlsx;*.pptx;*.md)|*.docx;*.xlsx;*.pptx;*.md|" +
                    "Word documents (*.docx)|*.docx|" +
                    "Excel workbooks (*.xlsx)|*.xlsx|" +
                    "PowerPoint presentations (*.pptx)|*.pptx|" +
                    "Markdown files (*.md)|*.md|" +
                    "All files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.CheckFileExists = true;
                dialog.RestoreDirectory = true;
                try {
                    dialog.InitialDirectory = getInitialBrowseDir(txtSource.Text);
                } catch { }
                if (dialog.ShowDialog(frm) == DialogResult.OK) {
                    string sPicked = dialog.FileName;
                    if (sPicked.Contains(" "))
                        sPicked = "\"" + sPicked + "\"";
                    txtSource.Text = sPicked;
                    if (string.IsNullOrEmpty(txtOut.Text))
                        txtOut.Text = Path.GetDirectoryName(dialog.FileName);
                }
            }
        };

        // Wire up Choose output: FolderBrowserDialog.
        btnChooseOut.Click += (s, e) => {
            using (var dialog = new FolderBrowserDialog()) {
                dialog.Description = "Choose the output directory";
                dialog.ShowNewFolderButton = true;
                try {
                    dialog.SelectedPath = getInitialBrowseDir(txtOut.Text);
                } catch { }
                if (dialog.ShowDialog(frm) == DialogResult.OK)
                    txtOut.Text = dialog.SelectedPath;
            }
        };

        // --- Row 3: Force replacements + View output ---
        y += iLayoutTextHeight + iLayoutRowGap * 2;
        int iChkW = (iFormW - iLayoutLeft - iLayoutRight) / 2;
        var chkForce = new CheckBox();
        chkForce.Text = "&Force replacements";
        chkForce.Checked = bForce;
        chkForce.Location = new System.Drawing.Point(iLayoutLeft, y);
        chkForce.Size = new System.Drawing.Size(iChkW, iLayoutTextHeight);
        chkForce.TabIndex = 4;
        frm.Controls.Add(chkForce);

        var chkView = new CheckBox();
        chkView.Text = "&View output";
        chkView.Checked = bView;
        chkView.Location = new System.Drawing.Point(iLayoutLeft + iChkW, y);
        chkView.Size = new System.Drawing.Size(iChkW, iLayoutTextHeight);
        chkView.TabIndex = 5;
        frm.Controls.Add(chkView);

        // --- Row 4: Log session + Use configuration ---
        // Both are "meta" options that affect persistence/diagnostics
        // rather than the conversion itself, so they sit together
        // below the conversion-control checkboxes.
        y += iLayoutTextHeight + iLayoutRowGap;
        var chkLog = new CheckBox();
        chkLog.Text = "&Log session";
        chkLog.Checked = bLog;
        chkLog.Location = new System.Drawing.Point(iLayoutLeft, y);
        chkLog.Size = new System.Drawing.Size(iChkW, iLayoutTextHeight);
        chkLog.TabIndex = 6;
        frm.Controls.Add(chkLog);

        var chkUseCfg = new CheckBox();
        chkUseCfg.Text = "&Use configuration";
        chkUseCfg.Checked = bUseCfg;
        chkUseCfg.Location = new System.Drawing.Point(iLayoutLeft + iChkW, y);
        chkUseCfg.Size = new System.Drawing.Size(iChkW, iLayoutTextHeight);
        chkUseCfg.TabIndex = 7;
        frm.Controls.Add(chkUseCfg);

        // --- Bottom row: commit buttons. Help and Default settings
        // on the left (they don't commit/cancel the dialog), OK and
        // Cancel on the right. Matches Microsoft's UX guidance.
        y += iLayoutTextHeight + iLayoutRowGap * 2;
        var btnHelp = new Button();
        btnHelp.Text = "&Help";
        btnHelp.Location = new System.Drawing.Point(iLayoutLeft, y);
        btnHelp.Size = new System.Drawing.Size(iLayoutButtonWidth, iLayoutButtonHeight);
        btnHelp.TabIndex = 8;
        btnHelp.UseVisualStyleBackColor = true;
        btnHelp.Click += (s, e) => showHelpMessage();
        frm.Controls.Add(btnHelp);

        var btnDefaults = new Button();
        btnDefaults.Text = "&Default settings";
        btnDefaults.Location = new System.Drawing.Point(
            iLayoutLeft + iLayoutButtonWidth + iLayoutGap, y);
        btnDefaults.Size = new System.Drawing.Size(iLayoutButtonWidth, iLayoutButtonHeight);
        btnDefaults.TabIndex = 9;
        btnDefaults.UseVisualStyleBackColor = true;
        btnDefaults.Click += (s, e) => {
            string sDefault = defaultSourceForGui();
            txtSource.Text = sDefault;
            txtOut.Text = "";
            chkForce.Checked = false;
            chkView.Checked = false;
            chkLog.Checked = false;
            chkUseCfg.Checked = false;
            configManager.eraseAll();
        };
        frm.Controls.Add(btnDefaults);

        var btnOk = new Button();
        btnOk.Text = "OK";
        btnOk.DialogResult = DialogResult.OK;
        btnOk.Location = new System.Drawing.Point(
            iFormW - iLayoutRight - 2 * iLayoutButtonWidth - iLayoutGap, y);
        btnOk.Size = new System.Drawing.Size(iLayoutButtonWidth, iLayoutButtonHeight);
        btnOk.TabIndex = 10;
        btnOk.UseVisualStyleBackColor = true;
        // Validate output directory before allowing the dialog to close.
        // If the user has typed a non-existent directory, prompt to
        // create it (default Yes). On No (or creation failure), set
        // DialogResult = None so the dialog stays open and the user can
        // edit the field. WinForms invokes Click handlers BEFORE the
        // automatic close, so this hook runs first.
        btnOk.Click += (s, e) => {
            string sOutCandidate = (txtOut.Text ?? "").Trim();
            if (sOutCandidate.Length >= 2 && sOutCandidate[0] == '"' && sOutCandidate[sOutCandidate.Length - 1] == '"')
                sOutCandidate = sOutCandidate.Substring(1, sOutCandidate.Length - 2).Trim();
            if (string.IsNullOrEmpty(sOutCandidate)) return;
            try {
                if (Directory.Exists(sOutCandidate)) return;
            } catch { return; }
            DialogResult dr = MessageBox.Show(frm,
                "Create " + sOutCandidate + "?",
                program.sProgramName,
                MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1);
            if (dr != DialogResult.Yes) {
                frm.DialogResult = DialogResult.None;
                txtOut.Focus();
                return;
            }
            try {
                Directory.CreateDirectory(sOutCandidate);
            } catch (Exception ex) {
                MessageBox.Show(frm,
                    "Could not create directory:\r\n" + sOutCandidate + "\r\n\r\n" + ex.Message,
                    program.sProgramName,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                frm.DialogResult = DialogResult.None;
                txtOut.Focus();
            }
        };
        frm.Controls.Add(btnOk);

        var btnCancel = new Button();
        btnCancel.Text = "Cancel";
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Location = new System.Drawing.Point(iBtnX, y);
        btnCancel.Size = new System.Drawing.Size(iLayoutButtonWidth, iLayoutButtonHeight);
        btnCancel.TabIndex = 11;
        btnCancel.UseVisualStyleBackColor = true;
        frm.Controls.Add(btnCancel);

        // Wire Enter = OK and Esc = Cancel.
        frm.AcceptButton = btnOk;
        frm.CancelButton = btnCancel;

        // Resize the form to wrap the bottom row tightly.
        frm.ClientSize = new System.Drawing.Size(iFormW,
            y + iLayoutButtonHeight + iLayoutTop);

        // Show modally.
        var dialogResult = frm.ShowDialog();
        if (dialogResult != DialogResult.OK) return false;

        sSource = txtSource.Text.Trim();
        sOutputDir = txtOut.Text.Trim();
        bForce = chkForce.Checked;
        bView = chkView.Checked;
        bLog = chkLog.Checked;
        bUseCfg = chkUseCfg.Checked;
        return true;
    }

    private static string defaultSourceForGui()
    {
        try {
            return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        } catch {
            return "";
        }
    }

    /// <summary>
    /// Choose the initial directory for a file or folder picker.
    /// Returns the directory derived from the user's current text-field
    /// content (in-session choice or value loaded from saved config) when
    /// it points to an existing path; falls back to the user's Documents
    /// folder otherwise. This follows Microsoft's guidance that pickers
    /// start where the user has shown intent, with Documents as a
    /// sensible no-context default.
    /// </summary>
    /// <param name="sFieldText">The current text of the source-input or
    /// output-directory field. May be empty, may contain wildcards or
    /// quoted paths, may contain space-separated tokens (only the first
    /// is inspected).</param>
    /// <returns>An existing directory path. Always non-empty.</returns>
    private static string getInitialBrowseDir(string sFieldText)
    {
        string sCandidate = "";
        string sFirstToken = "";
        string sParent = "";

        sFieldText = (sFieldText ?? "").Trim();
        if (string.IsNullOrEmpty(sFieldText)) return defaultSourceForGui();
        // Inspect first space-separated token only.
        int iSpace = sFieldText.IndexOf(' ');
        sFirstToken = (iSpace < 0) ? sFieldText : sFieldText.Substring(0, iSpace);
        sFirstToken = sFirstToken.Trim('"');
        if (string.IsNullOrEmpty(sFirstToken)) return defaultSourceForGui();
        sCandidate = sFirstToken;
        try {
            // Wildcards: strip the basename and inspect the parent directory.
            if (sCandidate.IndexOfAny(new[] { '*', '?' }) >= 0) {
                sCandidate = Path.GetDirectoryName(sCandidate) ?? "";
            }
            if (!string.IsNullOrEmpty(sCandidate) && Directory.Exists(sCandidate))
                return Path.GetFullPath(sCandidate);
            if (!string.IsNullOrEmpty(sCandidate) && File.Exists(sCandidate))
                return Path.GetDirectoryName(Path.GetFullPath(sCandidate));
            sParent = string.IsNullOrEmpty(sCandidate) ? "" : (Path.GetDirectoryName(sCandidate) ?? "");
            if (!string.IsNullOrEmpty(sParent) && Directory.Exists(sParent))
                return Path.GetFullPath(sParent);
        } catch { }
        return defaultSourceForGui();
    }

    private static void showHelpMessage()
    {
        string sMsg =
            "extCheck checks Word, Excel, PowerPoint, and Markdown files for " +
            "accessibility problems and writes a CSV report for each file.\r\n\r\n" +
            "Source files: a single file path, a wildcard pattern, or several " +
            "of either separated by spaces. Wrap paths containing spaces in " +
            "double quotes.\r\n\r\n" +
            "Output directory: where each <basename>.csv is written. Blank " +
            "means the current working directory.\r\n\r\n" +
            "Options:\r\n" +
            "  Force replacements - overwrite an existing CSV instead of " +
            "skipping the input\r\n" +
            "  View output - open the output directory in Explorer when done\r\n" +
            "  Log session - write extCheck.log (replacing any prior log) " +
            "to the output directory, or to the current directory if no " +
            "output directory is set\r\n" +
            "  Use configuration - remember these settings for next time, in " +
            "%LOCALAPPDATA%\\extCheck\\extCheck.ini\r\n\r\n" +
            "Press Cancel to exit without checking.\r\n\r\n" +
            "Open the full README in your browser?";
        var dialogResult = MessageBox.Show(sMsg,
            "extCheck — Help",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Information,
            MessageBoxDefaultButton.Button2);
        if (dialogResult == DialogResult.Yes) launchReadMe();
    }

    private static void launchReadMe()
    {
        string sExeDir = Path.GetDirectoryName(
            Assembly.GetExecutingAssembly().Location);
        string sHtm = Path.Combine(sExeDir, "ReadMe.htm");
        string sMd = Path.Combine(sExeDir, "ReadMe.md");
        string sTarget = File.Exists(sHtm)
            ? sHtm
            : (File.Exists(sMd) ? sMd : null);
        if (sTarget == null) {
            MessageBox.Show(
                "Documentation (ReadMe.htm or ReadMe.md) was not found in " +
                "the extCheck install folder:\r\n\r\n" + sExeDir + "\r\n\r\n" +
                "If extCheck was installed via the installer, reinstall it. " +
                "If you deployed extCheck.exe manually, place ReadMe.htm " +
                "(or ReadMe.md) in the same folder.",
                "extCheck — Documentation not found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }
        try {
            var processStartInfo = new ProcessStartInfo(sTarget) { UseShellExecute = true };
            Process.Start(processStartInfo);
        } catch (Exception ex) {
            MessageBox.Show(
                "Could not open the documentation:\r\n\r\n" + ex.Message,
                "extCheck — Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }
}

// ===========================================================================
// guiProgress: a small modeless status form shown during the file-checking
// loop in GUI mode. Mirrors 2htm's guiProgress so users get visible
// progress feedback during long runs (e.g. multi-megabyte .docx files
// that take Word several seconds each to open and analyze).
//
// The displayed count reflects files ALREADY COMPLETED, not the file
// being started. When starting file 1 of 5, we show
// "report.docx -- 0 of 5, 0%". The percent and count advance only
// after the file finishes. This avoids the confusion of seeing "100%"
// while the (only) file is still being processed.
// ===========================================================================
public static class guiProgress
{
    private static Form frm;
    private static Label lblStatus;

    public static void open(int iTotal)
    {
        frm = new Form();
        frm.Text = "extCheck — Checking";
        frm.FormBorderStyle = FormBorderStyle.FixedDialog;
        frm.StartPosition = FormStartPosition.CenterScreen;
        frm.MaximizeBox = false;
        frm.MinimizeBox = false;
        frm.ControlBox = false;
        frm.ShowInTaskbar = true;
        frm.ClientSize = new System.Drawing.Size(480, 92);
        frm.Font = System.Drawing.SystemFonts.MessageBoxFont;

        var lblIntro = new Label();
        lblIntro.Text = "Checking files. Please wait...";
        lblIntro.AutoSize = false;
        lblIntro.Location = new System.Drawing.Point(14, 14);
        lblIntro.Size = new System.Drawing.Size(452, 22);
        lblIntro.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
        frm.Controls.Add(lblIntro);

        lblStatus = new Label();
        lblStatus.Text = "Starting...";
        lblStatus.AutoSize = false;
        lblStatus.Location = new System.Drawing.Point(14, 42);
        lblStatus.Size = new System.Drawing.Size(452, 22);
        lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
        lblStatus.AccessibleName = "Check status";
        lblStatus.AccessibleRole = AccessibleRole.StatusBar;
        frm.Controls.Add(lblStatus);

        frm.Show();
        Application.DoEvents();
    }

    public static void update(string sBase, int iIndex, int iTotal)
    {
        if (frm == null || lblStatus == null) return;
        // iIndex is 1-based; we show "iIndex - 1" as the completed count
        // so the percentage and count reflect work DONE, not work in
        // progress.
        int iCompleted = iIndex - 1;
        int iPercent = iTotal > 0 ? (iCompleted * 100 / iTotal) : 0;
        lblStatus.Text = sBase + " \u2014 " + iCompleted + " of " + iTotal +
            ", " + iPercent + "%";
        Application.DoEvents();
    }

    public static void close()
    {
        if (frm == null) return;
        try { frm.Close(); frm.Dispose(); } catch { }
        frm = null;
        lblStatus = null;
    }
}

// ===========================================================================
// Main program. Parses arguments, optionally shows a GUI dialog, dispatches
// to the format modules, and writes per-file CSV reports.
// ===========================================================================
static class program {

    public const string sProgramName = "extCheck";
    public const string sProgramVersion = "2.0";
    public const string sConfigDirName = "extCheck";
    public const string sConfigFileName = "extCheck.ini";
    public const string sLogFileName = "extCheck.log";
    public static readonly string[] aSupportedExtensions = { ".docx", ".xlsx", ".pptx", ".md" };

    // Trim and shorten a message for inline display next to a
    // basename. COM exceptions can produce multi-paragraph text;
    // we want a single short line. Returns "" if the input is
    // null or empty.
    public static string firstLine(string s) {
        const int iMaxLen = 120;
        if (string.IsNullOrEmpty(s)) return "";
        int i = s.IndexOfAny(new[] { '\r', '\n' });
        if (i >= 0) s = s.Substring(0, i);
        s = s.Trim();
        if (s.Length > iMaxLen) s = s.Substring(0, iMaxLen - 3) + "...";
        return s;
    }

    // A single failure record: basename and a short reason. Used
    // in the structured results summary to render
    // "basename: reason" on one line. The full exception (if any)
    // is in the log when -l is on.
    public class failure {
        public string sBase;
        public string sReason;
        public failure(string sB, string sR) { sBase = sB; sReason = sR; }
    }

    // Globals set by command-line switches and/or the GUI dialog.
    public static bool bGuiMode = false;
    public static bool bShowRules = false;
    public static bool bForce = false;
    public static bool bLog = false;
    public static bool bUseConfig = false;
    public static bool bViewOutput = false;
    public static bool bHideConsoleMode = false;

    // Parallel "was this set on the command line" flags. When true,
    // the saved config file must NOT overwrite the command-line-supplied
    // value during the pre-GUI load.
    public static bool bForceFromCli = false;
    public static bool bLogFromCli = false;
    public static bool bViewOutputFromCli = false;
    public static bool bOutputDirFromCli = false;
    public static bool bSourceFromCli = false;

    // Output directory. Empty means "current working directory".
    public static string sOutputDir = "";

    // ---- GUI / console detection (Win32 P/Invoke) ----
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint GetConsoleProcessList(
        [Out] uint[] aiProcessIds, uint iCount);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private const int iSwHide = 0;

    private static bool isLaunchedFromGui()
    {
        try {
            var aiList = new uint[16];
            uint iCount = GetConsoleProcessList(aiList, (uint)aiList.Length);
            return iCount == 1;
        } catch {
            return false;
        }
    }

    private static void hideOwnConsoleWindow()
    {
        try {
            IntPtr hwnd = GetConsoleWindow();
            if (hwnd != IntPtr.Zero) ShowWindow(hwnd, iSwHide);
        } catch { }
    }

    // ---- Argument splitter for source field ----
    // Re-parses a space-separated string of file paths, honoring
    // double-quoted segments. Used by configManager.loadInto and by
    // the dialog's OK-path.
    public static List<string> splitSourceField(string sField)
    {
        // Friendlier parsing rules:
        //   1. Trim the input.
        //   2. Strip a single layer of surrounding double quotes.
        //   3. Test the entire (unquoted) trimmed field as a single spec
        //      (existing file, existing directory, or wildcard pattern
        //      that matches at least one file). If usable, return it
        //      as one token.
        //   4. Otherwise fall back to space-tokenization, honoring
        //      "..." segments so the user can mix paths-with-spaces
        //      and ones without.
        // The user only needs to use quotes when supplying multiple
        // specs and at least one contains a space. For a single path,
        // no quotes are needed -- the entire trimmed input is tested
        // as one spec.
        var lsResult = new List<string>();
        if (string.IsNullOrWhiteSpace(sField)) return lsResult;
        string sTrimmed = sField.Trim();
        string sUnquoted = sTrimmed;
        if (sUnquoted.Length >= 2 && sUnquoted[0] == '"' && sUnquoted[sUnquoted.Length - 1] == '"')
            sUnquoted = sUnquoted.Substring(1, sUnquoted.Length - 2).Trim();

        // Test the full unquoted field as a single spec first.
        if (isUsableSingleSpec(sUnquoted)) {
            lsResult.Add(sUnquoted);
            return lsResult;
        }

        // Fall back: tokenize on whitespace, honoring quoted segments.
        var sb = new StringBuilder();
        bool bInQuotes = false;
        foreach (char c in sTrimmed) {
            if (c == '"') {
                bInQuotes = !bInQuotes;
                continue;
            }
            if (c == ' ' && !bInQuotes) {
                if (sb.Length > 0) {
                    lsResult.Add(sb.ToString());
                    sb.Clear();
                }
                continue;
            }
            sb.Append(c);
        }
        if (sb.Length > 0) lsResult.Add(sb.ToString());
        return lsResult;
    }

    /// <summary>
    /// Return true if sSpec, taken whole, is a usable file specification
    /// (existing file, existing directory, or wildcard pattern that
    /// matches at least one file). Used by splitSourceField to decide
    /// whether the entire trimmed field is one path that happens to
    /// contain spaces, vs multiple space-separated paths.
    /// </summary>
    private static bool isUsableSingleSpec(string sSpec)
    {
        if (string.IsNullOrEmpty(sSpec)) return false;
        try {
            if (File.Exists(sSpec)) return true;
            if (Directory.Exists(sSpec)) return true;
            // Wildcard? Try to enumerate.
            if (sSpec.IndexOfAny(new[] { '*', '?' }) >= 0) {
                string sDir = Path.GetDirectoryName(sSpec);
                if (string.IsNullOrEmpty(sDir)) sDir = Directory.GetCurrentDirectory();
                string sPattern = Path.GetFileName(sSpec);
                if (Directory.Exists(sDir) && !string.IsNullOrEmpty(sPattern)) {
                    try {
                        string[] aMatched = Directory.GetFiles(sDir, sPattern);
                        if (aMatched != null && aMatched.Length > 0) return true;
                    } catch { }
                }
            }
        } catch { }
        return false;
    }

    // ---- Resolve filespecs (with wildcards) to a flat file list. ----
    private static List<string> resolveFiles(List<string> lsSpecs)
    {
        var lsFiles = new List<string>();
        foreach (string sSpec in lsSpecs) {
            string sRawDir = Path.GetDirectoryName(sSpec);
            string sDir = (sRawDir == null || sRawDir == "")
                ? Directory.GetCurrentDirectory()
                : Path.IsPathRooted(sRawDir) ? sRawDir : Path.Combine(Directory.GetCurrentDirectory(), sRawDir);
            string sPattern = Path.GetFileName(sSpec);

            bool bStarExt = sPattern.EndsWith(".*");
            if (bStarExt) {
                string sBase = sPattern.Substring(0, sPattern.Length - 2);
                bool bAnyFound = false;
                foreach (string sExt in aSupportedExtensions) {
                    string sPat = (sBase == "") ? "*" + sExt : sBase + sExt;
                    try {
                        string[] aExt = Directory.GetFiles(sDir, sPat);
                        lsFiles.AddRange(aExt);
                        if (aExt.Length > 0) bAnyFound = true;
                    } catch { }
                }
                if (!bAnyFound) Console.WriteLine("No supported files matched: " + sSpec);
                continue;
            }

            string[] aFound;
            try {
                aFound = Directory.GetFiles(sDir, sPattern);
            } catch (Exception ex) {
                Console.WriteLine("ERROR resolving '" + sSpec + "': " + ex.Message);
                continue;
            }
            if (aFound.Length == 0) Console.WriteLine("No files matched: " + sSpec);
            lsFiles.AddRange(aFound);
        }
        return lsFiles;
    }

    // ---- Open the given directory in File Explorer. Reuses an
    // already-open Explorer window on the same folder if possible. ----
    private static void openOutputInExplorer(string sDir)
    {
        try {
            var processStartInfo = new ProcessStartInfo("explorer.exe", "\"" + sDir + "\"");
            processStartInfo.UseShellExecute = true;
            Process.Start(processStartInfo);
        } catch (Exception ex) {
            Console.Error.WriteLine("[WARN] Could not open output directory: " + ex.Message);
        }
    }

    // ---- Show a final MessageBox in GUI mode summarizing the run. ----
    private static void showFinalMessage(string sBody, string sTitle = "extCheck — Results")
    {
        try {
            MessageBox.Show(string.IsNullOrEmpty(sBody) ? "Done." : sBody,
                sTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
        } catch { }
    }

    // ---- Print CLI help. ----
    private const string sUsage = @"
Usage:
  extCheck.exe [-h] [-g] [-rules] [-o <dir>] [--view-output]
               [-l] [-u] [-f] <filespec> [<filespec> ...]

Arguments:
  <filespec>   Path to one or more files to evaluate.
               Wildcards are supported: *.docx  docs\*.md  C:\work\*.pptx

Supported file formats:
  .docx   Microsoft Word document
  .xlsx   Microsoft Excel workbook
  .pptx   Microsoft PowerPoint presentation
  .md     Pandoc Markdown file

Options:
  -h, --help            Show this help screen and exit.
  -g, --gui-mode        Show the parameter dialog. GUI mode is also entered
                        automatically when extCheck is launched without
                        arguments from a GUI shell (File Explorer, Start
                        menu, desktop hotkey).
  -rules                Write extCheck.csv (the complete rule registry) to
                        the output directory or current directory and exit.
  -o, --output-dir <d>  Directory in which the per-file CSV reports are
                        written. Defaults to the current working directory.
                        Created if it does not exist.
  --view-output         After all files are checked, open the output
                        directory in File Explorer.
  -f, --force           Overwrite an existing <basename>.csv. Without this
                        flag, an input file is skipped if its CSV already
                        exists in the output directory.
  -l, --log             Write detailed diagnostics to extCheck.log (UTF-8
                        with BOM) in the output directory if one is set,
                        else the current working directory. Any prior
                        extCheck.log is overwritten so the file always
                        reflects only the current session.
  -u, --use-configuration
                        Load saved settings from
                        %LOCALAPPDATA%\extCheck\extCheck.ini at startup,
                        and (in GUI mode) write them back on OK. Without
                        this flag extCheck leaves no filesystem footprint
                        of its own.

Output:
  For each file evaluated, a CSV named <basename>.csv is written to the
  output directory. The CSV columns are:
    RuleID, Source, Category, Location, Context, Message, Remediation

  Source values:
    MSAC  Rule derived from MS Office Accessibility Checker categories
    AXE   Rule derived from axe-core WCAG 2.1 equivalents

  A brief progress line is printed for each file (basename, plus
  ""skipped"" or ""ERROR"" if applicable). The detailed issue list is in
  the CSV; full paths and stack traces (when relevant) are in the
  log when -l is given.

Notes:
  Word, Excel, and PowerPoint must be installed to check .docx, .xlsx,
  .pptx files. PowerPoint requires a visible application window; it is
  minimized automatically. No Office installation is needed for .md files.
  extCheck is a 64-bit program; the installed Office must also be 64-bit.

Examples:
  extCheck.exe report.docx
  extCheck.exe *.md -o reports
  extCheck.exe docs\*.docx data\*.xlsx slides\*.pptx --view-output
  extCheck.exe -rules -o C:\references
  extCheck.exe -g
";

    // ---- Main run() ----
    public static int run(string[] aArgs)
    {
        // Phase 1: GUI auto-detection. We need this before parsing
        // args because a no-argument GUI launch should never print
        // usage; instead it should fall through to the dialog.
        bool bOwnConsole = isLaunchedFromGui();
        var lsFileArgs = new List<string>();

        // Phase 2: parse CLI args. Recognized flags are consumed;
        // unrecognized tokens are file specs.
        for (int i = 0; i < aArgs.Length; i++) {
            string sArg = aArgs[i];
            if (sArg == "-h" || sArg == "/h" || sArg == "--help") {
                Console.WriteLine(sProgramName + " " + sProgramVersion +
                    " — Accessibility Checker for .docx .xlsx .pptx .md");
                Console.WriteLine();
                Console.Write(sUsage);
                return 0;
            }
            if (sArg == "-v" || sArg == "--version") {
                Console.WriteLine(sProgramName + " " + sProgramVersion);
                return 0;
            }
            if (sArg == "-g" || sArg == "--gui-mode") { bGuiMode = true; continue; }
            if (sArg == "-rules" || sArg == "/rules") { bShowRules = true; continue; }
            if (sArg == "-f" || sArg == "--force") {
                bForce = true; bForceFromCli = true; continue;
            }
            if (sArg == "-l" || sArg == "--log") {
                bLog = true; bLogFromCli = true; continue;
            }
            if (sArg == "-u" || sArg == "--use-configuration") {
                bUseConfig = true; continue;
            }
            if (sArg == "--view-output") {
                bViewOutput = true; bViewOutputFromCli = true; continue;
            }
            if (sArg == "-o" || sArg == "--output-dir") {
                if (i + 1 >= aArgs.Length) {
                    Console.Error.WriteLine("[ERROR] " + sArg +
                        " requires a directory argument.");
                    return 1;
                }
                sOutputDir = aArgs[++i];
                bOutputDirFromCli = true;
                continue;
            }
            // Unrecognized -> file spec.
            lsFileArgs.Add(sArg);
        }
        if (lsFileArgs.Count > 0) bSourceFromCli = true;

        // Phase 3: GUI auto-launch when invoked from a GUI shell with
        // no arguments. The console window (if any) belongs to us
        // alone, so we hide it.
        if (aArgs.Length == 0 && bOwnConsole) {
            bGuiMode = true;
            hideOwnConsoleWindow();
        } else if (bOwnConsole && lsFileArgs.Count > 0) {
            // CLI-style invocation from Explorer (e.g., right-click
            // verb) -- console flash would be confusing. Hide.
            bHideConsoleMode = true;
            hideOwnConsoleWindow();
        }

        // Phase 4: load config if requested or if GUI mode and a
        // config exists (we treat the existence of a config as
        // implicit opt-in for GUI mode, matching 2htm).
        if (bGuiMode && !bUseConfig && configManager.configExists()) {
            bUseConfig = true;
        }
        if (bUseConfig) configManager.loadInto(lsFileArgs);

        // Phase 5: in GUI mode, show the parameter dialog. Cancel = exit.
        if (bGuiMode) {
            string sSource = string.Join(" ",
                lsFileArgs.ConvertAll(s => s.Contains(" ") ? "\"" + s + "\"" : s));
            string sOut = sOutputDir;
            bool bForceLocal = bForce;
            bool bView = bViewOutput;
            bool bLogLocal = bLog;
            bool bUseCfgLocal = bUseConfig;
            if (!guiDialog.show(ref sSource, ref sOut, ref bForceLocal,
                ref bView, ref bLogLocal, ref bUseCfgLocal)) {
                return 0;
            }
            bForce = bForceLocal;
            bViewOutput = bView;
            bLog = bLogLocal;
            bUseConfig = bUseCfgLocal;
            sOutputDir = sOut;
            lsFileArgs = splitSourceField(sSource);
            if (bUseConfig)
                configManager.save(sSource, sOutputDir, bForce, bViewOutput, bLog);
        }

        // Phase 6: resolve output directory and open the log if
        // requested. Output dir defaults to CWD when empty.
        string sResolvedOutDir;
        if (string.IsNullOrWhiteSpace(sOutputDir)) {
            sResolvedOutDir = Directory.GetCurrentDirectory();
        } else {
            try {
                sResolvedOutDir = Path.GetFullPath(sOutputDir);
                if (!Directory.Exists(sResolvedOutDir))
                    Directory.CreateDirectory(sResolvedOutDir);
            } catch (Exception ex) {
                string sErr = "Output directory '" + sOutputDir +
                    "' could not be created: " + ex.Message;
                Console.Error.WriteLine("[ERROR] " + sErr);
                if (bGuiMode) showFinalMessage(sErr, "extCheck — Error");
                return 1;
            }
        }
        if (bLog) logger.open(sResolvedOutDir);

        // Write the run header to the log: program version, friendly
        // start time, and the resolved parameter list (showing both
        // explicit and defaulted values). Mirrors the GUI dialog
        // controls so the user can map a logged run to the dialog.
        var lsParams = new List<KeyValuePair<string, string>>();
        lsParams.Add(new KeyValuePair<string, string>("Source",
            lsFileArgs.Count == 0
                ? "(none)"
                : string.Join(" ", lsFileArgs.ConvertAll(s => s.Contains(" ") ? "\"" + s + "\"" : s))));
        lsParams.Add(new KeyValuePair<string, string>("Output directory", sResolvedOutDir));
        lsParams.Add(new KeyValuePair<string, string>("Force replacements", bForce.ToString().ToLowerInvariant()));
        lsParams.Add(new KeyValuePair<string, string>("View output",        bViewOutput.ToString().ToLowerInvariant()));
        lsParams.Add(new KeyValuePair<string, string>("Use configuration",  bUseConfig.ToString().ToLowerInvariant()));
        lsParams.Add(new KeyValuePair<string, string>("Log session",        bLog.ToString().ToLowerInvariant()));
        lsParams.Add(new KeyValuePair<string, string>("Show rules",         bShowRules.ToString().ToLowerInvariant()));
        lsParams.Add(new KeyValuePair<string, string>("GUI mode",           bGuiMode.ToString().ToLowerInvariant()));
        lsParams.Add(new KeyValuePair<string, string>("Working directory",  Directory.GetCurrentDirectory()));
        lsParams.Add(new KeyValuePair<string, string>("Command line",       string.Join(" ", aArgs)));
        logger.header(sProgramName, sProgramVersion, lsParams);

        // Phase 7: capture stdout/stderr in GUI mode for the final
        // dialog. In hide-console mode, capture stderr only.
        TextWriter writerOriginalOut = Console.Out;
        TextWriter writerOriginalErr = Console.Error;
        StringWriter stringWriterOut = null;
        StringWriter stringWriterErr = null;
        if (bGuiMode) {
            stringWriterOut = new StringWriter();
            Console.SetOut(stringWriterOut);
            Console.SetError(stringWriterOut);
        } else if (bHideConsoleMode) {
            stringWriterErr = new StringWriter();
            Console.SetError(stringWriterErr);
        }

        try {
            Console.WriteLine(sProgramName + " " + sProgramVersion +
                " — Accessibility Checker for .docx .xlsx .pptx .md");
            Console.WriteLine();

            // Phase 8: handle -rules (immediate write of rule registry).
            if (bShowRules) {
                string sRulesPath = Path.Combine(sResolvedOutDir, "extCheck.csv");
                shared.writeRulesCsv(sRulesPath);
                Console.WriteLine("Wrote rule registry: " + sRulesPath);
                logger.info("Wrote rule registry: " + sRulesPath);
                if (lsFileArgs.Count == 0) {
                    if (bViewOutput) openOutputInExplorer(sResolvedOutDir);
                    return 0;
                }
                // -rules combined with file specs: continue and
                // process the files too.
            }

            // Phase 9: file processing.
            if (lsFileArgs.Count == 0) {
                if (!bGuiMode && !bShowRules) {
                    Console.Write(sUsage);
                    return 0;
                }
                Console.WriteLine("No files to process.");
                return 1;
            }

            var lsFiles = resolveFiles(lsFileArgs);
            if (lsFiles.Count == 0) {
                Console.WriteLine("No files to process.");
                return 1;
            }

            // ---- Pre-pruning ----
            //
            // Two pruning passes happen BEFORE the conversion/check
            // loop and BEFORE the progress UI opens, so the progress
            // counter's denominator reflects only files that will
            // actually be processed:
            //
            //   1. Drop files whose extension extCheck cannot check.
            //      These are silently dropped; they do not appear in
            //      any of the three result sections (Checked / Failed /
            //      Skipped) since the user cannot do anything about
            //      "Force replacements" to make them processable. The
            //      log records the unsupported skip when -l is given.
            //
            //   2. If --force is NOT set, drop files whose target CSV
            //      already exists. These ARE counted as "skipped" and
            //      surface in the final results MessageBox with a hint
            //      about Force replacements.
            //
            // The lists below feed three sections in the final
            // MessageBox: lsToCheck (the survivors of pruning) feed
            // the loop and end up split into lsChecked + lsFailed; a
            // separate lsSkippedExisting tracks files dropped in
            // pass (2).
            var lsToCheck = new List<string>();
            var lsSkippedExisting = new List<string>();
            foreach (string sFilePath in lsFiles) {
                string sExt = Path.GetExtension(sFilePath).ToLower();
                bool bSupported = false;
                foreach (string sE in aSupportedExtensions) {
                    if (sExt == sE) { bSupported = true; break; }
                }
                if (!bSupported) {
                    logger.info("Skipped (unsupported format " + sExt + "): " + sFilePath);
                    continue;
                }
                string sCsvPath = Path.Combine(sResolvedOutDir,
                    Path.GetFileNameWithoutExtension(sFilePath) + ".csv");
                if (File.Exists(sCsvPath) && !bForce) {
                    lsSkippedExisting.Add(sFilePath);
                    logger.info("Skipped (CSV exists; use --force to overwrite): " + sFilePath);
                    continue;
                }
                lsToCheck.Add(sFilePath);
            }

            int iTotalIssues = 0;
            int iFileIndex = 0;
            var lsChecked = new List<string>();
            var lsFailed = new List<failure>();
            bool bGuiOrHidden = bGuiMode || bHideConsoleMode;

            // In GUI/hidden modes, open a small modeless status form
            // so the user gets visible progress feedback during long
            // runs. In pure CLI mode (real console attached), no
            // form -- the inline basenames printed below ARE the
            // progress indicator. The counter denominator is
            // lsToCheck.Count (post-pruning) so percentages reflect
            // actual work.
            if (bGuiOrHidden) guiProgress.open(lsToCheck.Count);

            try {
            foreach (string sFilePath in lsToCheck) {
                iFileIndex++;
                if (bGuiOrHidden)
                    guiProgress.update(Path.GetFileName(sFilePath), iFileIndex, lsToCheck.Count);

                string sExt = Path.GetExtension(sFilePath).ToLower();
                string sCsvPath = Path.Combine(sResolvedOutDir,
                    Path.GetFileNameWithoutExtension(sFilePath) + ".csv");
                string sBase = Path.GetFileName(sFilePath);

                logger.info("Checking " + sFilePath);

                // CLI mode: print the basename immediately; no
                // newline yet. On success, terminate with "\n"; on
                // failure, append ": <reason>\n". GUI/hidden mode:
                // no inline write -- the captured stdout becomes
                // the final MessageBox, and we want only the
                // structured summary there.
                if (!bGuiOrHidden) Console.Write(sBase);

                results.clear();
                bool bOpened = false;
                string sCheckError = null;

                if (sExt == ".docx") {
                    bOpened = docxModule.open(sFilePath);
                    if (bOpened) {
                        try { docxModule.checkAll(sFilePath); }
                        catch (Exception ex) { sCheckError = ex.Message; }
                        docxModule.quit();
                    }
                } else if (sExt == ".xlsx") {
                    bOpened = xlsxModule.open(sFilePath);
                    if (bOpened) {
                        try { xlsxModule.checkAll(sFilePath); }
                        catch (Exception ex) { sCheckError = ex.Message; }
                        xlsxModule.quit();
                    }
                } else if (sExt == ".pptx") {
                    bOpened = pptxModule.open(sFilePath);
                    if (bOpened) {
                        try { pptxModule.checkAll(sFilePath); }
                        catch (Exception ex) { sCheckError = ex.Message; }
                        pptxModule.quit();
                    }
                } else if (sExt == ".md") {
                    bOpened = mdModule.open(sFilePath);
                    if (bOpened) {
                        try { mdModule.checkAll(sFilePath); }
                        catch (Exception ex) { sCheckError = ex.Message; }
                        mdModule.quit();
                    }
                }

                if (!bOpened) {
                    string sReason = "could not open";
                    logger.warn("Could not open " + sFilePath);
                    lsFailed.Add(new failure(sBase, sReason));
                    if (!bGuiOrHidden) Console.WriteLine(": " + sReason);
                    continue;
                }
                if (sCheckError != null) {
                    string sReason = firstLine(sCheckError);
                    logger.error("Check failed for " + sFilePath + ": " + sCheckError);
                    lsFailed.Add(new failure(sBase, sReason));
                    if (!bGuiOrHidden) Console.WriteLine(": " + sReason);
                    continue;
                }

                results.writeCsv(sCsvPath);
                logger.info("Wrote " + sCsvPath + " (" + results.lIssues.Count + " issue(s))");
                lsChecked.Add(sBase);
                if (!bGuiOrHidden) Console.WriteLine();
                iTotalIssues += results.lIssues.Count;
            }
            } finally {
                if (bGuiOrHidden) guiProgress.close();
            }

            // ---- Final results summary ----
            //
            // Three sections, each printed only when its count is
            // non-zero. Singular "file" when count == 1; plural
            // "files" otherwise. In CLI mode the basenames have
            // already scrolled by inline; the section headers are
            // still printed but the per-name lists are suppressed
            // (since they would just repeat). In GUI/hidden mode
            // the captured stdout is the final MessageBox, so we
            // include the per-name lists there.
            int iChecked = lsChecked.Count;
            int iFailed = lsFailed.Count;
            int iSkippedExisting = lsSkippedExisting.Count;

            if (iChecked > 0) {
                Console.WriteLine();
                Console.WriteLine("Checked " + iChecked + " " +
                    (iChecked == 1 ? "file" : "files") + ":");
                if (bGuiOrHidden) {
                    foreach (string sName in lsChecked)
                        Console.WriteLine(sName);
                }
            }
            if (iFailed > 0) {
                Console.WriteLine();
                Console.WriteLine("Failed to check " + iFailed + " " +
                    (iFailed == 1 ? "file" : "files") + ":");
                if (bGuiOrHidden) {
                    foreach (var oFail in lsFailed) {
                        if (string.IsNullOrEmpty(oFail.sReason))
                            Console.WriteLine(oFail.sBase);
                        else
                            Console.WriteLine(oFail.sBase + ": " + oFail.sReason);
                    }
                }
            }
            if (iSkippedExisting > 0) {
                Console.WriteLine();
                Console.WriteLine("Skipped " + iSkippedExisting + " " +
                    (iSkippedExisting == 1 ? "file" : "files") +
                    ". Check \"Force replacements\" to overwrite.");
            }
            if (iChecked == 0 && iFailed == 0 && iSkippedExisting == 0) {
                Console.WriteLine();
                Console.WriteLine("No supported files to check.");
            }

            if (bViewOutput && (iChecked > 0 || bShowRules))
                openOutputInExplorer(sResolvedOutDir);

            return (iChecked > 0 || bShowRules) ? 0 : 1;
        }
        finally {
            // Restore stdout/stderr and surface captured output.
            if (bGuiMode && stringWriterOut != null) {
                Console.SetOut(writerOriginalOut);
                Console.SetError(writerOriginalErr);
                showFinalMessage(stringWriterOut.ToString());
            } else if (bHideConsoleMode && stringWriterErr != null) {
                Console.SetError(writerOriginalErr);
                string sErr = stringWriterErr.ToString();
                if (!string.IsNullOrWhiteSpace(sErr))
                    showFinalMessage(sErr, "extCheck — Errors");
            }
            logger.info("Run complete.");
            logger.close();
        }
    }
}
