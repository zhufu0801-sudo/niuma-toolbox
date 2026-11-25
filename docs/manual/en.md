# Office Toolbox Detailed User Manual

Source date: 2026-11-25. A4, 11-point body. [中文](zh.md) / [日本語](ja.md) / English.

## Contents

1. [Introduction and reading guide](#intro)
2. [Installation and removal](#install)
3. [Interface and settings](#interface)
4. [WPS Writer](#word)
5. [WPS Spreadsheets](#sheet)
6. [WPS Presentation](#slides)
7. [Files and PDF](#files)
8. [Printing and fax preparation](#print)
9. [Dates and business documents](#dates)
10. [Paste bar and content cards](#cards)
11. [Masking and restoration](#mask)
12. [Macro checking and VBA assistant](#macro)
13. [Japanese postal codes and addresses](#postal)
14. [Optional AI features](#ai)
15. [Microsoft Office differences](#office)
16. [Privacy and local data](#privacy)
17. [Backups undo and processing limits](#limits)
18. [Troubleshooting](#faq)
19. [Appendix](#appendix)

<a id="intro"></a>
## 1 Introduction and reading guide

This manual is for first-time Office Toolbox users. Start with the chapter for WPS Writer, Spreadsheets, or Presentation, or search for a button in the contents. Each feature has five parts: purpose, steps, options, limits, and recovery. The English manual does not imply an English product interface; the current UI languages are Chinese and Japanese.

Ordinary formatting, data cleanup, templates, PDF tools, and masking run locally without AI. Content is sent to an AI provider only when you configure a service and actively use it. External agents and web AI have separate data paths, explained in the masking chapter. Postal-data updates, version checks, and initial runtime downloads can use the network without uploading documents.

Begin with a saved, small sample. Text selections, the active cell, and the current slide determine many operation targets. Whole-document, batch, and file-generation operations may extend beyond the selection. Read confirmations, inspect results, and then save. The toolbox does not automatically send fax or email or start physical printing.

This manual covers version 0.16.0. The Mac and Microsoft Office editions have not completed full testing on real machines. Please report problems through GitHub Issues.

<a id="install"></a>
## 2 Installation and removal

### Windows

1. Extract the complete Windows 64-bit package. Do not run it inside a ZIP.
2. Run Start.cmd and select WPS, Microsoft Office, or both. Keep the installation in a stable directory writable by the current user.
3. Office asks you to trust a certificate used only for localhost. Check its name and decide yourself; ask IT on managed devices.
4. Save documents, fully exit WPS or Office, and reopen it. Open Office Toolbox from the WPS ribbon or the Office add-in entry.
5. After starting the computer, launch the desktop Office Toolbox shortcut to start its local service. Installation does not configure automatic startup.

Before upgrading, back up generated and important documents. Update program files, repeat registration, and fully close all WPS windows before reopening them. An old sidebar can retain cached scripts.

### macOS

1. Extract the Mac package. Right-click install.command in Finder and open it.
2. Select WPS or Microsoft Office. If Node.js is missing, the installer asks whether to obtain it from nodejs.org.
3. You or IT decide whether to trust the Office localhost certificate. macOS may ask for your login password.
4. Fully quit and reopen the host. In Office, use Insert, Add-ins, My Add-ins if necessary.

Mac WPS add-in support varies by version. System OCR needs macOS 11 or later for Chinese and 13 or later for Japanese. Features such as PDF image extraction may be unavailable without their image component. This round did not include a Mac desktop test.

### Enterprise deployment

Deploy per user; the background service is not a shared multi-user server. Initialize.ps1 accepts -Targets wps, office, or wps,office. -NoLaunch suppresses opening the app and -NoShortcut suppresses the desktop shortcut. Office certificate confirmation can prevent a fully unattended deployment. Fixed ports are 127.0.0.1:39127 and, for Office, 39128. Concurrent signed-in users on one computer can conflict.

preset/settings.json supplies initial UI and font settings; preset/ai.json supplies AI connection and model settings; preset/plugins supplies company plugins. Presets copy only when the destination file does not exist, so they do not enforce overwrites. Read the enterprise guide in the package and pilot with a test account.

### Removal and retained data

On Windows, run Uninstall.cmd. It removes the toolbox registrations, its own certificates and shortcuts, and generated/ai.json containing saved AI settings and the Key. Cards and ordinary settings remain. On Mac, follow uninstall.command. Uninstalling does not erase personal documents; verify backups before deleting retained folders.

<a id="interface"></a>
## 3 Interface and settings

### Where to start

The ribbon shows tools for the current host. A large button, menu item, and sidebar item may lead to the same operation. Common tools contains search and guidance; Create accepts tables and outlines; Cards manages reusable content; Files contains PDF, templates, printing, masking, and related panels. Scroll down and expand collapsed headings in a narrow sidebar.

Search the navigator by a task, such as postal lookup, masking, or a wide table. Results can be executable tools or guides. A guide explains steps without changing your file. Star an item to add it to favourites and the ribbon favourites group; remove the star to remove it. Choosing which functions to display hides entries without deleting files or uninstalling functionality.

### Settings and a first trial

1. Open Settings and choose Chinese, Japanese, or follow system.
2. Review default body font, size, spacing, table-header colours, and card count. These affect the relevant tool output.
3. Check the backup directory and automatic-backup setting. Save the document and run Back up current file to confirm the location.
4. Enter an AI endpoint, model, and Key only if you want AI. Ordinary tools do not require it.
5. If versions disagree after an update, save and close every WPS window before reopening. Closing only the sidebar may leave old ribbon scripts active.

Changing the UI language does not translate document content, filenames, or custom word lists. Example sheet names do not need to exist: capture your real selection to use its actual sheet name. Diagnostic summaries help report the environment. An unset version-information URL does not mean installation failed; the toolbox does not download or install updates automatically.

<a id="word"></a>
## 4 WPS Writer

### Selection and structure

Apply heading styles first, format body text next, and update the contents last. Text that looks like a heading may have no heading level. The contents tool uses levels 1–3; manually typed page numbers do not update. Body formatting skips recognized headings, but complex templates still need a trial copy.

Quick replacement, blank-paragraph cleanup, and prose conversion operate only on supported body ranges. Headers, fields, tables, and tracked changes can restrict them. Distinguish floating from inline pictures. Crop the background before setting card dimensions; the outer edge of a photographed scene is not the edge of the card.

### Related small tools

| Feature and purpose | Steps and options | Limits and recovery |
|---|---|---|
| <a id="tool-addCjkSpace"></a>Add CJK spacing · 中英文间加空格 — Insert a half-width space between Chinese text and Latin letters or digits. | 1. Select the text to normalize. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Writer。 Undo and recovery: Try Ctrl+Z and verify; otherwise restore a backup. |
| <a id="tool-removeCjkSpace"></a>Remove CJK spacing · 去掉中英文空格 — Remove spaces between Chinese text and Latin letters or digits. | 1. Select the text to normalize. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Writer。 Undo and recovery: Try Ctrl+Z and verify; otherwise restore a backup. |
| <a id="tool-heading1"></a>Heading 1 · 一级标题 — Apply the first heading level to selected paragraphs. | 1. Select the chapter heading paragraphs. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Writer。 The shared text-selection limit is 50,000 characters. Undo and recovery: Try Ctrl+Z and verify; otherwise restore a backup. |
| <a id="tool-heading2"></a>Heading 2 · 二级标题 — Apply the second heading level to selected paragraphs. | 1. Select the section heading paragraphs. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Writer。 The shared text-selection limit is 50,000 characters. Undo and recovery: Try Ctrl+Z and verify; otherwise restore a backup. |


<a id="tool-deleteBlankPages"></a>
### Remove blank pages · 删除空白页

**Purpose** Remove pages that contain no content after confirmation.

**Steps**

1. Open the document and check the number of pages proposed for deletion.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-officialFormatCn"></a>
### Chinese official document layout · 公文格式（国标）

**Purpose** Apply the supported margins, fonts, lines, and page numbers for Chinese official documents.

**Steps**

1. Save a copy and inspect the entire document after applying the layout.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-businessFormatJa"></a>
### Japanese business document layout · 日文商务文书格式

**Purpose** Apply A4 paper, 10.5-point Mincho text, a centred title, and supported Japanese closing conventions.

**Steps**

1. Save a copy and check the title and closing lines before applying.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-fixPageJump"></a>
### Fix unexpected page jumps · 段落乱跳页修复

**Purpose** Clear paragraph settings that force unexpected page breaks.

**Steps**

1. Select the affected paragraphs and inspect pagination after processing.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-landscapePages"></a>
### Landscape selection · 选中内容改横向

**Purpose** Place selected content in a landscape section while keeping surrounding pages portrait.

**Steps**

1. Select the complete content to put on landscape pages.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-tableBeautify"></a>
### Format Word tables · 表格一键美化

**Purpose** Apply borders, header formatting, alignment, and supported picture fitting.

**Steps**

1. Select the table content to format.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-pictureFix"></a>
### Stabilize pictures · 图片不乱跑

**Purpose** Convert supported floating pictures to inline placement and correct clipped display.

**Steps**

1. Select content containing the affected pictures.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-boldInlineHeads"></a>
### Bold inline headings · 加粗行内标题

**Purpose** Bold the heading portion before a colon in recognized numbered lines.

**Steps**

1. Select the paragraphs containing inline headings.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-alignOptions"></a>
### Align answer choices · 选项 ABCD 对齐

**Purpose** Use tab stops to align multiple-choice options.

**Steps**

1. Select the lines containing A, B, C, and D choices.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-noBreakLatin"></a>
### Keep Latin text together · 数字英文不断行

**Purpose** Prevent supported Latin words, URLs, and digit strings from breaking across lines.

**Steps**

1. Select the paragraphs containing those strings.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-removeWatermark"></a>
### Remove watermarks · 一键去水印

**Purpose** Remove supported header watermarks and page backgrounds.

**Steps**

1. Save a copy and check headers and backgrounds after processing.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-wordToPpt"></a>
### Word to presentation · Word 转 PPT

**Purpose** Create a presentation from document heading levels.

**Steps**

1. Use real heading styles to organize the document, then start conversion.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-sendCheck"></a>
### Pre-send document check · 发送前检查

**Purpose** Report comments, revisions, hidden text, highlighting, and author information.

**Steps**

1. Save a copy; inspect the report before choosing any cleanup.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-proofCheck"></a>
### Spelling and punctuation checks · 错别字与标点检查

**Purpose** Flag supported rule-based errors and punctuation problems.

**Steps**

1. Select text and choose comments, safe fixes plus comments, or removal of checker comments.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Mode：add comments / safe fixes and comments / remove checker comments; default add comments

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-jaStyle"></a>
### Japanese prose style · 日文文体转换

**Purpose** Convert supported sentence endings between polite and plain Japanese.

**Steps**

1. Select Japanese prose and choose polite, da, or dearu style.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Target style：plain style / polite style; default plain style

Plain-style ending：dearu / da; default dearu

**Limits** Local rules and tokenization change supported sentence endings, not meaning. Quoted text inside Japanese corner brackets is preserved. Review ambiguous sentences, unusual honorifics, and literary expressions.

Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-compareDocs"></a>
### Compare two versions · 比较两个文档

**Purpose** Create a document showing changes between two versions.

**Steps**

1. Open one version and select the second file when prompted.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-exportComments"></a>
### Export comments and revisions · 批注修订汇总导出

**Purpose** Summarize comments and revisions in a new document table.

**Steps**

1. Open the reviewed document and run the export.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-captionNumbers"></a>
### Automatic figure and table captions · 图表自动编号

**Purpose** Insert supported automatic figure and table captions.

**Steps**

1. Choose pictures, tables, or both; check caption placement afterwards.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Caption targets：pictures and tables / pictures / tables; default pictures and tables

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-fixListNumbering"></a>
### Repair heading numbering · 修复多级编号

**Purpose** Rebuild automatic numbering from heading levels.

**Steps**

1. Apply heading styles, choose a numbering scheme, and decide whether to remove typed prefixes.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Numbering scheme：1 / 1.1 / 1.1.1 / Chinese numbering / chapter numbering; default 1 / 1.1 / 1.1.1

Remove typed prefixes：yes / no; default yes

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-diagonalHeader"></a>
### Diagonal header · 斜线表头

**Purpose** Create a supported diagonal table header.

**Steps**

1. Select the upper-left table cell and follow the header-text prompt.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets / WPS Writer。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. 


<a id="tool-replaceText"></a>
### Replace selected text · 查找替换

**Purpose** Replace supported matching text in the current selection.

**Steps**

1. Select the source text or cells, then enter find and replacement text.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Writer supports normal body selections without tracked changes, tables, fields, headers, or footers, up to 500 matches. Find text is limited to 200 characters and replacement text to 1,000; neither may contain paragraph or table separators.

Hosts: WPS Spreadsheets / WPS Writer。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. 


<a id="tool-insertTable"></a>
### Insert a table · 快速插入表格

**Purpose** Create a table from tab-separated columns and newline-separated rows.

**Steps**

1. Enter the table text in the Create panel and choose the insertion point.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Table text; shorter rows are padded with empty cells.

**Limits** Maximum 20 rows, 10 columns, and 10,000 characters. Writer inserts after the selection; Presentation adds table slides after the current slide.

Hosts: WPS Presentation / WPS Writer。

**Undo and recovery** Writer: try Ctrl+Z and verify. Slides: some operations have no undo; retain the original.


<a id="tool-paraNormalize"></a>
### Normalize paragraphs · 段落规范化

**Purpose** Replace space-based indents and manual breaks with paragraph formatting.

**Steps**

1. Select the body paragraphs; check headings after processing.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-tableOnePage"></a>
### Fit a table on one page · 表格压到一页

**Purpose** Try to pull a table that spills onto the next page back to one page.

**Steps**

1. Select the affected table and check readability after processing.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-fitOneLine"></a>
### Fit text on one line · 文字挤成一行

**Purpose** Tighten character spacing to bring a slightly overlong line back together.

**Steps**

1. Select the short text that needs adjustment.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-photoToPrint"></a>
### Clean a document photo · 照片转打印稿

**Purpose** Reduce the grey background of a photographed document for printing.

**Steps**

1. Choose the document image and inspect the generated result.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-signatureClean"></a>
### Remove a signature background · 签名去白底

**Purpose** Make the white background of a signature image transparent.

**Steps**

1. Choose the signature image and check the resulting edges.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-removeHeaderLine"></a>
### Remove the header rule · 删除页眉横线

**Purpose** Remove the line below a header while preserving header text.

**Steps**

1. Open the document and inspect its header afterwards.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-deleteAnswers"></a>
### Clear bracketed answers · 删除括号里的答案

**Purpose** Clear supported answer characters inside exam brackets.

**Steps**

1. Select the exercise text and check which brackets were changed.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-contractSign"></a>
### Align contract signatures · 甲乙双方落款对齐

**Purpose** Use tab stops to place two parties' signature blocks side by side.

**Steps**

1. Select the signature lines to align.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-tablePictureFit"></a>
### Fit table pictures · 表格插图不变形

**Purpose** Fix table column widths and proportionally shrink pictures to fit.

**Steps**

1. Select the table containing the pictures.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-superSubscript"></a>
### Chemical and unit scripts · 化学式和单位上下标

**Purpose** Apply supported superscripts and subscripts to units and chemical expressions.

**Steps**

1. Select text such as m2 or H2O.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-alignSignature"></a>
### Right-align signature lines · 落款靠右

**Purpose** Right-align selected closing lines and remove alignment spaces.

**Steps**

1. Select only the closing or signature lines.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-bodyFormat"></a>
### Format selected body text · 正文排版

**Purpose** Unify body font and spacing while skipping recognized headings.

**Steps**

1. Select the body text; Ctrl+A may include the whole document because recognized headings are skipped.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Font, size, and spacing come from text settings. The confirmation lists affected paragraphs and skipped headings.

**Limits** Recognized headings are skipped. Kana in the text also affects the font choice. Ordinary blank paragraphs may be removed; blanks inside tables are excluded. The shared text-selection limit is 50,000 characters.

Hosts: WPS Writer。 The shared text-selection limit is 50,000 characters.

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-toc"></a>
### Insert or update a contents table · 插入 / 更新目录

**Purpose** Build a table of contents from heading levels one to three.

**Steps**

1. Apply heading styles first; place the cursor for a new contents table or run to update existing ones.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-textTable"></a>
### Text to table · 文字转表格

**Purpose** Convert tab-separated text into a Word table.

**Steps**

1. Select text that uses tabs between columns and paragraph breaks between rows.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。 The shared text-selection limit is 50,000 characters.

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-removeBlankParas"></a>
### Remove empty paragraphs · 清理空段落

**Purpose** Remove complete empty paragraphs inside the selection.

**Steps**

1. Select normal body text without tables or tracked changes.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。 The shared text-selection limit is 50,000 characters.

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-unifyPictures"></a>
### Unify picture width · 图片统一大小

**Purpose** Use the first inline picture's width for the other selected pictures.

**Steps**

1. Select at least two inline pictures; their heights keep their own proportions.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Only 2–500 inline pictures in the selection. Width follows the first picture. Convert floating pictures to inline first.

Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-fitPictures"></a>
### Fit pictures within margins · 图片不超出页面

**Purpose** Shrink inline pictures that exceed the available text width.

**Steps**

1. Select the content containing inline pictures.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Up to 500 inline pictures; oversized images shrink to the section text width. Table column width still needs review.

Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-idCardSize"></a>
### Set card dimensions · 证件照按实际尺寸

**Purpose** Set one or two card images to 8.56 by 5.4 centimetres.

**Steps**

1. Crop away the background first and select only the card images.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Only one or two inline pictures, with an aspect ratio within 6% of 1.59:1. Print at actual size and measure a trial copy.

Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="sheet"></a>
## 5 WPS Spreadsheets

### Values text and display formats

IDs and postal codes are not amounts. Keep long IDs as text; digits already lost by numeric storage cannot be recovered by conversion. Large-number units change display only, uppercase amounts produce text, and ranking writes formulas. Reserve empty output columns where required. Whether a result tracks source changes depends on the individual tool.

WPS spreadsheet tool writes cannot be undone with Ctrl+Z. Automatic backup copies the saved file on disk and excludes unsaved changes. Save and make a manual backup before deleting rows, replacing text, renaming, or clearing data.

### Related small tools

| Feature and purpose | Steps and options | Limits and recovery |
|---|---|---|
| <a id="tool-extractNumbers"></a>Extract digits · 提取数字 — Write extracted digits in the empty column to the right. | 1. Select one column and reserve a blank output column. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits. Undo and recovery: No Ctrl+Z recovery; restore a backup or remove a newly created result sheet. |
| <a id="tool-extractChinese"></a>Extract CJK text · 提取中文 — Keep Chinese and Japanese characters in a new output column. | 1. Select one column and reserve a blank output column. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits. Undo and recovery: No Ctrl+Z recovery; restore a backup or remove a newly created result sheet. |
| <a id="tool-extractLetters"></a>Extract Latin letters · 提取英文 — Keep Latin letters and spaces in a new output column. | 1. Select one column and reserve a blank output column. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits. Undo and recovery: No Ctrl+Z recovery; restore a backup or remove a newly created result sheet. |
| <a id="tool-removeBrackets"></a>Remove bracketed text · 去掉括号内容 — Remove supported brackets and their contents into a new output column. | 1. Select one column and reserve a blank output column. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits. Undo and recovery: No Ctrl+Z recovery; restore a backup or remove a newly created result sheet. |
| <a id="tool-formatHeader"></a>Format a header row · 美化表头 — Emphasize the first row of the selected range. | 1. Select the range including its header. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits. Undo and recovery: No Ctrl+Z recovery; restore a backup or remove a newly created result sheet. |
| <a id="tool-fitColumns"></a>Auto-fit column widths · 自动列宽 — Adjust column widths to their contents. | 1. Select the columns or cells to fit. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits. Undo and recovery: No Ctrl+Z recovery; restore a backup or remove a newly created result sheet. |
| <a id="tool-wrapCells"></a>Wrap cell text · 自动换行 — Enable wrapping so cell text can display on several lines. | 1. Select the target cells. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits. Undo and recovery: No Ctrl+Z recovery; restore a backup or remove a newly created result sheet. |
| <a id="tool-filterCells"></a>Toggle filter · 筛选开关 — Enable or remove filtering on the selected data. | 1. Select the data range including headers. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits. Undo and recovery: No Ctrl+Z recovery; restore a backup or remove a newly created result sheet. |
| <a id="tool-trimCells"></a>Trim cell whitespace · 清理空格 — Remove leading and trailing whitespace from text cells. | 1. Select the text cells to clean. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits. Undo and recovery: No Ctrl+Z recovery; restore a backup or remove a newly created result sheet. |
| <a id="tool-centerAcross"></a>Centre across selection · 跨列居中 — Centre a title across cells without merging them. | 1. Select the title and adjacent cells. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits. Undo and recovery: No Ctrl+Z recovery; restore a backup or remove a newly created result sheet. |
| <a id="tool-shrinkToFit"></a>Shrink text to fit · 缩小字体填充 — Reduce displayed text size without changing column width. | 1. Select the target cells. 2. Open this tool from the ribbon or navigator, review its options, and confirm. 3. Inspect the result and its location before saving. Options: No extra parameters. | Limits: Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits. Undo and recovery: No Ctrl+Z recovery; restore a backup or remove a newly created result sheet. |


<a id="tool-idCardInfo"></a>
### ID card details · 身份证提取生日性别

**Purpose** Validate a Chinese ID number and extract its birth date and sex fields.

**Steps**

1. Select the ID number cells with their full digits preserved.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-phoneCheck"></a>
### Check mobile numbers · 检查手机号

**Purpose** Check mobile number length and supported prefixes.

**Steps**

1. Select the column of phone numbers.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-furigana"></a>
### Show phonetic readings · 加注假名（读音）

**Purpose** Use PHONETIC to display readings stored with cell input.

**Steps**

1. Select cells that already contain phonetic input information.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** PHONETIC reads phonetic metadata stored during input. It cannot reliably invent readings for pasted Kanji.

Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-sheetCheck"></a>
### Spreadsheet diagnostic · 表格体检

**Purpose** Explain common lookup, sum, and numeric-display problems without changing data.

**Steps**

1. Select the affected data and review the diagnostic report.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** The check makes no edits; review any subsequent cleanup separately.


<a id="tool-wanUnit"></a>
### Display large-number units · 万元显示

**Purpose** Display amounts in selected units without changing their stored values.

**Steps**

1. Select numeric cells and choose the display unit.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Display unit：ten thousand / hundred million / thousand / million / thousand yen / normal numbers; default ten thousand

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-idCheck"></a>
### Validate identifiers · 号码校验

**Purpose** Check supported identifier lengths, formats, and checksums.

**Steps**

1. Select identifiers stored without lost digits and choose their type.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Identifier type：automatic / Chinese unified social credit code / Chinese ID / invoice number / Japanese corporate / invoice ID / bank card; default automatic

**Limits** Checks implemented formats, lengths, and check digits only. It does not verify existence, validity, or ownership online. Store long IDs as text.

Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-mergeWorkbooks"></a>
### Merge workbooks · 合并多个工作簿

**Purpose** Stack data aligned by headers or copy complete worksheets into a result workbook.

**Steps**

1. Select source workbooks and choose data stacking or sheet copying.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Mode：stack data by headers / copy sheets; default stack data by headers

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-codesBatch"></a>
### Batch QR codes and barcodes · 批量二维码 / 条形码

**Purpose** Generate a code image for each selected value and place it to the right.

**Steps**

1. Select one column of URLs or identifiers and choose QR or Code 128.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Code type：QR code / Code 128; default QR code

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-attendanceStats"></a>
### Attendance counts · 考勤统计

**Purpose** Count attendance marks for each person.

**Steps**

1. Select an attendance table containing names and daily marks.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-spotlight"></a>
### Row and column spotlight · 聚光灯

**Purpose** Highlight the current row and column.

**Steps**

1. Activate a cell; click again to turn the effect off.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-bandRows"></a>
### Alternating row colours · 隔行填色

**Purpose** Apply an alternating-row conditional format.

**Steps**

1. Select at least three rows, including the header; click again to remove the effect.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-textToNumber"></a>
### Convert numeric text · 文本转数字

**Purpose** Convert eligible numeric and date text into real values.

**Steps**

1. Select a range without merged cells; review the conversion count before confirming.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Formulas and non-text cells are skipped. Leading-zero IDs and IDs longer than 15 digits remain text. Digits lost by previous numeric conversion cannot be recovered.

Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-compareRanges"></a>
### Compare two ranges · 两表核对

**Purpose** Colour differences between two selected ranges.

**Steps**

1. Select one range, hold Ctrl, and select a second range on the same sheet.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-payrollSlips"></a>
### Payslip layout · 工资条

**Purpose** Create a new sheet with a header and data row for each person.

**Steps**

1. Select the payroll table including its single header row; do not include merged cells.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-mergedNumbering"></a>
### Number merged cells · 合并单元格编号

**Purpose** Fill differently sized merged cells with consecutive numbers.

**Steps**

1. Select the single column to number and review any values that will be overwritten.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-rankKeepOrder"></a>
### Rank without sorting · 排名不打乱顺序

**Purpose** Write descending RANK formulas to the right without moving source rows.

**Steps**

1. Select one numeric column without its header and leave the right column empty.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-dueReminder"></a>
### Due-date highlighting · 到期提醒

**Purpose** Mark past dates red and dates within 30 days yellow.

**Steps**

1. Select one column of real date values without its header; click again to remove the rule.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-progressBar"></a>
### In-cell progress bars · 进度条

**Purpose** Display numeric values as conditional-format data bars.

**Steps**

1. Select numeric or percentage cells.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-joinColumns"></a>
### Join columns · 多列合一格

**Purpose** Join each row into one cell to the right.

**Steps**

1. Select the source columns, leave output space empty, and choose a separator.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Separator

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-splitToRows"></a>
### Split a list into rows · 一格拆成多行

**Purpose** Create rows from delimited cell contents and copy the remaining row fields.

**Steps**

1. Select the source data and inspect the new result sheet.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-stackColumns"></a>
### Stack columns · 多列堆成一列

**Purpose** Create a new sheet containing selected columns stacked into one column.

**Steps**

1. Select the columns to stack.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-deleteBlankRows"></a>
### Delete blank rows · 删除空白行

**Purpose** Remove rows that are blank within the selected range.

**Steps**

1. Select the bounded data range and inspect the deletion confirmation.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-sheetIndex"></a>
### Worksheet index · 工作表目录

**Purpose** Create a linked index sheet for the workbook.

**Steps**

1. Open the workbook whose sheets need an index.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-renameSheets"></a>
### Rename worksheets · 批量改工作表名

**Purpose** Replace text and add prefixes or suffixes in worksheet names.

**Steps**

1. Enter the name change and review it before confirming.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Find text

Replacement text (blank deletes)

Prefix

Suffix

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-splitByColumn"></a>
### Split by category · 按列拆分成多表

**Purpose** Create one worksheet for each value in the current column.

**Steps**

1. Activate the category column in the data table.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-mergeSheets"></a>
### Merge sheets · 多表合并

**Purpose** Stack worksheets with matching headers in a new sheet.

**Steps**

1. Check that source worksheets use the same header structure.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-quickChart"></a>
### Create a chart · 一键图表

**Purpose** Choose a column or line chart from the selected data.

**Steps**

1. Select labels and values; Ctrl-select nonadjacent columns if needed.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-swapColumns"></a>
### Swap two columns · 两列互换

**Purpose** Exchange the contents of two selected columns.

**Steps**

1. Select two adjacent columns or Ctrl-select two separate columns.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-transposeRange"></a>
### Transpose rows and columns · 行列转置

**Purpose** Create a transposed copy in a new sheet.

**Steps**

1. Select the rectangular source range.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-alignNames"></a>
### Align names · 姓名对齐

**Purpose** Use distributed alignment without inserting spaces into names.

**Steps**

1. Select the name cells.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-colorSumCount"></a>
### Summarize by fill colour · 按颜色求和计数

**Purpose** Count cells and sum values by their fill colour.

**Steps**

1. Select the coloured cells to summarize.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-addPrefixSuffix"></a>
### Add prefixes and suffixes · 批量加前后缀

**Purpose** Add text before or after selected cell contents.

**Steps**

1. Select the cells and enter the prefix and suffix.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Prefix

Suffix

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-wrapIferror"></a>
### Hide formula errors · 公式错误显示为空

**Purpose** Wrap selected formulas in IFERROR with a blank fallback.

**Steps**

1. Select the formulas to change and retain a backup of the originals.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-maskData"></a>
### Mask cell values · 敏感信息打码

**Purpose** Replace the middle of supported sensitive values with asterisks.

**Steps**

1. Select the sensitive cells; use the full masking panel when you need a restoration table.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** This shortcut writes asterisks into cells; it is not the encrypted mapping workflow in the masking panel. Use that panel and retain the original if restoration is needed.

Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-unitSum"></a>
### Sum text with units · 带单位求和

**Purpose** Sum supported numeric text that includes unit suffixes.

**Steps**

1. Select values such as 12元 or 3.5kg and inspect the output units.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-listToTable"></a>
### Turn a list into a table · 一列整理成表格

**Purpose** Arrange a repeated vertical list into one record per row.

**Steps**

1. Select the list and enter the number of lines per record.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Lines per record

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-picturesByName"></a>
### Insert pictures by filename · 按名称批量插入图片

**Purpose** Match selected names to image filenames and insert pictures to the right.

**Steps**

1. Select the name column and choose the image folder.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-removeDuplicates"></a>
### Remove duplicate rows · 删除重复行

**Purpose** Keep the first row of each duplicate set.

**Steps**

1. Select the table and inspect the duplicate criteria and confirmation.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-fillBlanks"></a>
### Fill from above · 空白填上一格

**Purpose** Fill blank cells from the preceding cell in the same column.

**Steps**

1. Select the bounded range and verify that blank cells should inherit the value above.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-fitPicturesToCells"></a>
### Fit pictures to cells · 图片适应单元格

**Purpose** Resize and centre supported pictures within selected cells.

**Steps**

1. Select the cells containing the pictures.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-autoBorder"></a>
### Automatic row borders · 输入后自动加边框

**Purpose** Add conditional borders to rows when they contain data.

**Steps**

1. Select the data-entry area.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="formula-panel"></a>
### Local formula assistant

**Purpose** Preview formulas for common tasks and write them to blank destinations linked to source data.

**Steps**

1. Open Common formulas and describe the task or choose a type
2. Capture source, lookup table, and destination separately
3. Verify key and return columns, conditions, and fallback text
4. Preview formulas and references; inspect sample formulas and warnings
5. Confirm writing, then verify recalculation when source values change

**Options** Types include compatible lookup, XLOOKUP, reverse/multiple-condition/one-to-many lookup, SUMIF/SUMIFS, COUNTIF, conditional average, IF, sum/average/nonempty count, TRIM, left/right text, ROUND, year/month/day, concatenation, date difference, distinct count, rank, whole-year age/tenure, and weekday. Relevant inputs appear for the selected type.

**Limits** No AI: intent recognition mainly uses local Chinese rules. It is not a general-language interpreter. Use explicit ranges, not entire columns, named ranges, or external-workbook references. Duplicate keys normally return the first match; choose the dedicated type for all matches. New-function support depends on the host. Writes target blank cells, up to 1,000 rows.

**Undo and recovery** Recent-write rollback retains up to 10 records. It removes only assistant formulas you have not subsequently changed and keeps formatting, without relying on WPS spreadsheet undo.


<a id="quick-selectionStats"></a>
### Copy selection statistics · 复制选区合计

**Purpose** Copy selection statistics.

**Steps**

1. Select data
2. capture the range
3. and click a sum
4. average
5. or count to copy it.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** Read-only statistics; check numeric-text cells separately.

**Undo and recovery** The check makes no edits; review any subsequent cleanup separately.


<a id="quick-filterValues"></a>
### Multi-keyword filter · 按关键词多选筛选

**Purpose** Multi-keyword filter.

**Steps**

1. Select data with headers
2. choose the column
3. tick matching values
4. and create a filtered copy.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** Checked values, not just the current search text, determine the result.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="quick-unmergeFill"></a>
### Unmerge and fill · 拆开合并并补齐

**Purpose** Unmerge and fill.

**Steps**

1. Select complete merged areas
2. preview
3. then apply to a copy.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** Partial merged areas must be corrected before processing.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="quick-mergeSame"></a>
### Merge adjacent equal values · 连续同名合并

**Purpose** Merge adjacent equal values.

**Steps**

1. Select one category column without a header
2. preview
3. and apply to a copy.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** Only adjacent, exactly equal nonempty values are merged.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="quick-dropdown"></a>
### Create dropdown choices · 制作下拉选项

**Purpose** Create dropdown choices.

**Steps**

1. Select the input area
2. enter one option per line or capture options
3. then preview and apply to a copy.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** Empty lines are ignored, duplicates removed, and choices total at most 255 characters.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="quick-preventDuplicates"></a>
### Prevent duplicate entry · 不许重复录入

**Purpose** Prevent duplicate entry.

**Steps**

1. Select one input column without its header
2. preview
3. and apply to a copy.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** Validation restricts normal typing; pasting can bypass it. Only the selected range is checked.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="quick-cleanInvisible"></a>
### Clean invisible characters · 清理看不见的字符

**Purpose** Clean invisible characters.

**Steps**

1. Select cells
2. compare before/after examples
3. and clean a copy.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** Removes supported zero-width, BOM, and control characters; preserves tabs and newlines and skips numbers and formulas.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="quick-groupSummary"></a>
### Summarize by category · 按分类汇总

**Purpose** Summarize by category.

**Steps**

1. Select a table with headers
2. choose category and amount columns
3. and create a summary sheet.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** Static output, not a pivot table; text amounts are excluded from sums.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="quick-clearConstants"></a>
### Clear values but keep formulas · 清空数据，保留公式

**Purpose** Clear values but keep formulas.

**Steps**

1. Select only the data area without headers
2. preview
3. and create a template copy.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** Review the selection carefully: constant values are the data being removed.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="quick-uniqueCount"></a>
### Count unique values · 去重并统计次数

**Purpose** Count unique values.

**Steps**

1. Select data with headers
2. choose a column
3. and create the frequency summary.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** A new static summary is produced; later source edits require running it again.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="quick-textIds"></a>
### Enter long IDs as text · 长编号按文本录入

**Purpose** Enter long IDs as text.

**Steps**

1. Select future input cells
2. preview
3. and set text format in a copy.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** Affects future input, skips formulas, and cannot restore digits already lost.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="quick-readingGuide"></a>
### Native reading highlight · 阅读时高亮行列

**Purpose** Native reading highlight.

**Steps**

1. Use WPS View
2. Reading mode or row/column highlight
3. and choose a colour.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** A guidance entry; host labels and availability depend on version.

**Undo and recovery** The check makes no edits; review any subsequent cleanup separately.


<a id="quick-renameGuide"></a>
### Rename an open file · 文件不用关闭就改名

**Purpose** Rename an open file.

**Steps**

1. Right-click the top file tab and choose Rename
2. or click its name where supported.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** A guidance entry; renaming can affect existing path references.

**Undo and recovery** The check makes no edits; review any subsequent cleanup separately.


<a id="quick-pivotGuide"></a>
### Find pivot fields · 快速找到透视表字段

**Purpose** Find pivot fields.

**Steps**

1. Click an existing pivot table
2. search its field list
3. then select or drag fields.

**Options** Use the displayed column, item, and preview controls; recheck the captured range before applying.

**Limits** Requires an existing pivot table and a host version with field search.

**Undo and recovery** The check makes no edits; review any subsequent cleanup separately.


<a id="slides"></a>
## 6 WPS Presentation

### Whole deck or current slide

Current-slide formatting and duplication use the active slide. Unifying all fonts, removing animations, and removing notes affect the whole deck: save a copy first. Applying a template does not guarantee every text block will fit. Inspect titles, tables, images, and overflowing text. Image and PDF exports are sharing formats, not editable presentation originals.

<a id="tool-slidesToImages"></a>
### Export slide images · 每页导出图片 / 长图

**Purpose** Export visible slides as PNGs and optionally a long image.

**Steps**

1. Choose separate images or images plus a long image and choose an output folder.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Mode：one image per slide / images and one long image; default one image per slide

**Limits** At most 200 slides; hidden slides are skipped. Images are 1,920 pixels wide with proportional height and contain no editable objects. Long-image stitching needs the local image component.

Hosts: WPS Presentation。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-handoutPdf"></a>
### Handout PDF · 讲义打印

**Purpose** Export several slides per page as a handout PDF where supported.

**Steps**

1. Choose one, two, three, four, six, or nine slides per page.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Slides per page：2 / 3 / 4 / 6 / 9 / 1; default 6

**Limits** Some WPS versions only set handout print options and instruct you to use Ctrl+P. The tool does not automatically print.

Hosts: WPS Presentation。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-slimDeck"></a>
### Remove unused layouts · PPT 瘦身

**Purpose** Remove unused layouts and masters without deleting slide content.

**Steps**

1. Save a copy and review the count before confirming.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Presentation。

**Undo and recovery** Some operations have no undo; recover from the original saved before processing.


<a id="tool-makeSlides"></a>
### Create slides from an outline · 按大纲制作 PPT

**Purpose** Insert new slides after the current slide using a text outline.

**Steps**

1. Open the Create panel, enter the outline, and select the insertion slide.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Presentation。

**Undo and recovery** Some operations have no undo; recover from the original saved before processing.


<a id="tool-formatSlide"></a>
### Format the current slide · 统一本页字体

**Purpose** Unify body text formatting while preserving title size.

**Steps**

1. Activate the slide to format.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Presentation。

**Undo and recovery** Some operations have no undo; recover from the original saved before processing.


<a id="tool-duplicateSlide"></a>
### Duplicate the current slide · 复制当前页

**Purpose** Create a copy of the current slide.

**Steps**

1. Click the source slide in the thumbnail pane.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Presentation。

**Undo and recovery** Some operations have no undo; recover from the original saved before processing.


<a id="tool-unifyDeckFont"></a>
### Unify presentation fonts · 全部页统一字体

**Purpose** Apply one font to text boxes and tables while preserving font sizes.

**Steps**

1. Enter a font or leave it blank to use the most common font in the deck.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Font

**Limits** Hosts: WPS Presentation。

**Undo and recovery** Some operations have no undo; recover from the original saved before processing.


<a id="tool-removeAnimations"></a>
### Remove animations · 删除全部动画

**Purpose** Remove supported animations and slide transitions throughout the deck.

**Steps**

1. Save a copy and check the reported effects before confirming.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Presentation。

**Undo and recovery** Some operations have no undo; recover from the original saved before processing.


<a id="tool-unifyTitles"></a>
### Unify slide titles · 统一各页标题

**Purpose** Match supported title position, size, and colour to the reference title.

**Steps**

1. Use at least two slides and optionally enter a title size.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Title font size

**Limits** Hosts: WPS Presentation。

**Undo and recovery** Some operations have no undo; recover from the original saved before processing.


<a id="tool-slideNumbers"></a>
### Slide numbers · 加页码

**Purpose** Add slide numbers at the lower right.

**Steps**

1. Choose whether to show a number on the cover slide.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Skip cover number

**Limits** Hosts: WPS Presentation。

**Undo and recovery** Some operations have no undo; recover from the original saved before processing.


<a id="tool-bodySpacing"></a>
### Unify body line spacing · 统一正文行距

**Purpose** Apply one body-text line-spacing multiplier across slides.

**Steps**

1. Enter a multiplier from 0.8 to 3, or leave blank for 1.2.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Line-spacing multiplier

**Limits** Hosts: WPS Presentation。

**Undo and recovery** Some operations have no undo; recover from the original saved before processing.


<a id="tool-arrangeVisuals"></a>
### Arrange pictures and charts · 图片图表自动排版

**Purpose** Arrange supported visual objects in a proportional grid.

**Steps**

1. Choose the current slide or all slides and optionally one to six columns.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Scope

Number of columns

**Limits** Hosts: WPS Presentation。

**Undo and recovery** Some operations have no undo; recover from the original saved before processing.


<a id="tool-applyDesign"></a>
### Apply a presentation template · 套用模板

**Purpose** Apply another presentation's supported design, colours, and layouts.

**Steps**

1. Save a copy and select the template file.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Presentation。

**Undo and recovery** Some operations have no undo; recover from the original saved before processing.


<a id="tool-removeNotes"></a>
### Remove speaker notes · 删除全部备注

**Purpose** Clear supported speaker-note text throughout the presentation.

**Steps**

1. Save a copy and inspect the note count before confirming.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Presentation。

**Undo and recovery** Some operations have no undo; recover from the original saved before processing.


<a id="studio"></a>
### PPT layout studio

**Purpose** Turn your outline into an editable presentation.

**Steps**

1. Open the studio and set title, subtitle, and template
2. Put a section title first, then one point per line; separate sections with a blank line
3. Optionally add brand colour, a logo, and image-page pictures
4. Generate preview, inspect every page, then download PPTX or open in WPS

**Options** Optional layout tags follow section titles: [卡片], [数据], [时间线], [引言], [对比], [图表], [图文], [要点]. Lines starting with > become speaker notes. Logo: PNG/JPG up to 2 MB; up to eight image-page pictures; outline up to 20,000 characters. AI outline generation is a separate optional action.

**Limits** Local layout does not expand content or fact-check. Split long text and match image order to image sections. Inspect fonts and slideshow playback after export.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="files"></a>
## 7 Files and PDF

### Paths and saved content

File tools read content saved on disk. Save first and provide a full path to an existing file, not just its name. Results normally go to Generated files or a selected folder. Use the result path or Open containing folder, and verify that you opened the processed copy before continuing to edit.

PDF page tools allow up to 50 merge inputs, 200 MB per PDF, 500 MB total, and 2,000 pages. Image-to-PDF accepts up to 200 images at 40 MB each. Rendering PDF pages to images is limited to 100 pages per operation. An authorized person must handle password or editing restrictions in the source application; the toolbox does not bypass protection.

<a id="tool-batchPdf"></a>
### Batch PDF export · 批量转 PDF

**Purpose** Export several Office files as separate PDFs.

**Steps**

1. Save the source files and select them in the file picker.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-batchReplace"></a>
### Batch text replacement · 批量替换文字

**Purpose** Replace text in several documents and save new copies.

**Steps**

1. Choose the source documents; enter the text to find and its replacement.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Find text

Replacement text (blank deletes)

**Limits** Hosts: WPS Writer。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-batchWatermark"></a>
### Batch text watermarks · 批量加水印

**Purpose** Add the same text watermark to document copies.

**Steps**

1. Choose the source documents and enter the watermark text.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Text or numbering pattern

**Limits** Hosts: WPS Writer。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-splitByHeading"></a>
### Split by heading · 按标题拆分文档

**Purpose** Create a separate document for each heading section.

**Steps**

1. Apply real heading styles, save the document, and choose a heading level from 1 to 3.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Heading level (1–3)

**Limits** Hosts: WPS Writer。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-sendToDoc"></a>
### Send to another document · 发送到另一个文档

**Purpose** Copy selected content, including supported formatting, to another open document.

**Steps**

1. Open the destination document in the same host; select the source content and choose the destination number if asked.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Destination number

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-batchRename"></a>
### Batch file renaming · 批量重命名

**Purpose** Add prefixes or suffixes, replace filename text, or number files.

**Steps**

1. Select files, choose the naming mode, and inspect the proposed change.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Mode：prefix / suffix / replace text / sequence number / undo last rename; default prefix

Text or numbering pattern

Find text

Start date or number

**Limits** Renaming can break path-based links. Avoid files held open by other programs. In a numbering pattern, # marks the number position.

Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Use the last-rename recovery option and its local record, not document Ctrl+Z. Check the current state first if files have since moved or been renamed.


<a id="tool-exportSheets"></a>
### Export worksheets · 工作表另存为文件

**Purpose** Save each worksheet as a separate XLSX file.

**Steps**

1. Save the workbook, then choose the output location if prompted.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-rangeToImage"></a>
### Export a range as an image · 选区存为图片

**Purpose** Save a spreadsheet range as a PNG image.

**Steps**

1. Select the range and choose the output location.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-backupFile"></a>
### Back up the current file · 备份当前文件

**Purpose** Create a timestamped backup of the saved file.

**Steps**

1. Save current changes before making the backup.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-exportPdf"></a>
### Export PDF · 导出 PDF

**Purpose** Export the current supported layout as a new PDF.

**Steps**

1. Save the source document and inspect the PDF after export.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="tool-insertFiles"></a>
### Merge Word documents · 合并多个文档

**Purpose** Append selected documents in filename order.

**Steps**

1. Save the destination and choose the source documents.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="pdf-merge"></a>
### Merge PDFs

**Purpose** Join files in list order.

**Steps**

1. Open PDF tools in Files
2. Choose at least two PDFs, reorder the list, and start.
3. Inspect output count and order

**Options** Page numbers start at 1; 8- means page 8 through the end. Merge and image inputs follow list order.

**Limits** Observe this chapter’s size and page limits. Raster images do not preserve searchable text or interactive features.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="pdf-pick"></a>
### Extract pages

**Purpose** Save selected pages in a new PDF.

**Steps**

1. Open PDF tools in Files
2. Select a PDF, enter pages such as 1-3,5,8-, and start.
3. Inspect output count and order

**Options** Page numbers start at 1; 8- means page 8 through the end. Merge and image inputs follow list order.

**Limits** Observe this chapter’s size and page limits. Raster images do not preserve searchable text or interactive features.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="pdf-remove"></a>
### Remove pages

**Purpose** Save the pages not listed for removal.

**Steps**

1. Open PDF tools in Files
2. Select a PDF and removal pages; check what will remain and start.
3. Inspect output count and order

**Options** Page numbers start at 1; 8- means page 8 through the end. Merge and image inputs follow list order.

**Limits** Observe this chapter’s size and page limits. Raster images do not preserve searchable text or interactive features.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="pdf-split"></a>
### Split PDF

**Purpose** Create files by page count or page ranges.

**Steps**

1. Open PDF tools in Files
2. Select a PDF and enter a page count per file or ranges such as 1-3,4-10,11-.
3. Inspect output count and order

**Options** Page numbers start at 1; 8- means page 8 through the end. Merge and image inputs follow list order.

**Limits** Observe this chapter’s size and page limits. Raster images do not preserve searchable text or interactive features.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="pdf-to-images"></a>
### PDF to images

**Purpose** Save each selected page as a PNG.

**Steps**

1. Open PDF tools in Files
2. Choose a PDF, pages, and resolution, then start.
3. Inspect output count and order

**Options** Page numbers start at 1; 8- means page 8 through the end. Merge and image inputs follow list order.

**Limits** Observe this chapter’s size and page limits. Raster images do not preserve searchable text or interactive features.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="pdf-images-to-pdf"></a>
### Images to PDF

**Purpose** Place JPG and PNG images on separate pages in order.

**Steps**

1. Open PDF tools in Files
2. Choose images, adjust order and paper size, and start.
3. Inspect output count and order

**Options** Page numbers start at 1; 8- means page 8 through the end. Merge and image inputs follow list order.

**Limits** Observe this chapter’s size and page limits. Raster images do not preserve searchable text or interactive features.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="ocrImage"></a>
### Image OCR

**Purpose** Recognize image text locally.

**Steps**

1. Choose an image or paste a screenshot
2. Recognize and verify amounts, IDs, and punctuation
3. Copy or insert the text

**Options** Use the panel’s file, OCR-language, or quality controls where offered.

**Limits** Processes supported saved formats. Inspect output; fonts, pictures, and pagination depend on host and system components.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="pdfToWord"></a>
### PDF to Word

**Purpose** Convert PDF content into an editable document.

**Steps**

1. Select a PDF
2. Convert, using OCR for scanned content
3. Open the result and inspect paragraphs and tables

**Options** Use the panel’s file, OCR-language, or quality controls where offered.

**Limits** Processes supported saved formats. Inspect output; fonts, pictures, and pagination depend on host and system components.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="extractImages"></a>
### Extract document images

**Purpose** Save embedded images from supported files into a folder.

**Steps**

1. Save the source
2. Choose image extraction in Files and images
3. Inspect the output folder

**Options** Use the panel’s file, OCR-language, or quality controls where offered.

**Limits** Processes supported saved formats. Inspect output; fonts, pictures, and pagination depend on host and system components.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="compressImages"></a>
### Compress document images

**Purpose** Compress images and save a document copy.

**Steps**

1. Save the document
2. Choose compression and create a copy
3. Zoom in and inspect small text and images

**Options** Use the panel’s file, OCR-language, or quality controls where offered.

**Limits** Processes supported saved formats. Inspect output; fonts, pictures, and pagination depend on host and system components.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="docHistory"></a>
### Backup history

**Purpose** Open a previously saved backup.

**Steps**

1. Open history
2. Select a candidate by time
3. Inspect it and save under a new name if needed

**Options** Use the panel’s file, OCR-language, or quality controls where offered.

**Limits** Only existing backups can be opened. Unsaved content with no backup cannot be recovered.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="mail-merge-panel"></a>
### Generate documents from a list

**Purpose** Generate one document per list row from a DOCX template and XLSX list.

**Steps**

1. Save a DOCX template with {{column name}} placeholders
2. Select an XLSX list whose first row contains column names
3. Check used, unused, and unknown placeholders
4. Choose separate files or one combined document and generate

**Options** For separate files, choose a naming column or keep sequential numbering.

**Limits** Up to 500 rows. Placeholder names must match columns; unknown placeholders are not guessed. Only saved content is read.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="print"></a>
## 8 Printing and fax preparation

<a id="tool-printPreflight"></a>
### Print preflight · 打印前检查

**Purpose** Check supported font, table-break, and print-area issues.

**Steps**

1. Open the document to print and review the reported issues.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** The check makes no edits; review any subsequent cleanup separately.


<a id="tool-printKeepRows"></a>
### Keep table rows together · 表格不跨页断行

**Purpose** Stop a table row from splitting across pages.

**Steps**

1. Select the affected table.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="tool-printSheetSetup"></a>
### Set print area and footer · 自动打印区域与页脚

**Purpose** Use the used range, one-page width, repeating headers, filename, and date for printing.

**Steps**

1. Open the sheet to print and inspect its used range.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Print gridlines

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-printTitles"></a>
### Repeat print headers · 每页打印表头

**Purpose** Repeat the header rows on each printed page.

**Steps**

1. Select the rows to use as print headers.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-fitOnePageWide"></a>
### Fit to one page wide · 打印一页宽

**Purpose** Scale printing to one page wide without limiting its height.

**Steps**

1. Open the sheet to print and check print preview afterwards.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-columnsForPrint"></a>
### Print a narrow list in columns · 细长表格分栏打印

**Purpose** Rearrange a long narrow table into side-by-side groups.

**Steps**

1. Select the list and choose two to six groups or automatic layout.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Print column groups (2–6)

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-printBigText"></a>
### Print large characters · 打印超大字

**Purpose** Create large-character pages for signs and labels.

**Steps**

1. Enter the characters to print and inspect the resulting pages.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer。

**Undo and recovery** Try Ctrl+Z and verify; otherwise restore a backup.


<a id="fax-cover"></a>
### Fax cover sheet

**Purpose** Create a document copy for review before printing.

**Steps**

1. Open the fax/envelope/shipping/template panel
2. Enter recipient, sender, fax number, subject, page count, and message.
3. Generate and inspect the copy before printing yourself

**Options** Paper choices and text fields change with the document type.

**Limits** No fax transmission, postage purchase, or automatic printing. The shipping information sheet is not an official carrier label. Verify addresses and page counts.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="address-labels"></a>
### Address labels

**Purpose** Create a document copy for review before printing.

**Steps**

1. Open the fax/envelope/shipping/template panel
2. Enter one postal code, address, and name per line, or read them from a sheet.
3. Generate and inspect the copy before printing yourself

**Options** Paper choices and text fields change with the document type.

**Limits** No fax transmission, postage purchase, or automatic printing. The shipping information sheet is not an official carrier label. Verify addresses and page counts.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="envelope"></a>
### Envelopes and postcards

**Purpose** Create a document copy for review before printing.

**Steps**

1. Open the fax/envelope/shipping/template panel
2. Choose paper, sender and recipient details, and optional postal-code boxes.
3. Generate and inspect the copy before printing yourself

**Options** Paper choices and text fields change with the document type.

**Limits** No fax transmission, postage purchase, or automatic printing. The shipping information sheet is not an official carrier label. Verify addresses and page counts.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="vertical"></a>
### Japanese vertical notices

**Purpose** Create a document copy for review before printing.

**Steps**

1. Open the fax/envelope/shipping/template panel
2. Choose a sample or enter title and body, one paragraph per line.
3. Generate and inspect the copy before printing yourself

**Options** Paper choices and text fields change with the document type.

**Limits** No fax transmission, postage purchase, or automatic printing. The shipping information sheet is not an official carrier label. Verify addresses and page counts.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="shipping"></a>
### Shipping information sheet

**Purpose** Create a document copy for review before printing.

**Steps**

1. Open the fax/envelope/shipping/template panel
2. Choose A4 or label paper and enter sender, item notes, recipient, phone, and address.
3. Generate and inspect the copy before printing yourself

**Options** Paper choices and text fields change with the document type.

**Limits** No fax transmission, postage purchase, or automatic printing. The shipping information sheet is not an official carrier label. Verify addresses and page counts.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="print-copy-panel"></a>
### Ink-saving and fax copies

**Purpose** Create monochrome or fax-optimized copies of saved files.

**Steps**

1. Save a supported source file
2. Enter its full path and choose monochrome, fax-optimized copy, or fax PDF
3. Generate and inspect page breaks, images, and thin lines

**Options** Monochrome copy, fax-optimized copy, or fax monochrome PDF; see the next entry for seam stamps.

**Limits** Fax optimization can change colours, font sizes, and lines. Raster PDF output does not preserve searchable text, links, or digital signatures.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="stamp"></a>
### Stamp images and PDF seam stamps

**Purpose** Create a stamp image or split an existing PNG across PDF page edges.

**Steps**

1. Choose shape and text for an image, or prepare an authorized transparent PNG
2. Select PDF seam stamp and enter the PDF and PNG paths
3. Generate a copy and check split order, direction, and height

**Options** Stamp shape and text. Seam stamps divide by page count and appear in sequence on right edges.

**Limits** An image provides no identity verification or digital-signature capability. Raster processing does not preserve original PDF signatures or searchable text.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="impose"></a>
### PDF multiple-up and booklet layout

**Purpose** Rearrange PDF pages for printing.

**Steps**

1. Save the PDF and open its layout controls
2. Choose multiple-up or booklet and pages per sheet
3. Generate and test-print a few pages to check duplex orientation

**Options** Choose layout and page count; set duplex in the printer dialog.

**Limits** Output is raster PDF without searchable text, links, or signatures. Paper trays and printers are not automatically controlled.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="dates"></a>
## 9 Dates and business documents

<a id="tool-approvalBox"></a>
### Approval boxes · 插入押印欄（审批栏）

**Purpose** Insert labelled approval or seal boxes.

**Steps**

1. Put the cursor at the insertion point; enter roles separated by commas.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Role labels, separated by commas

**Limits** Hosts: WPS Writer / WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. 


<a id="tool-insertToday"></a>
### Insert today's date · 插入今天日期

**Purpose** Insert today's date in the selected format.

**Steps**

1. In Writer, place the cursor at the insertion point; an active selection inserts at its end. In spreadsheets select the destination cell; in slides enter text editing.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Date format：year/month/day / year-month-day / Chinese date / Chinese date with weekday / Japanese date with weekday / Japanese era date / Short era date; default year/month/day

**Limits** Spreadsheets write only the first selected cell and ask before replacing existing content.

Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-unifyDates"></a>
### Normalize dates · 统一日期格式

**Purpose** Convert recognized date expressions to one display format.

**Steps**

1. Select text or cells containing dates and choose the output date format.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Date format：year/month/day / year-month-day / Chinese date / Chinese date with weekday / Japanese date with weekday / Japanese era date / Short era date; default year/month/day

**Limits** Writer accepts supported separated dates, Chinese dates and era dates. Eight-digit dates without separators are supported only in spreadsheets. Review ambiguous dates.

Hosts: WPS Spreadsheets / WPS Writer。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits. The shared text-selection limit is 50,000 characters.

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. 


<a id="tool-workdayCalc"></a>
### Workday calculator · 算工作日

**Purpose** Count workdays between dates or calculate a date after a workday offset.

**Steps**

1. Enter a start date, then an end date or an offset; choose the holiday region.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Start date or number

End date

Workday offset

Holiday region：China / Japan / Japan plus 29 December–3 January; default China

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-dateSummary"></a>
### Summarize by period · 按年 / 季度 / 月汇总

**Purpose** Sum a date-and-amount range by month, quarter, or year in a new sheet.

**Steps**

1. Select the date and amount columns, then choose the period.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Period：month / quarter / year; default month

**Limits** Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-insertPhrase"></a>
### Insert a phrase · 常用语一键插入

**Purpose** Insert a built-in phrase or a saved custom phrase.

**Steps**

1. In Writer place the cursor at the insertion point; a selection inserts at its end. Select a destination cell or enter slide text editing, then choose or enter a phrase.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Phrase：此致
敬礼！ / 特此通知。 / 特此证明。 / 以上报告，请审阅。 / 如有疑问，请随时与我联系。 / 感谢您的支持与配合！ / 请于    年  月  日前回复。 / 附件： / （此页无正文） / （以下无正文）; default 此致
敬礼！

Custom phrase

**Limits** Custom phrases are limited to 500 characters and saved locally. Undoing insertion does not remove saved phrases. Spreadsheets write only the first cell and confirm replacement.

Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-rmbUpper"></a>
### Uppercase RMB amount · 金额转人民币大写

**Purpose** Convert an amount to Chinese uppercase RMB wording.

**Steps**

1. In Writer select a single amount or place the cursor inside it, excluding the paragraph mark. In spreadsheets select amount cells.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Writer accepts at most two decimal places. It fills a following uppercase-amount blank when present, otherwise appends parentheses. Selecting a paragraph mark places output in the next paragraph. Existing uppercase text in parentheses is not overwritten.

Hosts: WPS Spreadsheets / WPS Writer。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. 


<a id="tool-zenHan"></a>
### Normalize character width · 全角半角统一

**Purpose** Use half-width letters and digits and full-width Katakana.

**Steps**

1. Select the text cells to normalize.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-toWareki"></a>
### Convert to Japanese era dates · 西历转和历

**Purpose** Convert recognized Gregorian dates to Japanese era notation.

**Steps**

1. Select the date cells.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-toSeireki"></a>
### Convert to Gregorian dates · 和历转西历

**Purpose** Convert recognized Japanese era dates to Gregorian notation.

**Steps**

1. Select the era-date cells, such as R8.11.12.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="library-panel"></a>
### Business and everyday document templates

**Purpose** Create editable business documents, everyday forms, and PPT copies from the library.

**Steps**

1. Choose business, everyday forms, or PPT
2. Select Chinese or Japanese and a template, then inspect preview
3. Open a copy and enter actual content
4. Verify amounts, rates, registration numbers, dates, and recipients before saving

**Options** Current library entries include PPT scenarios, business sheets, leave requests, and meeting minutes. Consumption-tax templates separate 10% and 8% and offer rounding choices.

**Limits** Templates assist entry but do not validate transactions or requirements. Replace sample companies and amounts with real data.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="cards"></a>
## 10 Paste bar and content cards

### Collect insert and link

The paste bar collects temporary pieces; named cards support repeated reuse. Both are local. Select source content before collecting, then choose the destination cursor or active cell before insertion. Do not mistake one active cell for a whole output grid. Ordinary insertion creates independent content and does not update when its source changes.

Only linked modules record source and destination positions. Open sources and review versions and destination snapshots before updating. Blocking an overwrite after manual or structural changes protects your work. Retain manual changes before reinserting. Deleting a card or paste-bar record does not remove content already inserted into a document.

<a id="tool-saveLinkedModule"></a>
### Save a linked module · 设为联动模块

**Purpose** Save content with a reference to its source.

**Steps**

1. Save the source file and select a supported text or table range.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-insertLinkedModule"></a>
### Insert a linked module · 联动插入

**Purpose** Insert a saved module and record the destination for later updates.

**Steps**

1. Choose a saved linked module and an editable destination.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-updateLinkedModules"></a>
### Update linked content · 更新联动内容

**Purpose** Refresh linked content after checking its source version.

**Steps**

1. Open the source files and inspect the proposed updates in the destination.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Sources must be readable. Changed destination snapshots, missing markers, or structural changes block overwrites and may require reinsertion. Updates are not silently applied in the background.

Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-matchFormat"></a>
### Match nearby formatting · 匹配周围格式

**Purpose** Adapt supported inserted content to nearby text or slide formatting.

**Steps**

1. Select the inserted content; in spreadsheets, follow the format-painter guidance.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets / WPS Presentation / WPS Writer。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-saveCard"></a>
### Save a content card · 设为内容模块

**Purpose** Save selected text or a table for later reuse.

**Steps**

1. Select the reusable content and enter its card details.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-insertCard"></a>
### Insert a content card · 插入内容卡片

**Purpose** Insert a saved card into the current document.

**Steps**

1. Select a card and an editable insertion point.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-capture"></a>
### Capture selected content · 放入中转区

**Purpose** Save selected content in the local transfer area.

**Steps**

1. Select supported text or a table in the source document.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-insertTransfer"></a>
### Insert transferred content · 插入中转内容

**Purpose** Insert content from the local transfer area.

**Steps**

1. Open the destination and choose an insertion point.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-insertLatest"></a>
### Insert the latest card · 插入最新卡片

**Purpose** Insert the most recently saved content card.

**Steps**

1. Open the destination and choose an insertion point.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-pasteBarCollect"></a>
### Collect in paste bar · 放进粘贴栏

**Purpose** Collect selected text or table content in the local paste bar.

**Steps**

1. Select the source content and run collection.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="tool-pasteBarInsert"></a>
### Insert from paste bar · 从粘贴栏插入

**Purpose** Insert a selected paste-bar item into the destination.

**Steps**

1. Choose an item and place the destination cursor or active cell.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Writer / WPS Spreadsheets / WPS Presentation。

**Undo and recovery** Writer: try Ctrl+Z and verify. Spreadsheets: no Ctrl+Z; restore a backup. Slides: some operations have no undo; retain the original.


<a id="mask"></a>
## 11 Masking and restoration

<a id="mask-panel"></a>
### Scan and save a masked copy

**Purpose** Replace supported personal data and word-list names with codes, asterisks, or black blocks.

**Steps**

1. Save the original and open Files, Masking
2. Select data types and a client word list, then choose the replacement mode
3. Scan and inspect counts and locations, including headers and comments
4. Keep the copy-saving default and save the masked copy, then inspect it

**Options** Choose data types and a word list. Partial masking keeps some display context; black blocks conceal the value; codes support exact restoration. Retain an original before choosing direct modification.

**Limits** Detection is not exhaustive; names absent from the word list may remain. Inspect unsupported objects and attachments. Do not rely on a masked copy’s undo history to recover original data.

**Undo and recovery** Retain the original. Use the matching file record in the restoration list and save a restored copy. Keep its mapping record.


<a id="mask-agent"></a>
### Automatic MCP agent masking

**Purpose** Give agents codes when they read through the toolbox and restore originals before writes.

**Steps**

1. Check that automatic masking for agents is enabled; it defaults to on
2. Keep the host toolbox open and read with the configured MCP client
3. Preserve the returned codes in generated text and review the host write confirmation
4. Verify that the document receives restored original values

**Options** Uses saved data types and the current word list, always in code mode. Turning the switch off exposes originals. Read results include Chinese and Japanese masking counts.

**Limits** Covers only toolbox MCP paths, not other file access by an agent. Maps are session-specific and memory-only, cleared after two idle hours or a service restart. A new or expired session cannot restore old codes; unknown codes remain and are reported. Read again before writing.

**Undo and recovery** Reading does not change the document. Write recovery depends on the host. The in-memory map is not a document backup and is not written to disk.


<a id="mask-clipboard"></a>
### Masked copy and restored paste

**Purpose** Mask at the copy boundary and restore when pasting an answer back from web AI.

**Steps**

1. Select body text, a spreadsheet range, or presentation text
2. Click masked copy, paste into the web input, and verify codes before sending
3. Copy the coded answer and place the document insertion point
4. Click restored paste; if clipboard access fails, paste into the panel and use Restore and insert
5. Check unknown-code warnings and the inserted content

**Options** Always uses codes and the panel’s current types and word list. You can copy restored text for manual insertion. If automatic copy fails, select the result box and press Ctrl+C.

**Limits** Maximum 200,000 characters and 20,000 cells. Tables copy as tabs and newlines. WPS spreadsheet restoration inserts into the active cell, not a reconstructed grid. Only the latest 20 encrypted copy maps are retained, with the newest record winning code collisions; do not mix old answers for different clients. Full/half-width brackets and internal spaces are tolerated; unknown codes remain.

**Undo and recovery** Copying does not modify the source; insertion follows host recovery rules. A deleted map cannot restore its record. The 20-entry history is not a long-term archive.


<a id="mask-modules"></a>
### Masking cards and paste-bar items

**Purpose** Apply local masking rules when saving reusable content.

**Steps**

1. Select word list, data types, and masking mode
2. Enable automatic masking for cards and paste bar, or scan existing items
3. Review matches and mask the items
4. Inspect the saved content before inserting it

**Options** Automatic processing uses current settings; scan existing items separately.

**Limits** Later edits can make a stored snapshot unsafe to restore. Changing a word list does not automatically rewrite every old card.

**Undo and recovery** Restore through the module record; items whose current snapshot differs are skipped.


<a id="macro"></a>
## 12 Macro checking and VBA assistant

<a id="vba-panel"></a>
### Static macro checking

**Purpose** Read or paste VBA and report implemented error, risk, and improvement rules without running it.

**Steps**

1. Open a macro-containing file and Files, Macro check
2. Try reading the current file’s macros
3. If access is restricted, read the instructions and copy code manually from your own VBA editor
4. Run the check and review line numbers and suggestions

**Options** Read from the file or paste manually. The report separates errors, security findings, warnings, and hints.

**Limits** Direct reading needs host VBA support and project access. The checker does not change trust settings. A clean report does not establish safety or business correctness. Supported rules include forward row deletion and missing Option Explicit.

**Undo and recovery** The check makes no edits; review any subsequent cleanup separately.


<a id="macro-panel"></a>
### VBA macro assistant

**Purpose** Generate inspectable VBA from built-in templates.

**Steps**

1. Select a range and open the VBA assistant
2. Describe the task or select a template, then capture the range
3. Choose header and sorting options, generate code, and read it
4. Copy, export .bas, or confirm installation into a new workbook module

**Options** Template, header, sort column, and order. Installation does not run code; save as .xlsm to retain macros.

**Limits** Without VBA or project access, use copy/export. Templates are designed to process a source-sheet copy; review any modified code again.

**Undo and recovery** Remove the newly added module in your VBA editor if unwanted. Undoing installation does not reverse later code execution.


<a id="postal"></a>
## 13 Japanese postal codes and addresses

<a id="tool-postalFormat"></a>
### Format postal codes · 邮编格式统一

**Purpose** Format Japanese seven-digit postal codes as 123-4567.

**Steps**

1. Select the postal-code cells.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** No extra parameters.

**Limits** Hosts: WPS Spreadsheets。

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="tool-postalFill"></a>
### Postal code and address lookup · 邮编 ↔ 住所

**Purpose** Fill addresses from postal codes or postal codes from Japanese addresses.

**Steps**

1. Select one input column and leave the output column on its right empty.
2. Open this tool from the ribbon or navigator, review its options, and confirm.
3. Inspect the result and its location before saving.

**Options** Lookup direction：automatic / postal code to address / address to postal code; default automatic

Address detail：prefecture + municipality + town / town only; default prefecture + municipality + town

**Limits** Uses the local Japan Post dataset. It cannot supply room numbers or validate recipients. Unmatched rows remain blank and are marked. Preserve leading zeros; chome ranges follow the dataset.

Hosts: WPS Spreadsheets。 Spreadsheet range: continuous and editable; at most 1,000 rows and 50 columns, subject to smaller tool limits.

**Undo and recovery** No Ctrl+Z recovery; restore a backup or remove a newly created result sheet.


<a id="postal-panel"></a>
### Address lookup and dataset updates

**Purpose** Look up Japanese addresses and postal codes with local data and update that dataset.

**Steps**

1. Open the address/postal panel from Files or the navigator
2. Enter a seven-digit code or address and inspect candidates
3. Copy or insert a chosen result
4. Update online or import an official ZIP, then read its dataset date

**Options** Lookup direction, address detail, online update, or ZIP import.

**Limits** Updates need network access; ordinary lookup is local. The date comes from the data file inside the ZIP, not the import date, and can differ from a webpage publication date. The 2026-11-25 download reported 2026-10-23 and 124,526 entries. It does not validate rooms or deliverability.

**Undo and recovery** Lookup does not edit a document; insertion follows host rules. Dataset replacement cannot be undone with document Ctrl+Z.


<a id="ai"></a>
## 14 Optional AI features

<a id="aiChat"></a>
### AI chat

**Purpose** AI chat.

**Steps**

1. Configure and test your own AI service
2. Ask a question, optionally include a selected excerpt, send, then inspect and copy or insert the answer.
3. Verify facts, numbers, and masking codes before insertion

**Options** Configure provider, endpoint, model, Key, and protection mode in AI settings; choose text-processing modes in the panel.

**Limits** Sent content goes to your chosen provider and may incur charges. Masking is not a guarantee of anonymity, and AI output is not fact verification. Local tool suggestions remain useful without AI.

**Undo and recovery** Document undo cannot retract a request already received by a provider. Recover inserted content according to host rules.


<a id="aiSummary"></a>
### Summarize

**Purpose** Summarize.

**Steps**

1. Configure and test your own AI service
2. Read selected text, request a summary, and inspect the conclusion and key points before inserting.
3. Verify facts, numbers, and masking codes before insertion

**Options** Configure provider, endpoint, model, Key, and protection mode in AI settings; choose text-processing modes in the panel.

**Limits** Sent content goes to your chosen provider and may incur charges. Masking is not a guarantee of anonymity, and AI output is not fact verification. Local tool suggestions remain useful without AI.

**Undo and recovery** Document undo cannot retract a request already received by a provider. Recover inserted content according to host rules.


<a id="aiContinue"></a>
### Continue writing

**Purpose** Continue writing.

**Steps**

1. Configure and test your own AI service
2. Read the existing passage, request a continuation, and review it before appending.
3. Verify facts, numbers, and masking codes before insertion

**Options** Configure provider, endpoint, model, Key, and protection mode in AI settings; choose text-processing modes in the panel.

**Limits** Sent content goes to your chosen provider and may incur charges. Masking is not a guarantee of anonymity, and AI output is not fact verification. Local tool suggestions remain useful without AI.

**Undo and recovery** Document undo cannot retract a request already received by a provider. Recover inserted content according to host rules.


<a id="aiExpand"></a>
### Expand text

**Purpose** Expand text.

**Steps**

1. Configure and test your own AI service
2. Read a short passage, request expansion, and review added claims before replacing or appending.
3. Verify facts, numbers, and masking codes before insertion

**Options** Configure provider, endpoint, model, Key, and protection mode in AI settings; choose text-processing modes in the panel.

**Limits** Sent content goes to your chosen provider and may incur charges. Masking is not a guarantee of anonymity, and AI output is not fact verification. Local tool suggestions remain useful without AI.

**Undo and recovery** Document undo cannot retract a request already received by a provider. Recover inserted content according to host rules.


<a id="aiShorten"></a>
### Shorten text

**Purpose** Shorten text.

**Steps**

1. Configure and test your own AI service
2. Read a passage, request a shorter version, and check that necessary details remain.
3. Verify facts, numbers, and masking codes before insertion

**Options** Configure provider, endpoint, model, Key, and protection mode in AI settings; choose text-processing modes in the panel.

**Limits** Sent content goes to your chosen provider and may incur charges. Masking is not a guarantee of anonymity, and AI output is not fact verification. Local tool suggestions remain useful without AI.

**Undo and recovery** Document undo cannot retract a request already received by a provider. Recover inserted content according to host rules.


<a id="aiBullets"></a>
### Create bullet points

**Purpose** Create bullet points.

**Steps**

1. Configure and test your own AI service
2. Read the passage, request structured points, and verify their order and meaning.
3. Verify facts, numbers, and masking codes before insertion

**Options** Configure provider, endpoint, model, Key, and protection mode in AI settings; choose text-processing modes in the panel.

**Limits** Sent content goes to your chosen provider and may incur charges. Masking is not a guarantee of anonymity, and AI output is not fact verification. Local tool suggestions remain useful without AI.

**Undo and recovery** Document undo cannot retract a request already received by a provider. Recover inserted content according to host rules.


<a id="aiFormal"></a>
### Formal wording

**Purpose** Formal wording.

**Steps**

1. Configure and test your own AI service
2. Read informal text, request formal wording, and inspect the tone before replacing.
3. Verify facts, numbers, and masking codes before insertion

**Options** Configure provider, endpoint, model, Key, and protection mode in AI settings; choose text-processing modes in the panel.

**Limits** Sent content goes to your chosen provider and may incur charges. Masking is not a guarantee of anonymity, and AI output is not fact verification. Local tool suggestions remain useful without AI.

**Undo and recovery** Document undo cannot retract a request already received by a provider. Recover inserted content according to host rules.


<a id="aiTranslate"></a>
### Translate and polish

**Purpose** Translate and polish.

**Steps**

1. Configure and test your own AI service
2. Choose language or polishing mode, read the selection, request processing, and verify names and terminology.
3. Verify facts, numbers, and masking codes before insertion

**Options** Configure provider, endpoint, model, Key, and protection mode in AI settings; choose text-processing modes in the panel.

**Limits** Sent content goes to your chosen provider and may incur charges. Masking is not a guarantee of anonymity, and AI output is not fact verification. Local tool suggestions remain useful without AI.

**Undo and recovery** Document undo cannot retract a request already received by a provider. Recover inserted content according to host rules.


<a id="ai-formula"></a>
### AI formula and outline assistance

**Purpose** Ask configured AI for formula explanations or a presentation outline for your review.

**Steps**

1. Review AI settings and outgoing content
2. Ask about formulas in chat, or enter topic, audience, and section count in the studio
3. Check references, function support, and facts
4. Preview before writing formulas or generating a deck

**Options** Outline section counts are 4, 6, 8, or 10. Ordinary layout generation does not require AI.

**Limits** Unlike the local formula assistant, AI output is not guaranteed correct or executable. Review generated macros and complex formulas before use.

**Undo and recovery** The original remains; inspect and delete unwanted output files.


<a id="office"></a>
## 15 Microsoft Office differences

### Entry points and runtime

Office uses an Office.js sidebar and local HTTPS, requiring certificate trust and add-in registration. Closing the sidebar stops MCP access through that page. WPS native APIs differ from Office.js; a WPS feature count is not an Office compatibility guarantee.

| Feature class | Office behaviour |
|---|---|
| Document spreadsheet and presentation edits | Uses public Office.js capabilities and provides guidance when unsupported |
| Files PDF and templates | The local service can process supported saved files |
| ID validation units QR and attendance | Entries with Office implementations execute directly |
| Macro reading native printing some layouts | WPS-only APIs are unavailable; guidance may replace automation |
| Masked copy and paste | Office APIs read selections and write text, subject to permissions and host limits |

Test a small sample in your installed Word, Excel, or PowerPoint first. Web-code tests do not establish desktop compatibility. Microsoft Office was unavailable in this round and was not tested on a desktop.

<a id="privacy"></a>
## 16 Privacy and local data

### Data flow

| Data or action | Location | What to check |
|---|---|---|
| Ordinary formatting and spreadsheet cleanup | Host and local computer | Retain the original and inspect the target scope |
| Cards paste bar settings and outputs | generated and configured output or backup locations | Local storage does not mean every file is encrypted |
| Masking maps and word lists | generated/masking.json | AES-256-GCM with a key protected by Windows DPAPI or Mac Keychain |
| MCP maps | Service memory | Cleared after two idle hours or a service restart |
| Built-in AI requests | Your chosen model provider | Inspect protection mode and actual input before sending |
| Web AI | The website where you paste | The toolbox cannot control the site's own collection or history |

DPAPI protects data for the current Windows user. Copying masking.json to another computer does not guarantee decryption. Mac requires Keychain access. If the system key is unavailable, retain the original files while troubleshooting; do not delete old maps as a repair step. Before sharing, inspect body text, headers, footers, comments, and accompanying materials. Automatic detection cannot guarantee that every piece of client information is found.

AI protection hides supported number patterns; register client names and other custom terms in the word list. Company plugins also participate in AI sending, and plugin-load failure disables AI. Running a local toolbox does not mean content already sent to an AI provider remains local.

<a id="limits"></a>
## 17 Backups undo and processing limits

### Recovery paths

| Situation | Recovery |
|---|---|
| Ordinary WPS Writer edits | Try Ctrl+Z immediately and inspect; use a retained copy if host undo is incomplete |
| WPS spreadsheet tool writes | Restore the saved backup; do not rely on Ctrl+Z |
| Whole-deck changes | Work in a saved copy; some operations have no undo |
| New PDF document or image | Keep the source and delete an unwanted output after checking it |
| Formula assistant writes | Use recent-write rollback, up to 10 records; only unchanged recorded formulas are removed |
| File renaming | Use the last-rename recovery feature while current files still match its record |
| Masked copies | Restore with the map; asterisks or black blocks cannot reconstruct missing originals |

Automatic backup does not replace saving. A disk backup may predate unsaved edits. Apply multiple operations one at a time and check each result; do not change documents or selections during processing.

### Common limits

Shared body-text selection is limited to 50,000 characters. Basic cell ranges are limited to 1,000 rows and 50 columns, with smaller limits in some tools. Quick table input allows 20 rows, 10 columns, and 10,000 characters. Masked copy allows 200,000 characters and 20,000 cells. MCP reads may be truncated: inspect the truncated flag. Split large inputs and do not treat partial output as a complete result.

Fonts used in a document must be available locally. Other hosts or operating systems can substitute fonts and repaginate. Before printing, check paper, orientation, scale, and page numbers. UI language switching changes buttons and messages, not the language of document content.

<a id="faq"></a>
## 18 Troubleshooting

### Symptoms causes and fixes

| Symptom | Common cause | Action |
|---|---|---|
| Toolbox tab missing | Registration or restart incomplete | Rerun installation, save, and reopen all host windows |
| Local service unavailable | Service stopped or port occupied | Launch Office Toolbox and inspect service-error.log |
| Version mismatch after update | Cached scripts | Save, fully exit WPS, and reopen |
| Empty or oversized selection | Wrong focus or an entire column selected | Select a bounded range and keep the same document active |
| Right-hand output refused | Existing cells protected | Insert blank output columns and retry |
| ID loses final digits | Previous numeric precision loss | Re-enter from the original source as text |
| Clipboard cannot be read | Host permission limits | Paste into the panel after the prompt, then restore and insert |
| Codes do not restore | Expired deleted or different-session map | Inspect unknown codes and reread or use the correct record |
| Macro reading fails | Missing VBA or project access | Read the guidance and use manual code paste |
| Japanese OCR unavailable | Missing system language component | Check OCR languages and test clear horizontal print |
| Postal update fails | Network or official file unavailable | Retain existing data and import an official ZIP if needed |
| PDF will not open | Incomplete file or protection | Check in its original reader; let an authorized person handle protection |
| Shortcut does not fire | Focus or key conflict | Focus the sidebar outside inputs, or use ribbon access keys |

### Restart and diagnostics

Save documents, close the host, and rerun the Office Toolbox launcher. Do not terminate every Node process: other applications may own them. If needed, inspect service.log and service-error.log in the installation directory. Include only relevant diagnostic lines and remove Keys, client names, and private path details.

Back up generated and independent documents before reinstalling. Register at the same location and reopen the host. Deleting the program folder alone does not remove all add-in registration; use the uninstaller for its own entries.

### Report through GitHub Issues

Use the project repository’s Issues page. Obtain its URL from the release or README; this manual does not guess an unconfigured repository address. Use the bug template with application/version, OS, minimal steps, expected/actual result, exact message, and a few relevant log lines. Provide a synthetic sample, never Keys or real client documents. The trilingual template is .github/ISSUE_TEMPLATE/bug_report.md.

<a id="appendix"></a>
## 19 Appendix

### Keyboard reference

| Key | Action and condition |
|---|---|
| Ctrl+C / Ctrl+V / Ctrl+S | Copy paste save; check focus first |
| Ctrl+Z | Host-supported undo only, not WPS spreadsheet tool writes |
| Alt then OT then K | WPS navigator, followed by your assigned access letter |
| Custom Ctrl+Alt combination | Opens a tool when the sidebar has focus outside an input; does not edit data |
| F9 or right-click update field | Update the whole contents table in the host |

### File locations

Installation paths vary. A common local example is %LOCALAPPDATA%/Programs/WpsContentWorkbench. Code and runtime live there; generated holds user data and outputs. Common files are generated/settings.json, generated/ai.json, and generated/masking.json. Cards and paste-bar items are managed locally. Backup location is configurable. Logs are service.log and service-error.log. Do not treat the code directory as your only document backup.

### Licences and versions

Read THIRD-PARTY-NOTICES.txt and WPS-SDK-LICENSE.txt in the installation, and respect separate asset licences. This manual uses the project’s existing docx and pdfkit packages without new npm dependencies.

0.16.0 includes proofing, data, postal, prose-style, and macro-checking tools. This continuation adds MCP masking, masked copy/restored paste, trilingual manuals, and Japanese formula-dialog completion. Check the installer and repository releases for the published version number. Sources are docs/manual/zh.md, ja.md, and en.md. Build with node scripts/docs/build-manual.mjs; output goes to release/docs.

### Native-operation guide cards

Purpose: find instructions for host-native operations. Steps: 1. Search for a card below; 2. Open and read its instructions; 3. Follow the native controls in a copy using your host version. Options belong to the native menu; one-click entries open the related tool. Limits: cards are not all automated, and menu names vary by version. Undo: opening a card changes no document; subsequent native operations follow host recovery rules. This index covers all 65 cards and links equivalent toolbox features to their detailed instructions.

| Guide card | Related tool or scope |
|---|---|
| Fit long text · 文字太长，不改列宽也能显示 | [Detailed equivalent tool](#tool-shrinkToFit) |
| Fill blank cells · 批量填充空白单元格 | [Detailed equivalent tool](#tool-fillBlanks) |
| Remove bracket notes · 去掉括号里的备注 | [Detailed equivalent tool](#tool-removeBrackets) |
| Number merged cells · 合并单元格大小不一也能编号 | [Detailed equivalent tool](#tool-mergedNumbering) |
| Flash Fill examples · 智能填充：加后缀、提取文字 | Native instructions in this card |
| Split a list with TEXTSPLIT · 一格里的名单拆成一列 | Native instructions in this card |
| Return all FILTER matches · 按工号查出全部匹配记录 | Native instructions in this card |
| Rank without sorting · 成绩排名，不打乱原表顺序 | [Detailed equivalent tool](#tool-rankKeepOrder) |
| Extract amounts carefully · 从描述文字里提金额再求和 | Native instructions in this card |
| Print one page wide · 超宽表格打印在一页宽 | [Detailed equivalent tool](#tool-fitOnePageWide) |
| Repeat print headers · 每页都打印表头 | [Detailed equivalent tool](#tool-printTitles) |
| Join text with line breaks · 几列内容合成一格并换行 | [Detailed equivalent tool](#tool-joinColumns) |
| Transpose a copy · 横表竖表互换 | [Detailed equivalent tool](#tool-transposeRange) |
| Lock formula references · 复制公式时固定范围 | Native instructions in this card |
| Consolidate by labels · 多块数据按标签合并计算 | Native instructions in this card |
| Centre without merging · 标题居中但不合并单元格 | [Detailed equivalent tool](#tool-centerAcross) |
| Copy a range as a picture · 表格区域变成一张图片 | [Detailed equivalent tool](#tool-rangeToImage) |
| Fit multiple pictures into cells · 批量插入图片并放进单元格 | [Detailed equivalent tool](#tool-fitPicturesToCells) |
| Compare two data lists · 快速核对两个表格的数据 | [Detailed equivalent tool](#tool-compareRanges) |
| Due-date conditional formatting · 合同到期自动提醒 | [Detailed equivalent tool](#tool-dueReminder) |
| Make payslips · 批量制作工资条 | [Detailed equivalent tool](#tool-payrollSlips) |
| Percentage data bars · 百分比进度条 | [Detailed equivalent tool](#tool-progressBar) |
| Borders for populated rows · 输入内容后自动加边框 | [Detailed equivalent tool](#tool-autoBorder) |
| Align names · 两字和三字姓名对齐 | [Detailed equivalent tool](#tool-alignNames) |
| Move columns · 快速互换两列位置 | [Detailed equivalent tool](#tool-swapColumns) |
| Stack columns with TOCOL · 多列姓名整理成一列 | [Detailed equivalent tool](#tool-stackColumns) |
| Change photo background · 证件照换底色 | Native instructions in this card |
| Create a QR code · 制作二维码 | [Detailed equivalent tool](#tool-codesBatch) |
| Heading-based contents · 用标题自动生成目录 | [Detailed equivalent tool](#tool-toc) |
| Unexpected page breaks · 表格或段落突然跳到下一页 | [Detailed equivalent tool](#tool-fixPageJump) |
| Move table rows · 表格整行上下移动 | Native instructions in this card |
| Landscape middle section · 只把中间一页改成横向 | [Detailed equivalent tool](#tool-landscapePages) |
| Keep picture tables narrow · 表格里插图不把列撑宽 | [Detailed equivalent tool](#tool-tablePictureFit) |
| Text to table · 文字按分隔符转成表格 | [Detailed equivalent tool](#tool-textTable) |
| Distribute rows and columns · 杂乱表格统一行高列宽 | Native instructions in this card |
| Line and character spacing · 调整行距和字间距 | Native instructions in this card |
| Superscript and subscript · 输入上标、下标 | [Detailed equivalent tool](#tool-superSubscript) |
| Picture positioning · 图片不乱跑 | [Detailed equivalent tool](#tool-pictureFix) |
| Bold inline headings · 批量加粗“1、标题：”这类行内标题 | [Detailed equivalent tool](#tool-boldInlineHeads) |
| Remove a watermark · 去掉文档水印 | [Detailed equivalent tool](#tool-removeWatermark) |
| Vertical text labels · 竖排标签文字 | Native instructions in this card |
| Fix a clipped picture · 插入图片只显示一条 | [Detailed equivalent tool](#tool-pictureFix) |
| Print card fronts and backs · 打印身份证正反面 | [Detailed equivalent tool](#tool-idCardSize) |
| Background images · 添加背景图片 | Native instructions in this card |
| Prepare document photos · 拍照的文件不要直接打印 | [Detailed equivalent tool](#tool-photoToPrint) |
| Remove blank pages · 删除空白页 | [Detailed equivalent tool](#tool-deleteBlankPages) |
| Align signature blocks · 落款、乙方信息靠右对齐 | [Detailed equivalent tool](#tool-alignSignature) |
| Align answer choices · 批量对齐 ABCD 选项 | [Detailed equivalent tool](#tool-alignOptions) |
| Aligned underlined blanks · 文字和下划线对齐（填空线） | Native instructions in this card |
| Diagonal header · 表格斜线表头 | [Detailed equivalent tool](#tool-diagonalHeader) |
| Merge documents · 把多个文档合成一个 | [Detailed equivalent tool](#tool-insertFiles) |
| Compare versions · 比对两份文件哪里不同 | [Detailed equivalent tool](#tool-compareDocs) |
| Transparent signature image · 手写签名转电子签名 | [Detailed equivalent tool](#tool-signatureClean) |
| Latin word wrapping · 数字、英文不自动换行 | [Detailed equivalent tool](#tool-noBreakLatin) |
| Official document layout · 公文排版标准格式 | [Detailed equivalent tool](#tool-officialFormatCn) |
| OCR and edit image text · 修改图片里的文字 | Native instructions in this card |
| Arrange a seam stamp · 用单位的电子章排骑缝章 | Native instructions in this card |
| Copy paste undo save · 复制粘贴撤销保存 | Native instructions in this card |
| Host-specific shortcuts · 同一快捷键在不同软件意思不同 | Native instructions in this card |
| Switch and cancel · 切换窗口、退出当前操作 | Native instructions in this card |
| Recover from backups · 找回没保存的文件 | [Detailed equivalent tool](#docHistory) |
| Explain one technique · 把一个技巧做成教程 | Native instructions in this card |
| Pre-delivery checks · 交付前的检查清单 | Native instructions in this card |
| Print labels at actual size · 按实物尺寸做文件夹侧标签 | Native instructions in this card |
| Move the backup location · 把 WPS 备份位置改到其他盘 | Native instructions in this card |
