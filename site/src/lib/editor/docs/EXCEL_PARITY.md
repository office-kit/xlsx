# Excel parity checklist

The authoritative list of user-facing Microsoft Excel operations the browser
editor has to cover. The goal is that a user can do any normal Excel task
(VBA/macros excluded) without opening Excel.

**Source.** Microsoft Excel for Mac **16.113.3**, inspected on 2026-10-03 through
macOS Accessibility (System Events/JXA plus a small `AXUIElement` dumper).
The raw dumps and screenshots are not kept in the repository: they reproduce
Microsoft's UI text and imagery, which is not ours to redistribute. Re-run the
dumper locally when a row needs re-checking.

AX does not expose ribbon **group names** (groups are anonymous `AXGroup`s).
The group names below come from Excel's visible labels and general knowledge.
Some per-option details (for example the option panes of each Number-format
category) are not exposed by AX either. Those rows are marked *(knowledge)*.

Status column legend: `todo` (not started), `partial`, `done`, `out-of-scope`.

`SOURCE WIRED / NOT RUN` marks a row whose whole path was traced in the source
(UI → action → model change inside `doc.transact` → library save/load) but
never executed on the current code. A UI label alone, a placeholder, an alert
or a disabled control does not qualify. The trace for each such row is listed
in `DONE_CRITERIA.md` under "Source traces".
Everything starts at `todo` except obvious out-of-scope items: VBA/macros,
add-ins, cloud/OneDrive/co-authoring/sharing, Copilot/Analyze Data, Power
Query/external data, 3D maps/3D models, Python in Excel, Office Scripts/Power
Automate, inking, and dictation/translation services.

---

## 1. Ribbon

### 1.1 Home

| Group | Operation | Excel exposure | Editor status |
| --- | --- | --- | --- |
| Clipboard | Paste (default) | Paste button, Cmd+V | todo |
| Clipboard | Paste dropdown: Paste / Keep Text Only / Use Text Import Wizard… / Paste Special… | Paste ▾ | todo |
| Clipboard | Cut | Cmd+X | todo |
| Clipboard | Copy | Cmd+C | todo |
| Clipboard | Copy as Picture… | Copy ▾ | todo |
| Clipboard | Format Painter (single use; double-click = sticky) | "Format" toggle | todo |
| Font | Font family | Font combo | SOURCE WIRED / NOT RUN |
| Font | Font size | Size combo | SOURCE WIRED / NOT RUN |
| Font | Increase / Decrease font size | A^ / A˅ buttons, Cmd+Shift+> / Cmd+Shift+< | SOURCE WIRED / NOT RUN |
| Font | Bold / Italic | Cmd+B / Cmd+I | SOURCE WIRED / NOT RUN |
| Font | Underline / Double underline | Underline ▾, Cmd+U | SOURCE WIRED / NOT RUN |
| Font | Borders preset menu: Bottom, Top, Left, Right, No Border, All, Outside, Thick Box, Bottom Double, Thick Bottom, Top and Bottom, Top and Thick Bottom, Top and Double Bottom | Borders ▾ | SOURCE WIRED / NOT RUN |
| Font | Draw Border / Draw Border Grid / Erase Border / Line Color / Line Style | Borders ▾ | todo |
| Font | More Borders… (opens Format Cells › Border) | Borders ▾ | todo |
| Font | Fill color (theme 10×6 + standard 10 + No Fill + More Colors…) | Shading ▾ | SOURCE WIRED / NOT RUN |
| Font | Font color (theme + standard + Automatic + More Colors…) | Font Color ▾ | SOURCE WIRED / NOT RUN |
| Font | Phonetic guide: Show Phonetic Fields / Edit Phonetic / Phonetic Settings… | abc▾ | todo |
| Alignment | Top / Middle / Bottom align | toggles | SOURCE WIRED / NOT RUN |
| Alignment | Left / Center / Right align | toggles (Cmd+L / Cmd+E; no right-align shortcut, Cmd+R is Fill Right) | SOURCE WIRED / NOT RUN |
| Alignment | Orientation: Angle Counterclockwise / Angle Clockwise / Vertical Text / Rotate Text Up / Rotate Text Down / Format Cell Alignment | Orientation ▾ | SOURCE WIRED / NOT RUN |
| Alignment | Decrease / Increase indent | buttons, Ctrl+Option+Tab / Ctrl+Option+Shift+Tab | SOURCE WIRED / NOT RUN |
| Alignment | Wrap Text / Shrink Text to Fit | Wrap Text ▾ | SOURCE WIRED / NOT RUN |
| Alignment | Merge & Center / Merge Across / Merge Cells / Unmerge Cells | Merge & Center ▾ | SOURCE WIRED / NOT RUN |
| Number | Number format combo (General, Number, Currency, Accounting, Short Date, Long Date, Time, Percentage, Fraction, Scientific, Text, More Number Formats…) | combo | todo |
| Number | Accounting format dropdown (¥ ja-JP, $ en-US, £ en-GB, € Euro, ¥ zh-CN, CHF fr-CH, More Accounting Formats…) | Accounting ▾ | todo |
| Number | Percent style | %, Cmd+Shift+% | SOURCE WIRED / NOT RUN |
| Number | Comma style | `,` button | SOURCE WIRED / NOT RUN |
| Number | Increase / Decrease decimal | buttons | SOURCE WIRED / NOT RUN |
| Styles | Conditional Formatting menu (see §4.6) | CF ▾ | SOURCE WIRED / NOT RUN |
| Styles | Format as Table gallery (Light / Medium / Dark), New Table Style…, New PivotTable Style… | Format as Table ▾ | todo |
| Styles | Cell Styles gallery (Good, Bad and Neutral; Data and Model; Titles and Headings; Themed Cell Styles; Number Format), New Cell Style…, Merge Styles… | Cell Styles gallery | todo |
| Cells | Insert: Insert Cells… / Insert Sheet Rows / Insert Sheet Columns / Insert Sheet | Insert ▾ | SOURCE WIRED / NOT RUN |
| Cells | Delete: Delete Cells… / Delete Sheet Rows / Delete Sheet Columns / Delete Table Rows / Delete Table Columns / Delete Sheet | Delete ▾ | todo |
| Cells | Format: Row Height… / AutoFit Row Height / Column Width… / AutoFit Column Width / Default Width… / Hide & Unhide (rows, columns, sheet) / Rename Sheet / Move or Copy Sheet… / Tab Color / Protect Sheet… / Lock Cell / Format Cells… | Format ▾ | todo |
| Editing | AutoSum: Sum / Average / Count Numbers / Max / Min / More Functions… | AutoSum ▾, Cmd+Shift+T | SOURCE WIRED / NOT RUN |
| Editing | Fill: Down / Right / Up / Left / Across Worksheets… / Series… / Justify / Flash Fill | Fill ▾ (same as Edit › Fill) | todo |
| Editing | Clear: All / Formats / Contents / Comments and Notes / Hyperlinks / Remove Hyperlinks | Clear ▾ | SOURCE WIRED / NOT RUN |
| Editing | Sort & Filter: Sort A to Z / Sort Z to A / Custom Sort… / Filter / Clear / Reapply | Sort & Filter ▾ | todo |
| Editing | Find & Select: Find… / Replace… / Go To… / Go To Special… / Formulas / Notes / Conditional Formatting / Constants / Data Validation / Select Objects / Selection Pane… | Find & Select ▾ | todo |
| Add-ins | Add-ins | button | out-of-scope |

### 1.2 Insert

| Group | Operation | Excel exposure | Editor status |
| --- | --- | --- | --- |
| Tables | PivotTable | button | SOURCE WIRED / NOT RUN |
| Tables | Recommended PivotTables | button | todo |
| Tables | Table (Create Table dialog, "My table has headers") | button, Cmd+T / Ctrl+T | SOURCE WIRED / NOT RUN |
| Forms | New / Preview / Edit / Send Form (Microsoft Forms) | Forms ▾ | out-of-scope |
| From Picture | Picture From File… / From Clipboard (image → table OCR) | ▾ | out-of-scope |
| Illustrations | Pictures › Place in Cell / Place over Cells (Photo Browser, From File, Stock Images, Online Pictures) | Pictures ▾ | todo (From File only; stock/online out-of-scope) |
| Illustrations | Shapes gallery (Recently Used, Lines, Rectangles, Basic Shapes, Block Arrows, Equation Shapes, Flowchart, Stars and Banners, Callouts) | Shapes ▾ | todo |
| Illustrations | Icons | button | out-of-scope |
| Illustrations | 3D Models | ▾ | out-of-scope |
| Illustrations | SmartArt (List, Process, Cycle, Hierarchy, Relationship, Matrix, Pyramid, Picture) | ▾ | out-of-scope |
| Illustrations | Screenshot / Screen Clipping | ▾ | out-of-scope |
| Controls | Checkbox (cell checkbox, boolean cell value) | button | todo |
| Charts | Recommended Charts | ▾ | SOURCE WIRED / NOT RUN |
| Charts | Column / Bar (2-D, 3-D) | ▾ | SOURCE WIRED / NOT RUN |
| Charts | Line / Area (2-D, 3-D) | ▾ | SOURCE WIRED / NOT RUN |
| Charts | Pie / Doughnut (2-D, 3-D) | ▾ | SOURCE WIRED / NOT RUN |
| Charts | Hierarchy: Treemap, Sunburst | ▾ | todo |
| Charts | Statistical: Histogram, Pareto, Box and Whisker | ▾ | todo |
| Charts | X Y (Scatter), Bubble | ▾ | SOURCE WIRED / NOT RUN |
| Charts | Waterfall, Funnel, Stock, Surface, Radar | ▾ | todo |
| Charts | Combo | ▾ | todo |
| Charts | Maps: Filled Map | ▾ | out-of-scope (needs Bing geodata) |
| Charts | PivotChart | button | todo |
| Sparklines | Line / Column / Win/Loss | Sparklines ▾ | SOURCE WIRED / NOT RUN |
| Filters | Slicer / Timeline | buttons | todo |
| Links | Link (Insert Hyperlink dialog) | button, Cmd+K | SOURCE WIRED / NOT RUN |
| Comments | New Comment (threaded) | button, context menu | done |
| Text | Text Box (horizontal / vertical) | ▾ | todo |
| Text | Header & Footer (switches to Page Layout view) | button | todo |
| Text | WordArt | ▾ | todo |
| Text | Object (embedded OLE) | button | out-of-scope |
| Symbols | Equation (gallery + Insert New Equation) | ▾ | out-of-scope |
| Symbols | Symbol | button | SOURCE WIRED / NOT RUN |

### 1.3 Draw

| Operation | Excel exposure | Editor status |
| --- | --- | --- |
| Draw mode, Eraser (stroke / small / medium / segment), Lasso Select, Pens gallery, Add Pen/Highlighter/Pencil, Draw with Trackpad | Draw tab | out-of-scope |

### 1.4 Page Layout

| Group | Operation | Excel exposure | Editor status |
| --- | --- | --- | --- |
| Themes | Themes (Office + Browse… + Save Current Theme…) | Themes ▾ | todo |
| Themes | Theme Colors / Theme Fonts | ▾ | todo |
| Page Setup | Margins presets (Normal 0.75/0.75/0.7/0.7", Wide 1", Narrow 0.75/0.25") + Custom Margins… | Margins ▾ | SOURCE WIRED / NOT RUN |
| Page Setup | Orientation Portrait / Landscape | ▾ | SOURCE WIRED / NOT RUN |
| Page Setup | Size (Letter, Tabloid, Legal, A3, A4, A5, B4, B5, …) | ▾ | SOURCE WIRED / NOT RUN |
| Page Setup | Print Area: Set / Clear | ▾ (also File › Print Area) | SOURCE WIRED / NOT RUN |
| Page Setup | Breaks: Insert / Remove / Reset All Page Breaks | ▾ | SOURCE WIRED / NOT RUN |
| Page Setup | Background (sheet background picture) | button | todo |
| Page Setup | Print Titles (Page Setup › Sheet) | button | SOURCE WIRED / NOT RUN |
| Page Setup | Page Setup dialog | button | SOURCE WIRED / NOT RUN |
| Scale to Fit | Width / Height (Automatic, 1 page, …) | combos | SOURCE WIRED / NOT RUN |
| Sheet Options | Gridlines View / Print | checkboxes | todo |
| Sheet Options | Headings View / Print | checkboxes | todo |

### 1.5 Formulas

| Group | Operation | Excel exposure | Editor status |
| --- | --- | --- | --- |
| Function Library | Insert Function (Formula Builder pane) | button, Shift+F3 | todo |
| Function Library | AutoSum ▾ (Sum, Average, Count Numbers, Max, Min, More Functions…) | ▾, Cmd+Shift+T | todo |
| Function Library | Recently Used / Financial / Logical / Text / Date & Time / Lookup & Reference / Math & Trig function menus | ▾ (full lists in raw dump) | todo |
| Function Library | More Functions › Statistical / Engineering / Cube / Information / Compatibility / Web | ▾ | todo |
| Python | Insert Python / Reset / Editor / Initialization | buttons | out-of-scope |
| Defined Names | Name Manager | button, Ctrl+F3 | SOURCE WIRED / NOT RUN |
| Defined Names | Define Name… / Apply Names… | ▾ | todo |
| Defined Names | Use in Formula (list of names + Paste Names…) | ▾ | todo |
| Defined Names | Create from Selection | button, Cmd+Shift+F3 | todo |
| Formula Auditing | Trace Precedents / Trace Dependents | buttons | todo |
| Formula Auditing | Remove Arrows / Remove Precedent Arrows / Remove Dependent Arrows | ▾ | todo |
| Formula Auditing | Show Formulas | toggle, Ctrl+` | todo |
| Formula Auditing | Error Checking… / Trace Error / Circular References | ▾ | todo |
| Formula Auditing | Watch Window | toggle | todo |
| Calculation | Calculation Options: Automatic / Partial (automatic except data tables) / Manual / Format Stale Values / Compatibility Version | ▾ | todo |
| Calculation | Calculate Now / Calculate Sheet | buttons, Cmd+= / Shift+F9 | todo |

### 1.6 Data

| Group | Operation | Excel exposure | Editor status |
| --- | --- | --- | --- |
| Get & Transform | Get Data (Power Query)…, Launch Power Query Editor…, Data Source Settings…, From Database (MS Query), From HTML, From Text (Legacy), From SQL Server ODBC | ▾ | out-of-scope (From Text/CSV import: todo) |
| Get & Transform | From Picture | ▾ | out-of-scope |
| Queries & Connections | Refresh All / Refresh / Refresh Status / Cancel Refresh / Properties | ▾ | out-of-scope |
| Queries & Connections | Queries & Connections pane, Properties, Workbook Links (Edit Links) | buttons | out-of-scope |
| Data Types | Stocks / Geography etc. gallery | gallery | out-of-scope |
| Sort & Filter | Sort A to Z / Sort Z to A | buttons | SOURCE WIRED / NOT RUN |
| Sort & Filter | Sort… (custom multi-level sort dialog) | button, Cmd+Shift+R | SOURCE WIRED / NOT RUN |
| Sort & Filter | Filter (AutoFilter toggle) | toggle, Cmd+Shift+F | SOURCE WIRED / NOT RUN |
| Sort & Filter | Clear / Reapply | buttons | SOURCE WIRED / NOT RUN |
| Sort & Filter | Advanced (Advanced Filter dialog) | button | SOURCE WIRED / NOT RUN |
| Data Tools | Text to Columns (wizard: Delimited / Fixed width) | button | SOURCE WIRED / NOT RUN |
| Data Tools | Analyze Data | button | out-of-scope |
| Data Tools | Flash Fill | button, Cmd+E | todo |
| Data Tools | Remove Duplicates | button | done (browser 2026-10-04; not checked in real Excel) |
| Data Tools | Data Validation… / Circle Invalid Data / Clear Validation Circles | ▾ | SOURCE WIRED / NOT RUN |
| Data Tools | Consolidate | button | SOURCE WIRED / NOT RUN |
| Forecast | What-If Analysis: Scenario Manager… / Goal Seek… / Data Table… | ▾ | SOURCE WIRED / NOT RUN |
| Outline | Group… / Auto Outline | ▾, Cmd+Shift+K | todo |
| Outline | Ungroup… / Clear Outline | ▾, Cmd+Shift+J | todo |
| Outline | Subtotal | button | SOURCE WIRED / NOT RUN |
| Outline | Show Detail / Hide Detail | buttons | SOURCE WIRED / NOT RUN |
| Analysis | Analysis Tools (Analysis ToolPak) | button | out-of-scope |

### 1.7 Review

| Group | Operation | Excel exposure | Editor status |
| --- | --- | --- | --- |
| Proofing | Spelling | button, F7 | todo |
| Proofing | Thesaurus | button | out-of-scope |
| Proofing | Workbook Statistics | button | todo |
| Accessibility | Check Accessibility / Alt Text / Accessibility Help | ▾ | todo (Alt Text), rest out-of-scope |
| Language | Translate | button | out-of-scope |
| Changes | Show Changes | toggle | out-of-scope (cloud) |
| Comments | New / Delete / Previous / Next Comment, Show Comments pane | buttons | done |
| Notes | New Note / Previous / Next / Show/Hide Note / Show All Notes / Convert to Comments | ▾, Shift+F2 | todo |
| Protect | Protect Sheet… | button | SOURCE WIRED / NOT RUN |
| Protect | Protect Workbook (structure) | toggle | SOURCE WIRED / NOT RUN |
| Protect | Always Open Read-Only | toggle | todo |
| Ink | Hide Ink / Delete All Ink | ▾ | out-of-scope |

### 1.8 View

| Group | Operation | Excel exposure | Editor status |
| --- | --- | --- | --- |
| Sheet View | Switch / Keep / Exit / New Sheet View, Options | controls | out-of-scope (co-authoring) |
| Workbook Views | Normal / Page Break Preview / Page Layout | toggles (also status bar) | todo |
| Workbook Views | Custom Views | button | todo |
| Show | Navigation pane | toggle | todo |
| Show | Ruler (Page Layout view) | checkbox | todo |
| Show | Gridlines | checkbox | todo |
| Show | Formula Bar | checkbox | todo |
| Show | Headings | checkbox | todo |
| Show | Data Type Icons | checkbox | out-of-scope |
| Show | Focus Cell (crosshair highlight of active row/col), Focus Cell Color, Show Auto-Highlight | ▾ | todo |
| Zoom | Zoom combo, Zoom to 100%, Zoom to Selection | controls | todo |
| Window | New Window / Arrange All / Hide / Unhide / View Side by Side / Synchronous Scrolling / Reset Window Position / Switch Windows | buttons | out-of-scope (single-window web app) |
| Window | Freeze Panes / Freeze Top Row / Freeze First Column (Unfreeze when frozen) | ▾ | SOURCE WIRED / NOT RUN |
| Window | Split | toggle | todo |
| Macros | View Macros / Record Macro / Use Relative References | buttons | out-of-scope |

### 1.9 Automate

| Operation | Excel exposure | Editor status |
| --- | --- | --- |
| Office Scripts (New Script, View Scripts, gallery), Flow Templates | Automate tab | out-of-scope |

### 1.10 Contextual tabs *(knowledge, appear only on selection)*

| Tab | Key operations | Editor status |
| --- | --- | --- |
| Table | Table Name, Resize Table, Summarize with PivotTable, Remove Duplicates, Convert to Range, Insert Slicer, Header Row / Total Row / Banded Rows / First Column / Last Column / Banded Columns / Filter Button, Table Styles gallery | todo |
| Chart Design | Add Chart Element (Axes, Axis Titles, Chart Title, Data Labels, Data Table, Error Bars, Gridlines, Legend, Lines, Trendline, Up/Down Bars), Quick Layout, Change Colors, Chart Styles, Switch Row/Column, Select Data, Change Chart Type, Move Chart | todo |
| Format (chart / shape) | Shape Fill / Outline / Effects, WordArt styles, Arrange (Bring Forward, Send Backward, Selection Pane, Align, Group, Rotate), Size | todo |
| Shape Format | Insert Shapes, Shape Styles, Text Fill/Outline/Effects, Alt Text, Arrange, Size | todo |
| Picture Format | Remove Background, Corrections, Color, Artistic Effects, Transparency, Compress, Change Picture, Reset, Picture Styles, Crop, Size | todo (Crop/Size/Reset), rest out-of-scope |
| PivotTable Analyze / Design | Field list, Refresh, Change Data Source, Group, Fields/Items/Sets, Subtotals, Grand Totals, Report Layout, Blank Rows, styles | todo |
| Sparkline | Edit Data, Type, Show (High/Low/Negative/First/Last/Markers), Style, Axis, Group/Ungroup, Clear | todo |
| Header & Footer | Page Number, Number of Pages, Date, Time, File Path, File Name, Sheet Name, Picture, Different First/Odd&Even | todo |

---

## 2. Menu bar (Mac only, from the menu dump)

Items that duplicate a ribbon command map to the same row above. Only the
shortcuts and the menu-only commands are listed here.

| Menu | Operation | Shortcut (as reported by AX) | Editor status |
| --- | --- | --- | --- |
| Excel | Preferences… | Cmd+, | todo (editor settings) |
| File | New / New from Template… | Cmd+N / Cmd+Shift+P | todo |
| File | Open… / Open Recent | Cmd+O | todo |
| File | Close / Save / Save As… / Save as Template… | Cmd+W / Cmd+S / Cmd+Shift+S | todo |
| File | Browse Version History, Share, Send Workbook/PDF | – | out-of-scope |
| File | Import (CSV/Text) | – | todo |
| File | Reduce File Size… | – | out-of-scope |
| File | Always Open Read-Only / Passwords… | – | todo (passwords: out-of-scope unless encryption added) |
| File | Page Setup… / Print Area › Set / Clear / Print… | – / Cmd+P | todo |
| File | Properties… (title, author, …) | – | todo |
| Edit | Undo / Redo (Repeat) | Cmd+Z / Cmd+Y (also Cmd+Shift+Z) | todo |
| Edit | Cut / Copy / Paste | Cmd+X / C / V | todo |
| Edit | Paste Special… | Ctrl+Cmd+V | todo |
| Edit | Paste and Match Formatting | Opt+Shift+Cmd+V | todo |
| Edit | Fill › Down / Right / Up / Left / Across Worksheets… / Series… / Justify / Flash Fill | Cmd+D / Cmd+R (shown in Fill submenu only when enabled) | todo |
| Edit | Clear › All / Formats / Contents / Comments and Notes / Hyperlinks / Series | – | todo |
| Edit | Select All | Cmd+A | todo |
| Edit | Delete… (cells dialog) | Ctrl+- / Cmd+- (verified) | todo |
| Edit | Sheet › Delete Sheet / Move or Copy Sheet… | – | todo |
| Edit | Find… / Find Next / Find Previous / Replace… / Go To… | Ctrl+F (Cmd+F) / Cmd+G / Cmd+Shift+G / Ctrl+H / Ctrl+G | todo |
| Edit | Toggle Drawing | Ctrl+Cmd+Z | out-of-scope |
| View | Normal / Page Layout | – | todo |
| View | Ribbon (show/hide) | Opt+Cmd+R | todo |
| View | Formula Bar / Message Bar | – | todo |
| View | Formula Builder | – | todo |
| View | Header and Footer… / Ruler / Custom Views… / Zoom… / Full Screen | – | todo (Full Screen out-of-scope) |
| Insert | Cells… / Rows / Columns | Ctrl+Shift+= (Cmd+Shift+=) | todo |
| Insert | Sheet › Insert Sheet / Chart Sheet | Shift+F11 | todo |
| Insert | Chart › (17 chart families, Manage Templates…) | – | todo |
| Insert | Sparklines… / Table | – | todo |
| Insert | Page Break / Reset All Page Breaks | – | todo |
| Insert | Function… | Shift+F3 | todo |
| Insert | Name › Define Name / Paste… / Create… / Apply… / Name Manager | Ctrl+F3 etc. | todo |
| Insert | New Comment | – | done |
| Insert | Picture / Audio / Movie / Symbol… / Shape / Icons / 3D Models / Text Box / SmartArt / WordArt / Object… | – | see Insert tab (Audio/Movie out-of-scope) |
| Insert | Hyperlink… | Cmd+K | todo |
| Insert | Checkbox | – | todo |
| Format | Cells… | Cmd+1 | todo |
| Format | Row › Height… / AutoFit / Hide / Unhide | Ctrl+9 / Ctrl+Shift+9 | todo |
| Format | Column › Width… / AutoFit Selection / Hide / Unhide / Standard Width… | Ctrl+0 / Ctrl+Shift+0 | todo |
| Format | Sheet › Rename / Hide / Unhide… / Background… | – | todo |
| Format | Conditional Formatting… (Manage Rules) | – | todo |
| Format | Style… (cell style dialog) | – | todo |
| Format | Phonetic Guide › Edit / Settings… / Show | – | todo |
| Format | Selection Pane… | Opt+Cmd+U | todo |
| Tools | Spelling… / Thesaurus… / Language… / AutoCorrect Options… / Error Checking… | Ctrl+Opt+Cmd+R (Thesaurus) | todo (Thesaurus/Language out-of-scope) |
| Tools | Translate… / Check Accessibility | Ctrl+Opt+Cmd+T | out-of-scope |
| Tools | Track Changes › Highlight / Accept or Reject, Merge Workbooks… | – | out-of-scope (legacy shared workbooks) |
| Tools | Protection › Protect Sheet… / Protect Workbook… | – | todo |
| Tools | Goal Seek… / Scenarios… | – | todo |
| Tools | Auditing › Trace Precedents / Dependents / Error / Remove All Arrows | – | todo |
| Tools | Macro › Macros… / Record New Macro… / Visual Basic Editor | – | out-of-scope |
| Tools | Excel Add-ins… / Customize Keyboard… | – | out-of-scope |
| Data | Sort… | Cmd+Shift+R | todo |
| Data | AutoFilter / Clear Filters / Advanced Filter… | Cmd+Shift+F | todo |
| Data | Subtotals… / Validation… / Table… (what-if data table) / Text to Columns… / Consolidate… | – | todo |
| Data | Group and Outline › Hide Detail / Show Detail / Group… / Ungroup… / Auto Outline / Clear Outline / Settings… | – | todo |
| Data | Edit Links… / Refresh / Get Data (Power Query) | – | out-of-scope |
| Data | Summarize with PivotTable / Chart Source Data… / Chart Add Data… | – | todo |
| Data | Table Tools › Remove Duplicates / Summarize with PivotTable / Rename / Convert to Range | – | todo |
| Window | Minimize / Zoom / New Window / Arrange… / Hide / Unhide / Bring All to Front / tiling | Cmd+M | out-of-scope |
| Window | Split / Freeze Panes | – | todo |
| Help | Excel Help / Feedback / Check for Updates | – | out-of-scope |

---

## 3. Context menus (right click)

| Target | Items (in order) | Editor status |
| --- | --- | --- |
| Cell | Cut, Copy, Paste, Paste Special ▸ [Paste, Formulas, Formulas & Number Formatting, Keep Source Formatting, No Borders, Keep Source Column Widths, Transpose, Values, Values & Number Formatting, Values & Source Formatting, Formatting, Paste Link, Paste Picture, Paste Picture Link, Paste Special…] · Thesaurus… · Insert… (Insert Copied Cells… when clipboard holds cells), Delete…, Clear Contents · Filter ▸ [Clear Filter, Reapply, Advanced Filter…, Filter by Selected Cell's Value / Color / Font Color / Icon] · Sort ▸ [Sort Smallest to Largest, Sort Largest to Smallest, Put Selected Cell Color / Font Color / Formatting Icon On Top, Custom Sort…] · New Comment, New Note · Format Cells… (Cmd+1), Pick From Drop-down List…, Show Phonetic Field, Define Name… · Hyperlink… (Cmd+K) | todo |
| Row header | Cut, Copy, Paste, Paste Special ▸ (same), Insert, Delete, Clear Contents, Format Cells… (Cmd+1), Row Height…, Hide (Ctrl+9), Unhide (Ctrl+Shift+9) | todo |
| Column header | Cut, Copy, Paste, Paste Special ▸ (same), Insert, Delete, Clear Contents, Format Cells…, Column Width…, Hide (Ctrl+0), Unhide (Ctrl+Shift+0) | todo |
| Sheet tab | Insert Sheet, Delete, Rename, Move or Copy…, View Code (out-of-scope), Protect Sheet…, Tab Color ▸ (color picker + More Colors…), Hide, Unhide…, Select All Sheets | todo |
| Table cell *(knowledge)* | adds Table ▸ [Totals Row, Convert to Range, Alt Text], Insert ▸ [Table Rows Above, Table Columns to the Left], Delete ▸ [Table Rows, Table Columns] | todo |
| Chart / shape *(knowledge)* | Cut/Copy/Paste, Edit Text, Group, Bring to Front, Send to Back, Format Shape/Chart Area…, Select Data…, Change Chart Type…, Save as Picture…, Alt Text | todo |

---

## 4. Dialogs

### 4.1 Format Cells (Cmd+1)

| Tab | Fields | Editor status |
| --- | --- | --- |
| Number | Category list: General, Number, Currency, Accounting, Date, Time, Percentage, Fraction, Scientific, Text, Special, Custom; Sample preview; per-category options *(knowledge)*: Number = decimal places, Use 1000 separator, negative-number style list; Currency = decimal places, symbol, negative style; Accounting = decimal places, symbol; Date/Time = type list + locale; Percentage/Scientific = decimal places; Fraction = type list (up to one/two/three digits, halves…hundredths); Special = Zip, Zip+4, Phone, SSN (+ locale); Custom = type field + list of existing formats, Delete | todo |
| Alignment | Horizontal (General, Left (Indent), Center, Right (Indent), Fill, Justify, Center Across Selection, Distributed (Indent)); Vertical (Top, Center, Bottom, Justify, Distributed); Indent; Justify distributed; Orientation dial + Degrees (−90…90) + vertical "Text"; Text control: Wrap text, Shrink to fit, Merge cells; *(knowledge)* Right-to-left text direction | todo |
| Font | Font list (theme Headings/Body first), Font style (Regular, Italic, Bold, Bold Italic), Size list (6–72), Underline (None, Single, Double, Single Accounting, Double Accounting), Color, Normal font checkbox, Effects: Strikethrough / Superscript / Subscript, Preview | todo |
| Border | Presets: None / Outline / Inside; border buttons: Top, Bottom, Left, Right, Inside Vertical, Inside Horizontal, Diagonal Up, Diagonal Down; Line Style: None, Hair, Dotted, Dash Dot Dot, Dash Dot, Dashed, Thin, Medium Dash Dot Dot, Slanted Dash Dot, Medium Dash Dot, Medium Dashed, Medium, Thick, Double; Line Color | todo |
| Fill | Background color, Pattern color, Pattern style, Sample *(knowledge: Fill Effects… gradient on Windows only)* | todo |
| Protection | Locked, Hidden (effective only when sheet is protected) | todo |

### 4.2 Other dialogs

| Dialog | Fields / options | Editor status |
| --- | --- | --- |
| Sort | Levels table (Column / Sort On: Values, Cell Color, Font Color, Conditional Formatting Icon / Order: A to Z, Z to A, Custom List… / Color/Icon), + add level, − remove level, Copy level, ▲▼ reorder, "My list has headers", Options… (Case sensitive; Orientation top-to-bottom / left-to-right) | todo |
| Data Validation › Settings | Allow: Any value, Whole number, Decimal, List, Date, Time, Text length, Custom; Data: between, not between, equal to, not equal to, greater than, less than, greater than or equal to, less than or equal to; Minimum/Maximum/Source/Formula; Ignore blank; In-cell dropdown (List); Apply these changes to all other cells with the same settings; Clear All | todo |
| Data Validation › Input Message | Show input message when cell is selected, Title, Input message | todo |
| Data Validation › Error Alert | Show error alert after invalid data is entered, Style: Stop / Warning / Information, Title, Error message | todo |
| Data Validation › IME Mode | No Control, On, Off (English mode), Disable, Hiragana, Full-width Katakana, Half-width Katakana, Full-width Alpha-Numeric, Half-width Alpha-Numeric | todo |
| Find & Replace | Tabs Find / Replace; Find what (+ recent ▾), Replace with; Options: Within (Sheet / Workbook), Search (By Rows / By Columns), Look in (Formulas, Values, Notes, Comments; Replace allows Formulas only), Match case, Match byte, Find entire cells only, Format… *(knowledge)*; buttons Find All (result list), Previous, Next, Replace, Replace All, Close; status "N cell(s) found" | todo |
| Go To (Ctrl+G / F5) | Go to list (defined names + recent references), Reference field, Special… | todo |
| Go To Special | Notes, Constants, Formulas (sub-options Numbers / Text / Logicals / Errors), Blanks, Current region, Current array, Row differences, Column differences, Precedents, Dependents (Direct only / All levels), Last cell, Visible cells only, Objects, Conditional formats, Data validation (All / Same) | todo |
| Insert (cells) | Shift cells right, Shift cells down, Entire row, Entire column | todo |
| Delete (cells) | Shift cells left, Shift cells up, Entire row, Entire column | todo |
| Paste Special | Paste: All, All except borders, All using Source theme, Column widths, Formula and number formats, Formulas, Values and number formats, Formats, Values, Comments and Notes, Validation, All, merge conditional formats; Operation: None, Add, Subtract, Multiply, Divide; Skip Blanks; Transpose; Paste Link | todo |
| Page Setup › Page | Orientation Portrait/Landscape; Scaling Adjust to N % / Fit to W pages wide by H tall; Paper size; Print quality; First page number (Auto) | todo |
| Page Setup › Margins | Top, Bottom, Left, Right, Header, Footer; Center on page Horizontally / Vertically | todo |
| Page Setup › Header/Footer | Header / Footer presets, Custom Header… / Custom Footer… (left/center/right sections with codes &P &N &D &T &Z &F &A), Different odd and even pages, Different first page, Scale with document, Align with page margins | todo |
| Page Setup › Sheet | Print area, Rows to repeat at top, Columns to repeat at left, Gridlines, Row and column headings, Black and white, Draft quality, Comments (none / at end / as displayed), Cell errors as, Page order Down-then-over / Over-then-down | todo |
| Manage Rules (Conditional Formatting) | Show rules for (Current Selection / This Sheet / each table / pivot), rule list (Rule, Format, Applies to, Stop if true), + / − / Edit Rule… / Duplicate Rule, reorder ▲▼ | todo |
| New Formatting Rule | Style: 2-Color Scale, 3-Color Scale, Data Bar, Icon Sets, Classic. Min/Mid/Max Type: Lowest/Highest Value, Number, Percent, Formula, Percentile. Classic types: Format only cells that contain; top or bottom ranked values (Top/Bottom, N, Percent); above or below average; unique or duplicate values; Use a formula. Format with: Light Red Fill with Dark Red Text, Yellow Fill with Dark Yellow Text, Green Fill with Dark Green Text, Light Red Fill, Red Text, Red Border, Custom Format… | todo |
| Create Table *(knowledge)* | Range, My table has headers | todo |
| Insert Hyperlink *(knowledge)* | Web Page or File / This Document (cell ref, defined names) / Email Address, Display, ScreenTip | todo |
| Name Manager / New Name *(knowledge)* | Name, Scope (Workbook / sheet), Comment, Refers to; list with filter; New / Edit / Delete | todo |
| Row Height / Column Width / Standard Width | single number (points / characters) | todo |
| Move or Copy Sheet *(knowledge)* | To book, Before sheet list, Create a copy | todo |
| Protect Sheet *(knowledge)* | Password, allow: Select locked / unlocked cells, Format cells / columns / rows, Insert columns / rows / hyperlinks, Delete columns / rows, Sort, Use AutoFilter, Use PivotTable, Edit objects, Edit scenarios | todo |
| Protect Workbook *(knowledge)* | Structure, Windows, Password | todo |
| Series *(knowledge)* | Series in Rows/Columns, Type Linear / Growth / Date / AutoFill, Date unit Day / Weekday / Month / Year, Trend, Step value, Stop value | todo |
| Text to Columns wizard *(knowledge)* | Delimited / Fixed width; delimiters Tab, Semicolon, Comma, Space, Other, Treat consecutive as one, Text qualifier; column data format General / Text / Date (MDY…) / Do not import; Destination | todo |
| Remove Duplicates *(knowledge)* | column checklist, My data has headers, Select All / Unselect All | todo |
| Subtotal *(knowledge)* | At each change in, Use function (Sum, Count, Average, Max, Min, Product, Count Numbers, StdDev, StdDevp, Var, Varp), Add subtotal to, Replace current subtotals, Page break between groups, Summary below data | todo |
| Advanced Filter *(knowledge)* | Filter in place / Copy to another location, List range, Criteria range, Copy to, Unique records only | todo |
| Goal Seek / Scenario Manager / Data Table *(knowledge)* | Set cell / To value / By changing cell; scenarios CRUD + Summary; Row input / Column input cell | todo |
| Consolidate *(knowledge)* | Function, Reference, All references, Use labels in Top row / Left column, Create links to source data | todo |
| AutoFilter dropdown *(knowledge)* | Sort Ascending / Descending / by Color, Filter by Color, Text/Number/Date Filters (Equals, Does Not Equal, Begins With, Ends With, Contains, Does Not Contain, Greater Than, …, Top 10, Above/Below Average, date periods), Search box, (Select All) value checklist, Clear Filter | todo |
| Format Painter, Insert Function / Formula Builder pane *(knowledge)* | search, category list, argument fields with live result | todo |
| Print *(knowledge)* | Printer, copies, pages, Print what: Active sheets / Entire workbook / Selection, Scale to fit, preview | todo (browser print/PDF) |

### 4.3 Conditional Formatting dropdown (Home)

| Submenu | Items *(knowledge for the submenu contents)* | Editor status |
| --- | --- | --- |
| Highlight Cells Rules | Greater Than…, Less Than…, Between…, Equal To…, Text that Contains…, A Date Occurring…, Duplicate Values…, More Rules… | todo |
| Top/Bottom Rules | Top 10 Items…, Top 10%…, Bottom 10 Items…, Bottom 10%…, Above Average…, Below Average…, More Rules… | todo |
| Data Bars | Gradient Fill (6), Solid Fill (6), More Rules… | todo |
| Color Scales | 12 presets, More Rules… | todo |
| Icon Sets | Directional, Shapes, Indicators, Ratings, More Rules… | todo |
| New Rule… / Clear Rules ▸ (Selected Cells / Entire Sheet / This Table / This PivotTable) / Manage Rules… | – | todo |

---

## 5. Mouse gestures *(knowledge, sample-verified where noted)*

| Gesture | Behavior | Editor status |
| --- | --- | --- |
| Click cell | Select cell, make it active | todo |
| Drag over cells | Extend selection rectangle from anchor | todo |
| Shift+click | Extend selection from active cell to clicked cell | todo |
| Cmd+click / Cmd+drag | Add another range (multi-area selection) | todo |
| Click row/column header | Select entire row/column; drag across headers selects several; Shift+click extends; Cmd+click adds | todo |
| Click corner (select-all) button | Select entire sheet | todo |
| Double-click cell | Enter Edit mode at the click position (caret at click) | todo |
| Double-click column header right edge | AutoFit column width (to widest content of selection) | todo |
| Double-click row header bottom edge | AutoFit row height | todo |
| Drag column/row header border | Resize; tooltip shows "Width: 8.43 (64 pixels)"-style value (Mac: width in characters + pixels) | todo |
| Drag fill handle (bottom-right square) | AutoFill: copy / series (1,2,3; Mon,Tue; Jan; dates; "Item 1") ; Auto Fill Options button afterwards (Copy Cells, Fill Series, Fill Formatting Only, Fill Without Formatting, Flash Fill) | todo |
| Option+drag fill handle *(Windows: Ctrl)* | Toggle copy vs series | todo |
| Double-click fill handle | Fill down to end of adjacent data column | todo |
| Right-drag fill handle | Context menu: Copy Cells, Fill Series, Fill Formatting Only, Fill Without Formatting, Fill Days/Weekdays/Months/Years, Linear Trend, Growth Trend, Series… | todo |
| Drag selection border | Move cells (Option+drag = copy, Shift+drag = insert-move) | todo |
| Click sheet tab / Shift+click / Cmd+click | Activate sheet / group-select sheets | todo |
| Double-click sheet tab | Rename inline | todo |
| Drag sheet tab | Reorder (Option+drag = copy) | todo |
| "+" next to tabs | Insert new sheet at end | todo |
| Tab scroll arrows | Scroll tab strip | todo |
| Drag horizontal-scrollbar splitter | Resize tab strip vs scrollbar | todo |
| Scroll wheel / trackpad | Vertical/horizontal scroll by rows/cols; Cmd? +scroll / pinch = zoom | todo |
| Hover cell with note / comment | Shows note (red triangle) or comment (purple corner) popup | done |
| Hover hyperlink / click hyperlink | Opens link (single click) ; click-and-hold selects cell | todo |
| Filter dropdown button click | Opens AutoFilter menu | todo |
| Click in formula bar | Enter Edit mode in formula bar | todo |
| Click while typing a formula (Point mode) | Inserts reference; drag inserts range; Cmd+click adds `,`-separated ref | todo |
| Drag formula-reference color frame | Moves/resizes the referenced range in the formula (range finder) | todo |
| Status bar right-click | Choose aggregates (Average, Count, Numerical Count, Min, Max, Sum) | todo |
| Zoom slider / +/− / percentage | Zoom 10–400% | todo |
| Name Box | Type a ref or name + Enter to go/select; dropdown lists defined names; type a new name to define it | todo |

---

## 6. Keyboard shortcuts (Mac Excel defaults)

"Menu" = the shortcut appears in the menu dump; "Verified" = sent
through System Events and the resulting selection/value checked;
"Knowledge" = documented Mac Excel default not exercised here.

### 6.1 Navigation and selection

| Keys | Action | Source | Editor status |
| --- | --- | --- | --- |
| Arrow | Move one cell | Verified | todo |
| Shift+Arrow | Extend selection by one cell (active cell stays) | Verified | todo |
| Cmd+Arrow | Jump to the edge of the data region. From a filled cell: last filled cell before a blank. From the last filled cell: next filled cell or the sheet edge (A6 → A1048576). From a blank cell: next filled cell (H20 + Cmd+Up → H1 when the column is empty) | Verified | todo |
| Cmd+Shift+Arrow | Extend the selection to the data edge (A1 → A1:C1 → A1:C6) | Verified | todo |
| Home (fn+Left) | Column A of the current row | Verified | todo |
| Ctrl+Home (fn+Ctrl+Left) | A1, or the top-left unfrozen cell when panes are frozen | Verified | todo |
| Ctrl+End (fn+Ctrl+Right) | Last used cell (bottom-right of used range) | Verified | todo |
| Cmd+Home / Cmd+End | **No effect** on this machine (the Windows Ctrl+Home equivalent is Ctrl+Home on Mac) | Verified | todo |
| Page Down / Page Up (fn+Down/Up) | One screen down/up; the active row shifts by the visible row count (A1 → A37 with 36 visible rows) | Verified | todo |
| Option+Page Down / Option+Page Up | One screen right/left (A37 → V37 with 21 visible columns) | Verified | todo |
| Option+Right / Option+Left (also Cmd+Page Down/Up) | Next / previous sheet | Knowledge | todo |
| Shift+Space | Select entire row(s) of the selection | Verified | todo |
| Ctrl+Space | Select entire column(s). **Swallowed by macOS** on this machine because it is the input-source toggle (JP setup) | Verified (no-op) | todo |
| Cmd+A | First press: current region (A1:C6). Second press: whole sheet (`$1:$1048576`). In a formula right after a function name: Formula Builder | Verified | todo |
| Cmd+Shift+Space | Same as Cmd+A (current region first) | Verified | todo |
| Ctrl+G / F5 | Go To dialog | Menu | todo |
| Ctrl+F (Cmd+F) | Find | Menu | todo |
| Ctrl+H | Replace | Menu | todo |
| Cmd+G / Cmd+Shift+G | Find next / previous | Menu | todo |
| Enter / Shift+Enter | Move down / up (direction set in Preferences › Edit; default down) | Verified | todo |
| Tab / Shift+Tab | Move right / left | Verified | todo |
| Enter inside a multi-cell selection | Moves the active cell **within** the selection, column-major, wrapping: B2→B3→C2→C3→B2. The selection is kept | Verified | todo |
| Tab inside a multi-cell selection | Row-major, wrapping: B3→C3→B2→C2→B3 | Verified | todo |
| Shift+Enter / Shift+Tab inside a selection | Reverse order, wrapping (B2 + Shift+Tab → C3) | Verified | todo |
| Tab … Tab, Enter | After entering across with Tab, Enter returns to the column where the Tab run started, one row down (L1 a ⇥ M1 b ⇥ N1 c ↵ → L2) | Verified | todo |
| Ctrl+. | Move the active cell to the next corner of the selection, clockwise (B2 → C2 → C3 …) | Verified | todo |
| Shift+Delete (Shift+Backspace) | Collapse the selection to the active cell | Verified | todo |
| F8 / Shift+F8 | Extend-selection mode / Add-to-selection mode | Knowledge | todo |
| Ctrl+Delete | Scroll the active cell into view | Knowledge | todo |

### 6.2 Editing

| Keys | Action | Source | Editor status |
| --- | --- | --- | --- |
| Typing a character | Replaces the content, **Enter** mode | Verified | todo |
| F2 / Ctrl+U | Edit mode on the active cell, caret at the end. Pressing it again while editing toggles Edit ↔ Enter mode | Verified | todo |
| Double-click | Edit mode with the caret at the click position | Knowledge | todo |
| Enter (while editing) | Commit, then move in the Enter direction | Verified | todo |
| Tab / Shift+Tab (while editing) | Commit, then move right / left | Verified | todo |
| Cmd+Enter (while editing) | Commit and **stay** on the cell | Verified | todo |
| Esc | Cancel the edit and restore the previous value | Verified | todo |
| Arrow (Enter mode) | Commit and move (typing `abc` then Right in E2 → value committed, F2 active) | Verified | todo |
| Arrow (Edit mode) | Move the caret (`xyz`, F2, Left, `Q` → `xyQz`) | Verified | todo |
| Arrow after `=` or an operator (Point mode) | Status bar shows **Point**. Arrow inserts a reference, then moves it; Shift+Arrow extends it to a range; typing an operator fixes it, and the next arrow starts a new reference (E4: `=`, Up, Shift+Up, `+`, Left, Enter → `=E2:E3+D4`) | Verified | todo |
| Ctrl+Enter (Ctrl+Return) | Commit the same entry into every cell of the selection; relative refs adjust (`=B2` in G2:G4 → `=B2`,`=B3`,`=B4`); the selection is kept | Verified | todo |
| Option+Enter | Line break inside the cell; Wrap Text turns on automatically and the row auto-heightens | Verified | todo |
| Delete (backspace key) | **Clears every cell of the selection** and stays in Ready mode (Mac behavior; unlike Windows, it does not enter Edit mode) | Verified | todo |
| fn+Delete (forward delete) | Clears the contents of the selection | Verified | todo |
| Cmd+D | Fill Down from the top row of the selection; for a single cell, copies the cell above | Verified | todo |
| Cmd+R | Fill Right from the left column of the selection | Verified | todo |
| Ctrl+; | Insert today's date as text in short-date form (`10/3/26`), which Excel then parses (watch the locale: en_JP read it as 2010-03-26) | Verified | todo |
| Cmd+; | Insert the current **time** (`9:33 PM`), static | Verified | todo |
| Cmd+Shift+; | Insert time (alternative) | Knowledge | todo |
| Ctrl+' | Copy the formula from the cell above (no ref adjustment) | Knowledge | todo |
| Ctrl+Shift+" | Copy the value from the cell above | Knowledge | todo |
| Cmd+T or F4 (while editing a formula) | Cycle the reference under or before the caret: `A1` → `$A$1` → `A$1` → `$A1` → `A1` | Verified | todo |
| Cmd+Y / F4 (not editing) | Repeat the last action | Menu (Edit › Repeat) | todo |
| Cmd+Z / Cmd+Y | Undo / Redo. AppleScript-driven changes clear the undo stack, so test only with UI actions | Verified | todo |
| Shift+F3 | Insert Function / Formula Builder | Knowledge | todo |
| Cmd+Shift+T | AutoSum | Knowledge (synthetic event inconclusive) | todo |
| Option+Down | Pick-from-list: a dropdown under the cell listing the column's distinct text values, sorted (apple, banana, cherry, date, elder). On a validation-list cell it opens that list; on an AutoFilter header it opens the filter menu | Verified | todo |
| Ctrl+Shift+A | Insert argument placeholders after a function name | Knowledge | todo |
| Tab (formula AutoComplete open) | Accept the highlighted function and add `(` | Knowledge | todo |
| Cmd+K | Insert Hyperlink dialog | Verified | todo |
| Shift+F2 | New / edit **Note** (opens the note box in edit mode) | Verified | todo |
| Cmd+Shift+F2 | New threaded **Comment** on the active cell (reply box when the cell already has a thread) | Knowledge | done |
| Ctrl+- or Cmd+- | Delete dialog (deletes directly when whole rows/columns are selected) | Verified | todo |
| Ctrl+Shift+= (Cmd+Shift+=) | Insert dialog | Knowledge (JIS layout: not testable by key code) | todo |
| Ctrl+9 / Ctrl+Shift+9 | Hide / Unhide rows (glyphs shown in the row-header context menu) | Context menu | todo |
| Ctrl+0 / Ctrl+Shift+0 | Hide / Unhide columns | Knowledge | todo |
| Cmd+Shift+K / Cmd+Shift+J | Group / Ungroup | Knowledge | todo |
| Cmd+8 | Show/hide outline symbols | Knowledge | todo |
| Ctrl+` | Show formulas | Knowledge | todo |
| Cmd+= / Shift+F9 | Calculate all / calculate sheet | Knowledge | todo |
| Ctrl+Shift+U | Expand/collapse the formula bar | Knowledge | todo |
| Shift+F11 | Insert sheet | Menu | todo |
| Opt+Cmd+R | Show/hide ribbon | Menu | todo |
| Cmd+Ctrl+U | Search ("Tell me") | Toolbar label | out-of-scope |

### 6.3 Formatting

| Keys | Action | Source | Editor status |
| --- | --- | --- | --- |
| Cmd+1 | Format Cells | Verified | todo |
| Cmd+B / Cmd+I / Cmd+U | Bold / Italic / Underline | Knowledge | todo |
| Cmd+Shift+X | Strikethrough | Knowledge | todo |
| Cmd+Shift+> / Cmd+Shift+< | Increase / decrease font size | Knowledge | todo |
| Cmd+E / Cmd+L | Center / Left align (right align has no default because Cmd+R is Fill Right) | Knowledge | todo |
| Ctrl+Shift+~ | General number format | Knowledge | todo |
| Ctrl+Shift+$ | Currency (2 decimals) | Knowledge | todo |
| Ctrl+Shift+% | Percentage (0 decimals) | Knowledge | todo |
| Ctrl+Shift+^ | Scientific | Knowledge | todo |
| Ctrl+Shift+# | Date (d-mmm-yy) | Knowledge | todo |
| Ctrl+Shift+@ | Time (h:mm AM/PM) | Knowledge | todo |
| Ctrl+Shift+! | Number with 2 decimals and thousands separator | Knowledge | todo |
| Cmd+Opt+0 | Outline border | Knowledge | todo |
| Cmd+Opt+Arrow | Top/bottom/left/right border | Knowledge | todo |
| Cmd+Opt+- | Remove borders | Knowledge | todo |
| Cmd+Shift+F | AutoFilter on/off (Data › AutoFilter) | Verified | todo |
| Cmd+Shift+L | **No effect** here (it is the Windows filter shortcut) | Verified (no-op) | todo |
| Cmd+T / Ctrl+T (not editing) | Create Table | Knowledge (synthetic event opened nothing within 15 s) | todo |
| Opt+Cmd+U | Selection pane | Menu | todo |

### 6.4 File / workbook

| Keys | Action | Source | Editor status |
| --- | --- | --- | --- |
| Cmd+N / Cmd+Shift+P | New / New from Template | Menu | todo |
| Cmd+O / Cmd+S / Cmd+Shift+S | Open / Save / Save As | Menu | todo |
| Cmd+P | Print | Menu | todo |
| Cmd+W | Close | Menu | todo |
| Cmd+, | Preferences | Menu | todo |

---

## 7. Interaction semantics

Behavior a grid editor must reproduce exactly. Items marked **(verified)**
were exercised on Excel 16.113.3 through System Events key codes, with the
result read back through AppleScript (`address of selection`, cell values)
and the status bar mode text (`Ready` / `Enter` / `Edit` / `Point`).

### 7.1 Modes (status bar, left)

The status bar shows exactly one of `Ready`, `Enter`, `Edit`, `Point`. The
editor should model these modes explicitly. (Excel only shows a fifth state,
`text`, while a note/shape text box is being edited.)

| Mode | Entered by | Arrow keys | Enter / Tab | Esc |
| --- | --- | --- | --- | --- |
| Ready | default; after commit/cancel | move/extend the selection | move the active cell (within the selection if it has >1 cell) | clears copy marquee |
| Enter | typing a character into a cell (content is replaced) | **commit** the entry and move | commit + move | cancel, restore old value |
| Edit | F2, Ctrl+U, double-click, click in formula bar | move the caret inside the text (Home/End to line ends) | commit + move | cancel |
| Point | in Enter mode, when the text is a formula and the caret follows `=`, `(`, `,` or an operator, and an arrow key or mouse click is used | insert a reference to the cell next to the edited cell, then move that reference; Shift+Arrow extends it to a range | commit | cancel |

- F2 (or Ctrl+U) while editing toggles Enter ↔ Edit. In toggled Enter mode an
  arrow commits again (verified: Ctrl+U, Ctrl+U, Left → commit, D3 active).
- In Point mode, typing an operator (`+`, `,`, `)` …) fixes the current
  reference, and the next arrow starts a **new** reference from the edited
  cell's position (verified: `=` ↑ ⇧↑ `+` ← ↵ in E4 → `=E2:E3+D4`).
- Cmd+T or F4 in Enter/Edit mode cycles the absolute/relative form of the
  reference at the caret (verified `A1 → $A$1 → A$1 → $A1 → A1`).

### 7.2 Commit and move

- Enter commits and moves down. Shift+Enter moves up, Tab right, Shift+Tab left,
  Cmd+Enter stays put (all verified).
- **Within a multi-cell selection**, the movement keys never collapse the
  selection; they cycle the active cell inside it with wrap-around:
  Enter goes column-major (down, then the top of the next column, then back
  to the first cell), Tab goes row-major. The Shift variants reverse the
  order. With several areas (Cmd+click), it continues into the next area *(knowledge)*.
- Tab-then-Enter return: Enter after a run of Tab entries returns to the
  column where the run started (verified).
- Ctrl+Enter writes the entry into every selected cell, adjusting relative
  references as if filled. The selection and active cell stay (verified).
- Option+Enter inserts a line break and switches on Wrap Text. The row height then
  auto-fits (verified).
- Esc restores the old value (verified).

### 7.3 Navigation edge rules (Cmd+Arrow)

Given the direction, let `c` be the active cell and `n` the next cell:
1. If `c` and `n` are both non-empty: move to the last non-empty cell before
   the first empty one.
2. Otherwise: move to the first non-empty cell in that direction, or to the
   sheet edge (row 1 / 1048576, column A / XFD) if there is none.
Cmd+Shift+Arrow applies the same rule to the selection's moving edge, which is
anchored at the active cell. Verified: A1→A6→A1048576, H20 (empty column)→H1,
H1→C1.

### 7.4 Selection gestures

- Cmd+A: current region first (the contiguous non-empty block around the
  active cell, which counts diagonals), and whole sheet on a second press. On a blank
  isolated cell it selects the whole sheet directly *(knowledge)*.
- Shift+Space selects entire rows. Ctrl+Space (entire columns) is captured by
  macOS input switching on many JP setups, so the editor should also offer a
  non-conflicting path (clicking a column header).
- Ctrl+. walks the corners clockwise. Shift+Delete collapses to the active cell.
- After a fill/paste the selection is the destination range.

### 7.5 Clearing

- On Mac, **Delete (backspace) clears the whole selection and stays in Ready**
  (verified on B2:C2). fn+Delete (forward delete) does the same. Neither
  touches formats. To edit after clearing, the user starts typing.
- Edit › Clear › All / Formats / Contents / Comments and Notes / Hyperlinks
  are separate commands.

### 7.6 Fill

- Cmd+D fills the top row down through the selection, and Cmd+R fills the left column
  right. On a single cell they copy from the cell above / to the left
  (verified). Formulas adjust their relative references.
- Fill handle drag: series detection (numbers with step inferred from 2+
  cells, dates, weekday/month names, trailing numbers in text) versus copy for a
  single number. Option toggles. Double-click fills down to the extent of the
  adjacent column *(knowledge)*.

### 7.7 Insert date/time

- Ctrl+; inserts today's date, Cmd+; inserts the current time (both static,
  verified). They work both in Ready mode (starting an entry) and inside an edit.

### 7.8 AutoComplete and pick list

- Typing text that prefixes an existing value in the same column
  auto-completes inline (selected suffix). Enter accepts it, Delete removes the
  suggestion *(knowledge)*.
- Option+Down opens a dropdown of the column's distinct text values, sorted
  ascending (verified). The same gesture opens a validation list or AutoFilter
  menu on those cells.
- Formula AutoComplete: after `=` plus letters, a function list appears. Tab
  accepts the highlighted function and inserts `(`. Up/Down navigate the list
  *(knowledge)*.

### 7.9 Clipboard

- Cut/Copy shows a marching-ants marquee and keeps it until Esc, an edit, or
  another command. The status bar then reads "Select destination and press
  ENTER or choose Paste" (observed). Enter pastes once and clears the
  marquee, while Cmd+V can paste repeatedly. A cut-paste moves the cells and rewrites
  references that point at them *(knowledge)*.
- With a copied range on the clipboard, the context menu's "Insert…" becomes
  "Insert Copied Cells…" (observed).

### 7.10 Dialog conventions observed

- Excel dialogs are separate windows (Format Cells, Sort, Data Validation …).
  Range fields have a collapse-dialog button for picking a range on the sheet.
- Opening some dialogs is slow (10–20 s under AppleScript). That is not a
  UX target.
- A malformed file opens with the alert "We found a problem with some content
  in '…'. Do you want us to try to recover as much as we can?" (Yes/No).

---

## 8. Visual metrics

Measured on Excel 16.113.3 at 100 % zoom on a 2× (Retina) display; all values
are in **points**, which equal CSS px. The macOS appearance was **Dark**, so
the chrome colors below are dark-mode values. The cell area stays white in
dark mode. Light-mode chrome values are given as *(knowledge)*.

Geometry is matched; colour is not. Excel's green is its brand colour, so
wherever the tables below say "green" the editor uses its own accent
(`--xl-accent` in `ui/tokens.css`) so it is not mistaken for Excel.

### 8.1 Default font and cell size

| Item | Value | How |
| --- | --- | --- |
| Normal style font on this machine | **游ゴシック (Yu Gothic) "Body", 12 pt** (theme minor East-Asian font, because the Office locale is ja) | AppleScript `font object of style "Normal"` |
| Excel › Preferences standard font | **Aptos Narrow 12** (theme minor Latin font, the default for new workbooks in non-East-Asian locales) | AppleScript `standard font` |
| Default column width | `standardWidth` 10 chars = **75 pt** wide with Yu Gothic 12 (screenshot: 150 device px) | AppleScript + screenshot |
| Default row height | **20 pt** with Yu Gothic 12 (screenshot: 40 device px) | AppleScript + screenshot |
| Reference for Windows/Calibri 11 workbooks | 8.43 chars = 64 px column, 15 pt (20 px) row *(knowledge)*. The editor should derive width from the workbook's Normal-font max digit width, not hard-code it | – |
| Cell padding | Text starts about 3 pt from the left gridline; numbers end about 3 pt before the right gridline; the default vertical alignment is **center** for this Normal style (Format Cells shows Vertical = Center) | screenshot, Format Cells dump |
| Font list order | Theme fonts first: "游ゴシック Light (Headings)", "游ゴシック (Body)", then the alphabetical system list | Format Cells › Font |

### 8.2 Grid

| Element | Value |
| --- | --- |
| Gridline | 1 pt, **#DCDCDC** |
| Header ↔ cells separator | 1 pt, #C8C7C6 |
| Freeze-pane line | 1 pt, **#ABABAB** (darker than a gridline; no thick bar) |
| Column header height | ≈ 24 pt |
| Row header width | 25 pt for 1–2 digit row numbers (grows with digit count) |
| Header background | dark mode #1D1D1D; light mode ≈ #F3F3F3 *(knowledge)* |
| Header text | dark mode ≈ #B7B7B7–#FFFFFF, system font (SF), centered; row numbers centered |
| Header of a selected row/column | background **#696969** (dark mode; light ≈ #D2D2D2 *(knowledge)*), plus a **green bar** #4F825B on the edge next to the cells (bottom of column headers, right of row headers), ≈ 1–1.5 pt |
| Select-all corner | dark triangle in the top-left header cell |

### 8.3 Selection

| Element | Value |
| --- | --- |
| Selection border | **2 pt solid #4F825B** (Excel green) around the whole selection |
| Inner gap | 1 pt white between the border and the fill |
| Selection fill | **#D1D1D1** semi-opaque gray over the cells (text stays readable) |
| Active cell | not shaded (white) inside a multi-cell selection; for a single cell only the 2 pt border |
| Fill handle | **5 × 5 pt** solid #4F825B square at the bottom-right corner, with a 1 pt white ring, centered on the border corner |
| Copy marquee | animated dashed border in the same green *(observed during copy)* |

### 8.4 Formula bar (from AX frames)

| Element | Size / position |
| --- | --- |
| Name Box | 87 × 28 pt combo box at the left, with a ▲▼ stepper (dropdown of names) |
| Cancel ✕ / Enter ✓ buttons | 26 × 26 pt each; disabled (gray) in Ready mode, active while editing |
| fx (Formula Builder) | 40 × 26 pt, "fx ▾" |
| Separator | 3 pt vertical bar |
| Formula text area | single line of 19 pt in a row ≈ 33 pt high; expand ▼ button (26 × 21) at the far right (Ctrl+Shift+U) |
| Layout order (top to bottom) | title/QAT toolbar 40 pt (AutoSave, Home, Save, Save As, Undo ▾, Redo, …, title, Search) → ribbon 105 pt (tabs + commands) → message bar 31 pt (when shown) → formula bar ≈ 33 pt → grid → sheet-tab strip ≈ 25 pt → status bar 28 pt |

### 8.5 Sheet tabs and status bar

| Element | Value |
| --- | --- |
| Tab strip | background #414141 (dark mode); ◀ ▶ scroll buttons at the left, tabs, then a "+" new-sheet button; the horizontal scrollbar shares the row on the right |
| Active tab | **white background, green (#217346-ish, anti-aliased ~#63966F) text**, no bold, small rounded rectangle |
| Inactive tab | strip background, light text |
| Status bar | background #353535; left: mode (`Ready`/`Enter`/`Edit`/`Point`) or a contextual hint, then "Accessibility: Good to go"; right: **Average, Count, Sum** (only when ≥2 cells are selected and some are numeric; e.g. `Average: 3.75  Count: 6  Sum: 22.5`), then view buttons Normal / Page Layout / Page Break Preview, zoom − slider +, and the zoom percentage |
| Status bar aggregates | Average and Sum over numeric cells; Count = non-empty cells. Numerical Count / Min / Max can be enabled by right-click *(knowledge)* |
