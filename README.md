# ODS Spreadsheet Editor for VSCode

[Русская версия](README.ru.md)

Full-featured OpenDocument Spreadsheet (.ods) editor, built natively into VSCode. No external apps required.

## Quick Start

1. Open any `.ods` file in VSCode
2. Or right-click a folder in Explorer > **New Table (.ods)** to create from template

---

## Features

### Cell Editing
- **Direct input** -- click a cell and start typing, or press `F2` / `Enter` to edit
- **Formula bar** -- edit cell value or formula in the top bar
- **Formula support** -- `=SUM(A1:A10)`, `=IF(A1>0, "yes", "no")`, `=VLOOKUP(...)`, and 50+ functions
- **Cross-sheet references** -- `=Sheet2.A1` or `=Sheet2.A1:B10` (ODS dot-notation)
- **Auto-fill** -- drag the fill handle (small square at bottom-right of selection) to extend series: numbers, letters, IP addresses, repeating patterns
- **Clipboard** -- `Ctrl+C`, `Ctrl+V`, `Ctrl+X` with tab-separated format (compatible with Excel/Google Sheets paste)

### Selection & Navigation
- **Click** a cell to select it
- **Shift+click** or **Shift+Arrow** to extend selection
- **Click column header** (A, B, C...) to select entire column
- **Click row header** (1, 2, 3...) to select entire row
- **Ctrl+A** -- select all cells
- **Tab** / **Shift+Tab** -- move right/left
- **Enter** -- confirm edit and move down
- **Merged cells** -- clicking any part of a merged cell selects the entire merged area

### Formatting Toolbar
| Button | Action |
|--------|--------|
| **B** / **I** / **U** | Bold, Italic, Underline |
| Color pickers | Text color / Background color |
| Alignment icons | Left / Center / Right align |
| Borders dropdown | All borders, Outer, None, Top/Bottom/Left/Right |
| **123** (Number format) | General, Number (1,234.56), Currency ($1,234.56), Percent (12.34%), Date, Integer |
| Wrap text icon | Toggle text wrapping in cells |
| Freeze icon | Freeze/unfreeze rows and columns |
| **A↓** / **Z↓** | Sort ascending / descending |
| Filter icon | Open filter dropdown for selected column |
| Conditional format icon | Open conditional formatting dialog |

### Borders
Click the borders button in the toolbar to open a dropdown with options:
- **All Borders** -- apply borders to all cell edges in selection
- **Outer Borders** -- border only around the perimeter of selection
- **No Borders** -- remove all borders from selection
- **Single side** -- Top, Bottom, Left, or Right border only

### Number Formats
Click **123** in the toolbar:
- **General** -- no formatting (raw value)
- **Number** -- `1,234.56` (thousands separator, 2 decimal places)
- **Currency** -- `$1,234.56`
- **Percent** -- `12.34%` (value is multiplied by 100)
- **Date** -- `2024-01-15` format
- **Integer** -- `1,235` (thousands separator, no decimals)

### Text Wrap
Toggle wrap text to make cell content break into multiple lines within the cell. Combined with vertical alignment (top/middle/bottom), gives control over how long text is displayed.

### Freeze Panes
1. Select the cell at the intersection point -- everything **above** and **to the left** of it will be frozen
2. Click the freeze button (dashed cross icon)
3. Frozen rows/columns stay fixed while scrolling
4. Click again to unfreeze

### Column & Row Resize
- **Drag** the header divider to resize columns or rows
- **Double-click** the column header divider to **auto-fit** width (scans all cells and adjusts to longest content)

### Find & Replace
- **Ctrl+F** -- open Find bar
- **Ctrl+H** -- open Find bar (with Replace field)
- Type to search across all cells; matches are highlighted
- **Enter** / arrows to navigate between matches
- **Replace** / **Replace All** to substitute values
- **Esc** to close

### Filter
1. Select a cell in the column you want to filter
2. Click the filter button (funnel icon) in the toolbar
3. A dropdown appears with **checkboxes for each unique value** in the column
4. Uncheck values you want to hide
5. **Apply** to filter, **Clear** to remove filter, **Cancel** to close
6. Filtered (hidden) rows are skipped during rendering

### Conditional Formatting
1. Select the range of cells to apply the rule to
2. Click the conditional format button (colored squares icon)
3. Choose a condition:
   - **Greater than** / **Less than** -- compare cell value to a number
   - **Equals** / **Not equals** -- exact match
   - **Contains** / **Not contains** -- substring match
   - **Between** -- value is in a range (two values)
   - **Is empty** / **Is not empty** -- null/blank detection
4. Set the formatting style: background color, text color, bold, italic
5. Click **Apply** -- the rule is saved and cells update immediately
6. Existing rules are listed at the bottom of the dialog with a delete button (x)
7. Multiple rules can stack; first matching rule wins

### Sort
1. Select a range of cells
2. Click **A↓** for ascending or **Z↓** for descending
3. Sorts the selected range by the active column

### Merge Cells
Right-click a selection > **Merge Cells** to combine multiple cells into one. **Unmerge** to split them back.

### Insert & Delete
Right-click to access:
- **Insert Row Above / Below**
- **Insert Column Left / Right**
- **Delete Row / Column**

### Sheet Management
- **Sheet tabs** at the bottom show all sheets
- Click a tab to switch sheets
- **Double-click** a tab to rename
- **Right-click** a tab to delete
- Click **+** to add a new sheet

### Status Bar
When multiple cells are selected, the bottom-right shows:
- **Sum** -- total of all numeric values
- **Avg** -- average
- **Count** -- number of non-empty cells

---

## Templates

### Built-in Templates
When creating a new table (right-click folder > **New Table**), choose from:

| Template | Description |
|----------|-------------|
| **Blank** | Empty spreadsheet |
| **Budget** | Monthly budget with Income/Expenses categories, 12 months, currency format |
| **Invoice** | Invoice layout with Bill To, line items, quantities, amounts, total |
| **Timesheet** | Weekly timesheet with days, start/end times, break, hours |
| **Contacts** | Address book with name, email, phone, company, address fields |
| **Project Tracker** | Task list with #, task, assignee, priority, status, due date |

### Custom Templates
Save any spreadsheet as a reusable template:

1. Open the `.ods` file you want to use as a template
2. Run command palette (`Ctrl+Shift+P`) > **ODS: Save Current as Template**
3. Enter a template name and optional description
4. The template now appears in the **New Table** template picker under "Custom Templates"

Custom templates store the complete `.ods` file (data, styles, formulas, formatting) and are available across all workspaces.

---

## Formulas

### Supported Functions

**Math:** `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `CEILING`, `FLOOR`, `MOD`, `POWER`, `SQRT`, `PI`, `RAND`, `INT`, `SIGN`, `SUMIF`, `COUNTIF`, `AVERAGEIF`, `SUMPRODUCT`

**Text:** `LEN`, `LEFT`, `RIGHT`, `MID`, `UPPER`, `LOWER`, `TRIM`, `CONCATENATE`, `SUBSTITUTE`, `FIND`, `SEARCH`, `REPLACE`, `REPT`, `TEXT`, `VALUE`, `EXACT`, `T`

**Logic:** `IF`, `AND`, `OR`, `NOT`, `IFERROR`, `ISBLANK`, `ISNUMBER`, `ISTEXT`, `TRUE`, `FALSE`

**Lookup:** `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `CHOOSE`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`

**Date:** `TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`

### Reference Syntax
- Cell: `A1`, `$A$1` (absolute)
- Range: `A1:B10`
- Cross-sheet: `Sheet2.A1`, `Sheet2.A1:B10`
- ODS bracket: `[.A1]`, `[.Sheet2.A1]`

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `F2` or `Enter` | Start editing cell |
| `Esc` | Cancel editing |
| `Tab` / `Shift+Tab` | Move right / left |
| `Enter` / `Shift+Enter` | Move down / up |
| `Arrow keys` | Navigate cells |
| `Shift+Arrow` | Extend selection |
| `Ctrl+A` | Select all |
| `Ctrl+C` / `Ctrl+V` / `Ctrl+X` | Copy / Paste / Cut |
| `Ctrl+B` | Bold |
| `Ctrl+I` | Italic |
| `Ctrl+U` | Underline |
| `Ctrl+F` / `Ctrl+H` | Find / Find & Replace |
| `Delete` / `Backspace` | Clear cell content |
| `Home` | Go to column A in current row |
| `Ctrl+Home` | Go to cell A1 |
| `End` | Go to last used cell in current row |
| `Ctrl+End` | Go to last used cell in sheet |
| `Page Up` / `Page Down` | Scroll one page up / down |
| Address bar + `Enter` | Navigate to typed address (e.g. `B5`, `A1:C10`) |

---

## File Format

This extension works with **OpenDocument Spreadsheet (.ods)** files -- an open standard used by LibreOffice Calc, Google Sheets (export), and other applications. Files can be freely exchanged between this editor and any ODS-compatible application.
