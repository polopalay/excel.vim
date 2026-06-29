# Vim Excel Viewer

Edit Microsoft Excel `.xlsx` workbooks directly inside Vim or Neovim.

Vim Excel Viewer renders Excel worksheets as editable ASCII tables while preserving workbook structure, formatting, and formulas. Changes are written back to the original workbook using a fast Rust backend with a built-in Excel formula engine.

No Microsoft Excel, LibreOffice, WPS Office, Python, or OpenPyXL is required.

---

## Features

### Workbook

* Open `.xlsx` files directly in Vim or Neovim
* Save changes back to the original workbook
* Multiple worksheet support
* Add worksheets
* Rename worksheets
* Delete worksheets
* Switch between worksheets
* Automatic reload after saving
* Automatic first-time Rust build
* Pure Rust backend

### Cell Editing

* Edit cells using normal Vim commands
* Jump directly to any cell (`A1`, `B25`, ...)
* Statusline displays the current worksheet, active cell, and formula
* Supports Normal mode and Visual mode operations

### Formula Engine

* Display calculated formula results
* View the original formula of the current cell from the statusline
* Toggle between calculated values and formulas
* Automatically recalculate formulas after editing
* Formula dependency evaluation
* Circular reference detection
* Built-in formula parser and evaluator
* No external calculation engine required

### Formatting

* Preserve existing formatting
* Preserve merged cells
* Merge cells
* Unmerge cells
* Toggle **Bold**
* Toggle *Italic*
* Change font color
* Change background color
* Supports named colors and custom `#RRGGBB` colors

### Rendering

* Automatic column sizing
* Built-in syntax highlighting
* URLs
* Numbers
* Dates
* Table borders
* Real Excel formatting (bold, italic, foreground color, background color)

---

## Requirements

### Vim / Neovim

* Vim 8.2+
* Neovim 0.7+

### Rust

Rust and Cargo are required only for the initial build.

Verify installation:

```bash
cargo --version
```

Install Rust:

```bash
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
```

---

## Installation

### vim-plug

```vim
Plug 'polopalay/excel.vim'
```

Then:

```vim
:PlugInstall
```

### lazy.nvim

```lua
{
    "polopalay/excel.vim",
}
```

---

## First Build

The Rust backend is built automatically the first time an Excel file is opened.

To build manually:

```vim
:ExcelBuild
```

Generated binary:

Linux / macOS

```text
rs/target/release/excel
```

Windows

```text
rs/target/release/excel.exe
```

---

## Opening Workbooks

```bash
vim report.xlsx
```

or

```bash
nvim report.xlsx
```

Example:

```text
+----------+--------+
| Name     | Amount |
+----------+--------+
| Alice    | 100    |
+----------+--------+
| Bob      | 250    |
+----------+--------+
```

Save normally:

```vim
:w
```

or

```vim
:ExcelSave
```

---

# Worksheet Commands

| Command                    | Description      |
| -------------------------- | ---------------- |
| `:ExcelSheets`             | List worksheets  |
| `:ExcelSheetOpen Sheet1`   | Open worksheet   |
| `:ExcelSheetAdd NewSheet`  | Create worksheet |
| `:ExcelSheetRename`        | Rename worksheet |
| `:ExcelSheetDelete Sheet1` | Delete worksheet |

---

# Cell Navigation

Jump directly to a cell.

```vim
:ExcelGoto C15
```

Tab completion is supported.

The current worksheet, active cell, and formula (if present) are displayed automatically in the Vim statusline.

Example:

```text
Sheet1 / C15  =SUM(A1:A10)
```

---

# Formula Engine

By default, worksheets display calculated values, just like Microsoft Excel.

You can toggle between displaying calculated values and the original formulas.

Show formulas:

```vim
:ExcelShowFormula
```

Example:

Calculated values:

```text
+--------+
| 125    |
+--------+
```

Formula view:

```text
+--------------+
| =SUM(A1:A5)  |
+--------------+
```

Whenever the workbook is saved, all supported formulas are recalculated automatically before being written back to the `.xlsx` file.

---

# Supported Formula Features

## Operators

Arithmetic

```text
+
-
*
/
^
%
```

Comparison

```text
=
<>
<
<=
>
>=
```

Concatenation

```text
&
```

## References

```text
A1
B25
$A$1
A1:A10
B2:D8
```

## Functions

```text
SUM
AVERAGE
MIN
MAX
COUNT
ABS
ROUND
IF
AND
OR
NOT
CONCAT
CONCATENATE
```

Additional features

* Automatic dependency evaluation
* Circular reference detection (`#CIRCULAR!`)
* Formula result caching

---

# Formatting Commands

## Font Color

```vim
:ExcelSetFg red
```

Visual mode:

```vim
:'<,'>ExcelSetFg blue
```

## Background Color

```vim
:ExcelSetBg yellow
```

Visual mode:

```vim
:'<,'>ExcelSetBg "#FFE599"
```

## Bold

```vim
:ExcelBold
```

Visual selection:

```vim
:'<,'>ExcelBold
```

## Italic

```vim
:ExcelItalic
```

Visual selection:

```vim
:'<,'>ExcelItalic
```

---

# Merge Cells

Merge selected cells:

```vim
:'<,'>ExcelMerge
```

Unmerge:

```vim
:ExcelUnmerge
```

or

```vim
:'<,'>ExcelUnmerge
```

---

# Supported Colors

Named colors:

```text
red
green
blue
yellow
orange
purple
gray
white
black
none
```

Custom colors:

```text
#RRGGBB
```

Example:

```vim
:ExcelSetBg "#D9EAD3"
```

---

# Commands

| Command             | Description             |
| ------------------- | ----------------------- |
| `:ExcelBuild`       | Build Rust backend      |
| `:ExcelSave`        | Save workbook           |
| `:ExcelSheets`      | List worksheets         |
| `:ExcelSheetOpen`   | Open worksheet          |
| `:ExcelSheetAdd`    | Add worksheet           |
| `:ExcelSheetRename` | Rename worksheet        |
| `:ExcelSheetDelete` | Delete worksheet        |
| `:ExcelGoto`        | Jump to a cell          |
| `:ExcelShowFormula` | Toggle formula display  |
| `:ExcelSetFg`       | Change font color       |
| `:ExcelSetBg`       | Change background color |
| `:ExcelBold`        | Toggle bold             |
| `:ExcelItalic`      | Toggle italic           |
| `:ExcelMerge`       | Merge cells             |
| `:ExcelUnmerge`     | Unmerge cells           |

---

# Syntax Highlighting

Built-in highlighting includes:

* Numbers
* Dates
* URLs
* Table borders
* Existing Excel bold text
* Existing Excel italic text
* Existing Excel font colors
* Existing Excel background colors

---

# Supported Features

| Feature                      | Status |
| ---------------------------- | ------ |
| Read XLSX                    | ✓      |
| Write XLSX                   | ✓      |
| Formula Evaluation           | ✓      |
| Formula Viewer               | ✓      |
| Automatic Recalculation      | ✓      |
| Dependency Evaluation        | ✓      |
| Circular Reference Detection | ✓      |
| Multiple Worksheets          | ✓      |
| Add Worksheet                | ✓      |
| Rename Worksheet             | ✓      |
| Delete Worksheet             | ✓      |
| Merge Cells                  | ✓      |
| Unmerge Cells                | ✓      |
| Preserve Merged Cells        | ✓      |
| Bold                         | ✓      |
| Italic                       | ✓      |
| Font Color                   | ✓      |
| Background Color             | ✓      |
| Cell Navigation              | ✓      |
| Statusline Formula Preview   | ✓      |
| XLS Format                   | ✗      |
| Charts Editing               | ✗      |
| Embedded Images              | ✗      |
| Pivot Tables                 | ✗      |
| VBA Macros                   | ✗      |

---

# ZIP Plugin Compatibility

Excel workbooks are ZIP containers internally.

The plugin automatically overrides Vim's built-in `zip.vim` handlers for `.xlsx` files, so no additional configuration is required.

---

# Limitations

* Only `.xlsx` files are supported
* Cross-worksheet references are not yet supported
* Shared formulas are not yet supported
* Named ranges are not yet supported
* Charts are not editable
* Embedded images are ignored
* Pivot tables are not editable
* VBA macros are preserved but cannot be edited
* Extremely complex Excel layouts may not render identically to Microsoft Excel

---

# License

MIT

---

# Credits

Built with

* Rust
* Vim
* Neovim
* ZIP
* quick-xml
