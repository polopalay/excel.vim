# Vim Excel Viewer

Edit Excel `.xlsx` files directly inside Vim or Neovim.

Vim Excel Viewer converts Excel worksheets into editable ASCII tables, allowing you to view and modify spreadsheet data without leaving Vim. Changes are written back to the original workbook while preserving merged cells.

---

## Features

* Open `.xlsx` files directly in Vim/Neovim
* Edit worksheet data using normal Vim commands
* Save changes back to the original Excel file
* Preserve merged-cell layouts
* Fast merge-cell detection using direct XML parsing
* Multiple worksheet support
* List workbook sheets from Vim
* Open a specific worksheet without leaving Vim
* Automatic reload after save
* No Microsoft Excel, LibreOffice, or WPS Office required

### Syntax Highlighting

Built-in highlighting for:

* Numbers
* Dates
* URLs
* UPPERCASE headers
* Text inside parentheses
* Table borders

---

## Requirements

### Vim / Neovim

* Vim 8.2+
* Neovim 0.7+

### Python

Python 3 is required.

Check:

```bash
python3 --version
```

### Python Packages

Install:

```bash
pip install openpyxl
```

---

## Installation

### vim-plug

```vim
Plug 'polopalay/excel.vim'
```

Install:

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

## Opening Excel Files

```bash
nvim report.xlsx
```

or

```bash
vim report.xlsx
```

The worksheet is displayed as an editable ASCII table:

```text
+----------+--------+
| Name     | Amount |
+----------+--------+
| Alice    | 100    |
+----------+--------+
| Bob      | 250    |
+----------+--------+
```

Edit cells normally using Vim motions.

Save:

```vim
:w
```

The original workbook is updated automatically.

---

## Working With Sheets

### List Available Sheets

```vim
:ExcelSheets
```

Example:

```text
Sheet1
Customers
Invoices
Summary
```

### Open a Specific Sheet

```vim
:ExcelOpenSheet Customers
```

Switches the current buffer to the selected worksheet.

---

## Commands

### Save Workbook

```vim
:ExcelSave
```

Saves the current worksheet and refreshes formatting.

### List Sheets

```vim
:ExcelSheets
```

Displays all worksheet names in the workbook.

### Open Sheet

```vim
:ExcelOpenSheet <sheet-name>
```

Opens the specified worksheet.

---

## Supported Features

| Feature              | Status |
| -------------------- | ------ |
| Read XLSX            | ✓      |
| Write XLSX           | ✓      |
| Merged Cells         | ✓      |
| Multiple Sheets      | ✓      |
| Formula Results      | ✓      |
| Large Files          | ✓      |
| Fast Merge Detection | ✓      |
| XLS Format           | ✗      |
| CSV Mode             | ✗      |
| Charts Editing       | ✗      |
| Image Editing        | ✗      |

---

## How It Works

1. User opens an `.xlsx` file.
2. Python loads workbook data using OpenPyXL.
3. Merge information is extracted.
4. Worksheet is rendered as an ASCII table.
5. User edits data directly inside Vim.
6. On save:

   * ASCII table is parsed back into rows and columns.
   * Workbook data is updated.
   * Merge regions are restored.
7. Workbook is written back to disk.

---

## ZIP Plugin Compatibility

Excel files are ZIP containers internally.

Many Vim installations load the built-in `zip.vim` plugin, which may try to open `.xlsx` files as archives.

Vim Excel Viewer automatically removes ZIP handlers registered for `.xlsx` files and takes ownership of opening and saving Excel workbooks.

No additional configuration is normally required.

---

## Limitations

* Only `.xlsx` files are supported
* Formulas are displayed as calculated values
* Charts are not editable
* Embedded images are ignored
* Complex Excel formatting is not rendered
* Cell styles are not preserved when editing data

---

## Example Workflow

Open workbook:

```bash
nvim invoices.xlsx
```

List sheets:

```vim
:ExcelSheets
```

Open worksheet:

```vim
:ExcelOpenSheet Invoices
```

Edit values:

```text
| INV001 | 1000 |
| INV002 | 2500 |
```

Save:

```vim
:w
```

Workbook is updated immediately.

---

## License

MIT

---

## Credits

Built with:

* Vim
* Neovim
* Python
* OpenPyXL
