# Vim Excel Viewer

Open, edit and save `.xlsx` files directly inside Vim/Neovim.

Instead of launching Microsoft Excel, LibreOffice or WPS Office, this plugin converts Excel worksheets into an ASCII table that can be edited directly in Vim. Changes are written back to the original `.xlsx` file while preserving merged cells.

---

## Features

* Open `.xlsx` files directly in Vim/Neovim
* Edit cell contents using normal Vim motions
* Save changes back to the original Excel workbook
* Preserve merged cells
* Insert rows above or below current row
* Insert columns left or right of current column
* Syntax highlighting for:

  * Numbers
  * Dates
  * URLs
  * Uppercase headers
  * Comments in parentheses
* No external office suite required
* Fast loading for large spreadsheets

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

### Python Libraries

Install:

```bash
pip install openpyxl
```

Required libraries:

| Library               | Purpose                   |
| --------------------- | ------------------------- |
| openpyxl              | Read/write Excel files    |
| zipfile               | Fast merge-cell detection |
| xml.etree.ElementTree | Parse workbook metadata   |

---
## Important: Disable ZIP Handling for XLSX

Vim's built-in `zipPlugin` treats `.xlsx` files as ZIP archives because the XLSX format is internally a ZIP container.

As a result, opening:

```bash
nvim report.xlsx
```

may open the workbook as a ZIP archive instead of using Vim Excel Viewer.

### Recommended Fix

Exclude `.xlsx` from the zip plugin extension list:

```vim
let g:zipPlugin_ext = '*.zip,*.jar,*.xpi,*.apk,*.war,*.ear'
```

### Alternative

Disable the ZIP plugin completely:

```vim
let g:loaded_zip = 1
let g:loaded_zipPlugin = 1
```

### Why?

An `.xlsx` file is actually a ZIP archive containing XML files:

```text
report.xlsx
├── xl/
├── docProps/
├── _rels/
└── [Content_Types].xml
```

Without this configuration, Vim may intercept the file before Vim Excel Viewer has a chance to load it.

If opening an XLSX file shows ZIP contents instead of a spreadsheet table, check your zipPlugin configuration first.

---

## Installation

### vim-plug

```vim
Plug 'yourname/vim-excel-viewer'
```

Then:

```vim
:PlugInstall
```

### Lazy.nvim

```lua
{
    "yourname/vim-excel-viewer",
}
```

---

## Usage

Open any Excel file:

```bash
vim report.xlsx
```

or

```bash
nvim report.xlsx
```

The spreadsheet will be displayed as an editable ASCII table:

```text
+----------+------------+
| Name     | Amount     |
+----------+------------+
| Alice    | 100        |
+----------+------------+
| Bob      | 250        |
+----------+------------+
```

Edit cells normally.

Save:

```vim
:w
```

Changes are written directly to the original `.xlsx` file.

---

## Commands

### Reload Workbook

```vim
:ExcelReload
```

Reload the workbook from disk.

### Save Workbook

```vim
:ExcelSave
```

Save and refresh formatting.

---

## Row Operations

### Insert Row Above

```vim
:call ExcelInsertRowAbove()
```

Insert a blank row above the current row.

### Insert Row Below

```vim
:call ExcelInsertRowBelow()
```

Insert a blank row below the current row.

---

## Column Operations

### Insert Column Left

```vim
:call ExcelInsertColLeft()
```

Insert a blank column to the left.

### Insert Column Right

```vim
:call ExcelInsertColRight()
```

Insert a blank column to the right.

---

## Suggested Key Mappings

Add these to your `vimrc` or `init.lua`.

### Vimscript

```vim
autocmd FileType excel nnoremap <buffer> <leader>ra :call ExcelInsertRowAbove()<CR>
autocmd FileType excel nnoremap <buffer> <leader>rb :call ExcelInsertRowBelow()<CR>

autocmd FileType excel nnoremap <buffer> <leader>cl :call ExcelInsertColLeft()<CR>
autocmd FileType excel nnoremap <buffer> <leader>cr :call ExcelInsertColRight()<CR>

autocmd FileType excel nnoremap <buffer> <leader>er :ExcelReload<CR>
autocmd FileType excel nnoremap <buffer> <leader>es :ExcelSave<CR>
```

### Default Mapping Table

| Mapping      | Action              |
| ------------ | ------------------- |
| `<leader>ra` | Insert row above    |
| `<leader>rb` | Insert row below    |
| `<leader>cl` | Insert column left  |
| `<leader>cr` | Insert column right |
| `<leader>es` | Save workbook       |
| `<leader>er` | Reload workbook     |

---

## Supported Features

| Feature                   | Status                |
| ------------------------- | --------------------- |
| Read XLSX                 | V                     |
| Write XLSX                | V                     |
| Merged Cells              | V                     |
| Insert Rows               | V                     |
| Insert Columns            | V                     |
| Large Files               | V                     |
| Formulas (display result) | V                     |
| Preserve Formatting       | Partial               |
| Multiple Sheets           | X (active sheet only) |
| XLS                       | X                     |
| CSV                       | X                     |

---

## How It Works

1. `.xlsx` file is opened.
2. Python reads workbook using `openpyxl`.
3. Worksheet is converted to an ASCII table.
4. User edits text inside Vim.
5. On save:

   * ASCII table is parsed back into rows and columns.
   * Workbook is updated.
   * Merge regions are restored.
6. Original `.xlsx` file is overwritten.

---

## Example Workflow

```bash
nvim invoices.xlsx
```

Edit:

```text
| Invoice | Amount |
| INV001  | 1000   |
```

Save:

```vim
:w
```

Workbook is updated immediately.

---

## Limitations

* Only the active worksheet is supported.
* Excel formulas are saved as displayed values.
* Charts are not editable.
* Images are ignored.
* Complex formatting is not rendered in Vim.

---

## License

MIT License

---

## Credits

Built with:

* Vim
* Neovim
* Python
* OpenPyXL

