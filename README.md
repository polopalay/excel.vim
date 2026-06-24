# Vim Excel Viewer

Edit Excel `.xlsx` files directly inside Vim or Neovim.

Vim Excel Viewer renders Excel worksheets as editable ASCII tables and writes changes back to the original workbook. The plugin is powered by a Rust backend and does not require Microsoft Excel, LibreOffice, WPS Office, Python, or OpenPyXL.

---

## Features

* Open `.xlsx` files directly in Vim/Neovim
* Edit worksheet data using normal Vim commands
* Save changes back to Excel workbooks
* Preserve merged cells
* Multiple worksheet support
* Add, rename, delete, and switch worksheets
* Automatic workbook reload after save
* Built-in syntax highlighting
* Pure Rust backend
* Automatic first-time build using Cargo
* No external office suite required

---

## Requirements

### Vim / Neovim

* Vim 8.2+
* Neovim 0.7+

### Rust

Rust and Cargo are required for the initial build.

Check:

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

## First Build

The plugin automatically builds the Rust backend on first use.

You can also build manually:

```vim
:ExcelBuild
```

The binary is generated under:

```text
rs/target/release/excel
```

or on Windows:

```text
rs/target/release/excel.exe
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

Example rendering:

```text
+----------+--------+
| Name     | Amount |
+----------+--------+
| Alice    | 100    |
+----------+--------+
| Bob      | 250    |
+----------+--------+
```

Edit normally using Vim motions.

Save:

```vim
:w
```

or

```vim
:ExcelSave
```

---

## Worksheet Commands

### List Worksheets

```vim
:ExcelSheets
```

### Open Worksheet

```vim
:ExcelSheetOpen Sheet1
```

### Create Worksheet

```vim
:ExcelSheetAdd NewSheet
```

### Rename Worksheet

```vim
:ExcelSheetRename
```

Interactive menu selection is displayed.

### Delete Worksheet

```vim
:ExcelSheetDelete Sheet1
```

Confirmation is required before deletion.

---

## Commands

| Command             | Description        |
| ------------------- | ------------------ |
| `:ExcelBuild`       | Build Rust backend |
| `:ExcelSave`        | Save workbook      |
| `:ExcelSheets`      | List worksheets    |
| `:ExcelSheetOpen`   | Open worksheet     |
| `:ExcelSheetAdd`    | Create worksheet   |
| `:ExcelSheetRename` | Rename worksheet   |
| `:ExcelSheetDelete` | Delete worksheet   |

---

## Syntax Highlighting

Built-in highlighting for:

* Numbers
* Dates
* URLs
* UPPERCASE headers
* Text inside parentheses
* Table borders

---

## Supported Features

| Feature         | Status |
| --------------- | ------ |
| Read XLSX       | ✓      |
| Write XLSX      | ✓      |
| Merged Cells    | ✓      |
| Multiple Sheets | ✓      |
| Add Sheet       | ✓      |
| Rename Sheet    | ✓      |
| Delete Sheet    | ✓      |
| Formula Results | ✓      |
| XLS Format      | ✗      |
| Charts Editing  | ✗      |
| Image Editing   | ✗      |

---

## ZIP Plugin Compatibility

Excel files are ZIP containers internally.

The plugin automatically overrides Vim's built-in `zip.vim` handlers for `.xlsx` files and takes ownership of workbook loading and saving.

No additional configuration is required.

---

## Limitations

* Only `.xlsx` files are supported
* Formulas are displayed as calculated values
* Charts are not editable
* Embedded images are ignored
* Complex Excel formatting is not rendered
* Cell styles are not preserved

---

## License

MIT

---

## Credits

Built with:

* Vim
* Neovim
* Rust
* ZIP
* quick-xml
