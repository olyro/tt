# Table TUI

Table TUI `tt` is a Terminal User Interface (TUI) for editing Microsoft Excel
files, inspired by Vim key bindings. It deviates from vim in a couple of places
(see the Key Bindings section).

## Demo

![Demo](./demo.svg)

## Features

- Create, open and save Excel files (.xlsx)
- Browse, Create and Delete Excel Sheets
- Vim-like navigation and editing
- Basic Search and replace
- Undo/Redo functionality
- Copy and paste
- System clipboard interoperability using TSV
- Merging and Unmerging
- Row and column selection
- Block selection
- Insert rows and columns
- Basic Formula Support

## Installation

Get the binary via `go install`

```bash
go install github.com/olyro/tt@latest
```

or check out the repository and build it from source

```bash
go build -o tt
```

## Usage

```bash
# Create new Excel file
tt

# Open existing Excel file
tt file.xlsx

# Show help
tt --help

# Show version
tt --version
```

## Key Bindings

### Navigation

| Key           | Description         |
| ------------- | ------------------- |
| `h`, `←`, `b` | Move left           |
| `j`, `↓`, `w` | Move down           |
| `k`, `↑`      | Move up             |
| `l`, `→`      | Move right          |
| `0`           | Go to row beginning |
| `$`           | Go to row end       |
| `gg`          | Go to first row     |
| `G`           | Go to last row      |
| `[Num]G`      | Jump to row [Num]   |
| `Ctrl+d`      | Move page down      |
| `Ctrl+u`      | Move page up        |

### Editing

| Key     | Description                                                                                         |
| ------- | --------------------------------------------------------------------------------------------------- |
| `i`     | Edit current cell (cursor at beginning)                                                             |
| `a`     | Edit current cell (cursor at end)                                                                   |
| `c`     | Change cell content (clear and edit)                                                                |
| `x`     | Clears the selected cells                                                                           |
| `d`     | Deletes the selected rows or columns, in case of block and cell select it clears the selected cells |
| `Enter` | Confirm input                                                                                       |
| `Esc`   | Cancel editing / return to normal mode                                                              |

### Rows and Columns

| Key            | Description                           |
| -------------- | ------------------------------------- |
| `I`            | Insert column before current position |
| `A`            | Insert column after current position  |
| `O`            | Insert row before current position    |
| `o`            | Insert row after current position     |
| `[Num]I/A/O/o` | Insert multiple rows/columns          |

### Selection

| Key      | Description            |
| -------- | ---------------------- |
| `v`      | Start column selection |
| `V`      | Start row selection    |
| `Ctrl+v` | Start block selection  |

### Merge

Merged Cells are highlighted, only the top left value is shown.

| Key | Description                                           |
| --- | ----------------------------------------------------- |
| `m` | Merge block selection (undoing restores old values)   |
| `M` | Unmerge block selection (does not restore old values) |

### Copy and Paste

| Key | Description                                                    |
| --- | -------------------------------------------------------------- |
| `y` | Copy current selection                                         |
| `p` | Paste copied content                                           |
| `Y` | Copy selection to the system clipboard as TSV                  |
| `P` | Paste TSV from the system clipboard at the current cell        |

System clipboard paste overlays cells starting at the cursor and is recorded as
one undoable operation. Tabs separate columns and newlines separate rows, making
the format compatible with Excel, LibreOffice, and other spreadsheet tools.
Clipboard access requires the platform clipboard service to be available.

### Search

| Key | Description            |
| --- | ---------------------- |
| `/` | Start search           |
| `n` | Next search result     |
| `N` | Previous search result |

### Commands

| Key | Description       |
| --- | ----------------- |
| `:` | Open command mode |

#### Available Commands

| Command                | Short   | Description                                  |
| ---------------------- | ------- | -------------------------------------------- |
| `:sheet [name]`        | `:b`    | Switch to sheet (shows current if no name)   |
| `:nextSheet`           | `:bn`   | Switch to next sheet                         |
| `:previousSheet`       | `:bp`   | Switch to previous sheet                     |
| `:deleteSheet [name]`  | `:bd`   | Delete sheet (current sheet if no name)      |
| `:addSheet <name>`     | `:badd` | Create new sheet with given name             |
| `:autoWidth [column]`  | `:aw`   | Auto-fit one column, or all populated columns |
| `:edit [filepath]`     | `:e`    | Open file, or reload current file if omitted |
| `:edit! [filepath]`    | `:e!`   | Open file ignoring non-empty undo history    |
| `:write [filename]`    | `:w`    | Save file (save as filename if provided)     |
| `:quit`                | `:q`    | Exit program                                 |
| `:quit!`               | `:q!`   | Exit program ignoring non-empty undo history |

While entering `:b ...`, press `Tab` to complete a matching sheet name. Column
widths are read from and written to the active Excel worksheet, so auto-fitted
widths are preserved when the workbook is saved. `:aw` sizes every populated
column from its longest displayed line (including calculated formula results);
pass a column name such as `:aw C` to resize only that column.

### Undo/Redo

Undo history is maintained per sheet. Creating and deleting sheets is tracked as
an unsaved change but is not itself undoable. Row and column deletion restores
cell values, formulas, and styles; worksheet metadata such as custom dimensions
or conditional formatting may still require reloading the original file.

| Key      | Description        |
| -------- | ------------------ |
| `u`      | Undo last action   |
| `Ctrl+r` | Redo undone action |

### General

| Key      | Description           |
| -------- | --------------------- |
| `Ctrl+c` | Exit if there are no unsaved changes |
| `Esc`    | Return to normal mode |

## Number Prefixes

Most navigation commands support number prefixes:

- `5j` - Move 5 rows down
- `10l` - Move 10 columns right
- `3I` - Insert 3 columns

## Modes

The program has different modes:

- **NORMAL**: Standard navigation mode
- **COMMAND**: Command input (with `:`)
- **SEARCH**: Search mode (with `/`)
- **INPUT**: Cell editing

## Cell Types

You can see the current cell type by executing `:type` or `:t`. Formulas have to
be start with `=`. Numbers are automatically recognized. To force a string
prefix your input with `'`.

## Dependencies

Built on top of these great projects:

- [Bubble Tea](https://github.com/charmbracelet/bubbletea) - TUI Framework
- [Excelize](https://github.com/xuri/excelize) - Excel Library
- [Bubbles](https://github.com/charmbracelet/bubbles) - TUI Components
- [Lipgloss](https://github.com/charmbracelet/lipgloss) - Styling

## License

This project is licensed under the MIT License.
