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

| Key | Description            |
| --- | ---------------------- |
| `y` | Copy current selection |
| `p` | Paste copied content   |

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
| `:columnWidth [width]` | `:cw`   | Set column width (shows current if no width) |
| `:write [filename]`    | `:w`    | Save file (save as filename if provided)     |
| `:quit`                | `:q`    | Exit program                                 |

### Undo/Redo

This only applies to actions on a per sheet basis (it does not track deleting sheets or creating them).

| Key      | Description        |
| -------- | ------------------ |
| `u`      | Undo last action   |
| `Ctrl+r` | Redo undone action |

### General

| Key      | Description           |
| -------- | --------------------- |
| `Ctrl+c` | Exit program          |
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
