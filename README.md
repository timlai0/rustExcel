# xlview

A fast, read-only Excel viewer for whole folders. Point it at a folder and it opens **one window** containing every workbook, navigated through three rows of tabs:

1. **Folders** – the subfolders (`All` and `(root)` included)
2. **Workbooks** – the files in the selected folder
3. **Sheets** – the sheets of the selected workbook

Single static `.exe`, no installer, no runtime needed. Files are never modified (nothing is written, there is no edit mode). Cached cell values are shown, not formulas.

Supported: `.xlsx .xlsm .xlsb .xls .ods .xlam` (Excel `~$` lock files are skipped).

## Usage

```
xlview.exe C:\path\to\folder
```

or start it empty and then drag a folder onto the window, press `Ctrl+O`, use **Open folder…**, or type a path in the path box and press Enter.

## Search

| Action | How |
|---|---|
| Focus search box | `Ctrl+F` |
| Next / previous hit | `Enter` / `Shift+Enter` in the box, or `F3` / `Shift+F3` |
| Match case | **Aa** toggle |
| Scope | This sheet · This workbook · This folder · Everything |
| Jump to a hit | click it in the results list |

Hits are highlighted in the grid (current hit in orange). Results are capped at 100 000.

## Command palette

`Ctrl+P` (or `Ctrl+Shift+P`) opens the palette. Type to fuzzy-filter, `↑`/`↓` to choose, `Enter` to run, `Esc` to close.

Commands include: open/reload folder, search (focus, clear, next/prev, case, scope, results list), copy cell, next/previous sheet / workbook / folder, toggle light/dark theme, and **Go to workbook: …** / **Go to sheet: …** for every workbook and sheet.

## Grid

- Click a cell to select it; arrow keys / `PageUp` / `PageDown` move; `Ctrl+C` copies it.
- The status bar shows the full text of the selected cell and the sheet size.
- Headers and row numbers stay fixed while scrolling; each sheet remembers its scroll position.
- `Ctrl` + `+` / `-` zooms.

## Performance notes

Workbooks load in background threads (the one you click is loaded first), search runs on its own thread, and only visible cells are drawn. All workbooks of the folder are held in memory, so huge folders use a lot of RAM.

## Build

Requires Rust and the MSVC build tools.

```
cargo build --release
```

Result: `target\release\xlview.exe` (CRT is statically linked; only system DLLs are used). On a Windows-on-ARM machine without the ARM64 MSVC linker, use the x64 toolchain: `cargo +stable-x86_64-pc-windows-msvc build --release`.

## Limitations

No editing, formulas are not displayed, merged cells are not drawn merged, text is not wrapped.
