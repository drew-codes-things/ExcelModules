<div align="center">

# ExcelModules

**A collection of Excel VBA macros for data cleaning and formatting, with a matching Python CLI (`excel_utils.py`) for macro-disabled environments.**

[![VBA](https://img.shields.io/badge/VBA-Excel-217346?style=flat-square&logo=microsoft-excel&logoColor=white)](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office)
[![Python](https://img.shields.io/badge/python-3.8+-blue?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green?style=flat-square)](LICENSE)

</div>

---

Six data-cleaning operations available as standalone `.bas` VBA modules (import directly into the Excel VBA editor) and as a Python CLI (`excel_utils.py`) for users on LibreOffice, macOS, Linux, or any environment with macros disabled.

---

## Modules

| VBA file | Python function | What it does |
|----------|----------------|--------------|
| `Alphabetical order sort.bas` | `sort_columns()` | Sorts text columns alphabetically; numeric columns are skipped |
| `Brackets around 4 digits.bas` | `brackets_around_4_digits()` | Wraps standalone 4-digit numbers (e.g. years) in parentheses across all text cells |
| `Capitalise first words.bas` | `capitalise_proper()` / `capitalise_first_only()` | Applies Title Case or first-character-only capitalisation to every text cell |
| `Highlight duplicates.bas` | `highlight_duplicates_multi_column()` / `highlight_duplicates_across_sheets()` | Highlights duplicate rows across selected columns (yellow) or matching values across two sheets (red) |
| `Remove exact copies.bas` | `remove_exact_duplicates()` | Removes duplicate rows, keeping first, last, or removing all occurrences |
| `Strip whitespace.bas` | `strip_whitespace()` | Trims leading/trailing and double spaces, or removes all spaces entirely |

---

## VBA Usage

1. Open Excel and press `Alt + F11` to open the VBA editor
2. Go to **File -> Import File** and select the `.bas` file you want
3. Run the imported macro from **Tools -> Macros** or assign it to a button

---

## Python CLI Usage

The Python script mirrors all six VBA operations and writes output to a new file so the original is never overwritten. It accepts **`.xlsx` or `.csv`** input (CSV in -> CSV out for the data-transform operations; highlight operations always produce `.xlsx` since CSV can't store cell fills).

Run non-interactively-friendly with flags:

```
python excel_utils.py --input data.csv --output cleaned.csv
```

| Flag | Description |
|------|-------------|
| `-i`, `--input` | Path to the `.xlsx` / `.csv` file (skips the path prompt) |
| `-o`, `--output` | Exact output path, overriding the default `_processed` / `_highlighted` naming |

**Install dependencies:**

```
pip install openpyxl pandas
```

**Run:**

```
python excel_utils.py
```

You will be prompted for the path to your `.xlsx` file and then presented with a numbered menu:

```
 1  Alphabetical sort
 2  Brackets around 4-digit numbers
 3  Capitalise (Proper Case)
 4  Capitalise (first letter only)
 5  Highlight duplicates - multi-column
 6  Highlight duplicates - across two sheets
 7  Remove duplicate rows (keep first)
 8  Remove ALL duplicate rows
 9  Strip whitespace (trim)
10  Strip whitespace (remove all spaces)
 0  Exit
```

**Output files:**

| Operation | Output filename |
|-----------|-----------------|
| Most operations | `<original>_processed.xlsx` |
| Multi-column highlight | `<original>_highlighted.xlsx` |
| Cross-sheet highlight | `<original>_cross_highlighted.xlsx` |

---

## Requirements

- Python 3.8+
- `openpyxl` and `pandas` (`pip install openpyxl pandas`)

---

---

## Install as a command (pipx)

Install this folder as a CLI so it is available on your PATH:

```bash
pipx install .
excel-utils
```

Logging: set `LOG_LEVEL` (e.g. `DEBUG`) and `LOG_FILE` to also write logs to a file.


## Get the Code

Clone with git:

```bash
git clone https://github.com/drew-codes-things/ExcelModules.git
```

Or with the [GitHub CLI](https://cli.github.com/):

```bash
gh repo clone drew-codes-things/ExcelModules
```

## License

MIT - made by [Drew](https://github.com/drew-codes-things)
