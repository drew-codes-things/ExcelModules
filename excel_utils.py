import argparse
import csv as csvlib
import re
import shutil
import sys
from copy import copy
from pathlib import Path
from typing import Optional

try:
    import pandas as pd
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.formula.translate import Translator
    from openpyxl.styles import PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Missing dependencies. Run: pip install openpyxl pandas")
    sys.exit(1)


YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
RED_FILL    = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

OUTPUT_OVERRIDE: Optional[str] = None
FIRST_DATA_ROW = 2
FORMULA = "f"
TEXT = "s"
ONLY_BLANK_CELLS_ARE_MISSING = {"keep_default_na": False, "na_values": [""]}


def is_csv(path: str) -> bool:
    return Path(path).suffix.lower() == ".csv"


def resolve_output(path: str, default_suffix: str, ext: str) -> str:
    if OUTPUT_OVERRIDE:
        return OUTPUT_OVERRIDE
    source = Path(path)
    return str(source.with_name(source.stem + default_suffix + ext))


def resolve_output_xlsx(path: str, default_suffix: str) -> str:
    if OUTPUT_OVERRIDE:
        o = OUTPUT_OVERRIDE
        if not o.lower().endswith(".xlsx"):
            o = str(Path(o).with_suffix(".xlsx"))
        return o
    source = Path(path)
    return str(source.with_name(source.stem + default_suffix + ".xlsx"))


def processed_sheet_name(workbook) -> str:
    return workbook.active.title


def load_df(path: str, sheet: Optional[str] = None) -> "tuple[pd.DataFrame, str]":
    if is_csv(path):
        return pd.read_csv(path, dtype=str, **ONLY_BLANK_CELLS_ARE_MISSING), "Sheet1"
    wb = load_workbook(path)
    sheet_name = sheet or processed_sheet_name(wb)
    df = pd.read_excel(path, sheet_name=sheet_name, **ONLY_BLANK_CELLS_ARE_MISSING)
    return df, sheet_name


def load_ws_workbook(path: str):
    if is_csv(path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        with open(path, newline="", encoding="utf-8") as f:
            for row in csvlib.reader(f):
                ws.append(row)
        return wb
    return load_workbook(path)


def save_output(path: str, df: pd.DataFrame, sheet: str) -> str:
    ext = ".csv" if is_csv(path) else ".xlsx"
    out = resolve_output(path, "_processed", ext)
    if out.lower().endswith(".csv"):
        df.to_csv(out, index=False)
    elif not is_csv(path):
        if Path(out).resolve() != Path(path).resolve():
            shutil.copyfile(path, out)
        with pd.ExcelWriter(out, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)
    else:
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)
    return out


def edits_xlsx_in_place(path: str) -> bool:
    return not is_csv(path) and not resolve_output(path, "_processed", ".xlsx").lower().endswith(".csv")


def save_workbook(path: str, workbook) -> str:
    out = resolve_output(path, "_processed", ".xlsx")
    workbook.save(out)
    return out


def set_value(cell, value, data_type: str) -> None:
    cell.value = value
    if data_type != FORMULA and isinstance(value, str):
        cell.data_type = TEXT


def change_sheet_text(sheet, change) -> int:
    changed = 0
    for row in sheet.iter_rows(min_row=FIRST_DATA_ROW):
        for cell in row:
            if not isinstance(cell.value, str) or cell.data_type == FORMULA:
                continue
            new_value = change(cell.value)
            if new_value != cell.value:
                set_value(cell, new_value, TEXT)
                changed += 1
    return changed


def change_text_cells(path: str, change) -> "tuple[int, str]":
    if not edits_xlsx_in_place(path):
        df, sheet = load_df(path)
        changed = transform_text(df, change)
        return changed, save_output(path, df, sheet)
    workbook = load_workbook(path)
    changed = change_sheet_text(workbook[processed_sheet_name(workbook)], change)
    return changed, save_workbook(path, workbook)


def moved_value(value, data_type: str, column: int, source_row: int, target_row: int):
    if data_type != FORMULA or not isinstance(value, str) or source_row == target_row:
        return value
    letter = get_column_letter(column)
    return Translator(value, origin=f"{letter}{source_row}").translate_formula(f"{letter}{target_row}")


def rewrite_data_rows(sheet, kept_positions: list, data_row_count: int) -> None:
    snapshot = [[(cell.value, cell.data_type, copy(cell._style)) for cell in row]
                for row in sheet.iter_rows(min_row=FIRST_DATA_ROW, max_row=FIRST_DATA_ROW + data_row_count - 1,
                                           max_col=sheet.max_column)]
    for target_offset, source_offset in enumerate(kept_positions):
        source_row = FIRST_DATA_ROW + source_offset
        target_row = FIRST_DATA_ROW + target_offset
        for column, (value, data_type, style) in enumerate(snapshot[source_offset], 1):
            cell = sheet.cell(target_row, column)
            set_value(cell, moved_value(value, data_type, column, source_row, target_row), data_type)
            cell._style = copy(style)
    removed = data_row_count - len(kept_positions)
    if removed:
        sheet.delete_rows(FIRST_DATA_ROW + len(kept_positions), removed)


def save_kept_rows(path: str, df: "pd.DataFrame", kept: "pd.DataFrame", sheet: str) -> str:
    if not edits_xlsx_in_place(path):
        return save_output(path, kept.reset_index(drop=True), sheet)
    workbook = load_workbook(path)
    rewrite_data_rows(workbook[sheet], list(kept.index), len(df))
    return save_workbook(path, workbook)


def count_changed(before: "pd.Series", after: "pd.Series") -> int:
    same = before.eq(after) | (before.isna() & after.isna())
    return int((~same).sum())


TEXT_DTYPES = ["object", "string"]


def is_text_column(series: "pd.Series") -> bool:
    if not (pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series)):
        return False
    values = series.dropna()
    return bool(pd.to_numeric(values, errors="coerce").isna().any())


def alphabetical_key(series: "pd.Series") -> "pd.Series":
    return series.map(lambda v: v if pd.isna(v) else str(v).casefold())


def transform_text(df: "pd.DataFrame", change) -> int:
    changed = 0
    for col in df.select_dtypes(include=TEXT_DTYPES).columns:
        original = df[col].copy()
        df[col] = df[col].apply(lambda v: change(v) if isinstance(v, str) else v)
        changed += count_changed(original, df[col])
    return changed


def require_columns(df: "pd.DataFrame", columns: list) -> bool:
    missing = [c for c in columns if c not in df.columns]
    if missing:
        print(f"Unknown column(s): {missing}. Available: {list(df.columns)}")
        return False
    return True


def sort_columns(path: str, columns: list, ascending: bool = True) -> str:
    df, sheet = load_df(path)

    existing_cols = [c for c in columns if c in df.columns]
    missing_cols = [c for c in columns if c not in df.columns]
    if missing_cols:
        print(f"Ignoring unknown column(s): {missing_cols}")

    text_cols = [c for c in existing_cols if is_text_column(df[c])]
    if not text_cols:
        print("No text columns found in the specified list - nothing sorted.")
        return path

    sorted_df = df.sort_values(by=text_cols, ascending=ascending, key=alphabetical_key)
    out = save_kept_rows(path, df, sorted_df, sheet)
    print(f"Sorted by {text_cols}. Output: {out}")
    return out


_DIGIT_RUN = re.compile(r"\d+(?:[ -]\d+)*")
_DIGIT_RUN_SEPARATOR = re.compile(r"[ -]")
_NUMBER_CELL = re.compile(r"\s*[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?\s*")
_DECIMAL_BEFORE = re.compile(r"\d[.,]$")
_DECIMAL_AFTER = re.compile(r"^[.,]\d")
YEAR_DIGITS = 4
MAX_YEARS_IN_A_RUN = 2


def looks_like_number(text: str) -> bool:
    try:
        float(text)
    except ValueError:
        return bool(_NUMBER_CELL.fullmatch(text))
    return True


def is_standalone_year_run(text: str, run: "re.Match") -> bool:
    groups = _DIGIT_RUN_SEPARATOR.split(run.group(0))
    if len(groups) > MAX_YEARS_IN_A_RUN or any(len(group) != YEAR_DIGITS for group in groups):
        return False
    before = text[:run.start()]
    after = text[run.end():]
    if before.endswith("(") or after.startswith(")"):
        return False
    return not (_DECIMAL_BEFORE.search(before) or _DECIMAL_AFTER.search(after))


def bracket_four_digit_numbers(text: str) -> str:
    if looks_like_number(text):
        return text

    def bracket_run(run: "re.Match") -> str:
        if not is_standalone_year_run(text, run):
            return run.group(0)
        return re.sub(r"\d+", r"(\g<0>)", run.group(0))

    return _DIGIT_RUN.sub(bracket_run, text)


def brackets_around_4_digits(path: str) -> str:
    changed, out = change_text_cells(path, bracket_four_digit_numbers)
    print(f"{changed} cell(s) updated. Output: {out}")
    return out


_PROPER_CASE_WORD = re.compile(r"[^\x00\t\n\v\f\r ]+")


def proper_case(text: str) -> str:
    return _PROPER_CASE_WORD.sub(lambda word: word.group(0)[:1].upper() + word.group(0)[1:].lower(), text)


def capitalise_proper(path: str) -> str:
    changed, out = change_text_cells(path, proper_case)
    print(f"{changed} cell(s) converted. Output: {out}")
    return out


def capitalise_first_only(path: str) -> str:
    changed, out = change_text_cells(path, lambda v: v[:1].upper() + v[1:])
    print(f"{changed} cell(s) updated. Output: {out}")
    return out


def highlight_duplicates_multi_column(path: str, columns: list) -> str:
    from collections import Counter
    wb = load_ws_workbook(path)
    ws = wb.active

    header = {cell.value: cell.column for cell in ws[1]}
    col_indices = [header[c] for c in columns if c in header]
    if not col_indices:
        print("No valid columns found for duplicate highlighting.")
        return path

    def row_key(row):
        values = [ws.cell(row, ci).value for ci in col_indices]
        if all(v is None or str(v).strip() == "" for v in values):
            return None
        return "\x00".join(str(v) for v in values)

    keys = [row_key(r) for r in range(2, ws.max_row + 1)]
    counts = Counter(k for k in keys if k is not None)
    highlighted = 0
    for i, key in enumerate(keys):
        if key is not None and counts[key] > 1:
            for ci in col_indices:
                ws.cell(i + 2, ci).fill = YELLOW_FILL
            highlighted += 1

    out = resolve_output_xlsx(path, "_highlighted")
    wb.save(out)
    print(f"{highlighted} row(s) highlighted. Output: {out}")
    return out


def highlight_duplicates_across_sheets(path: str, col_name: str, sheet_a: str, sheet_b: str) -> str:
    if is_csv(path):
        print("Cross-sheet highlight needs a multi-sheet .xlsx, not a CSV.")
        return path
    wb = load_workbook(path)
    unknown = [name for name in (sheet_a, sheet_b) if name not in wb.sheetnames]
    if unknown:
        print(f"Unknown sheet(s): {unknown}. Available: {wb.sheetnames}")
        return path
    ws_a = wb[sheet_a]
    ws_b = wb[sheet_b]

    header_a = {cell.value: cell.column for cell in ws_a[1]}
    header_b = {cell.value: cell.column for cell in ws_b[1]}
    ci_a = header_a.get(col_name)
    ci_b = header_b.get(col_name)
    if not ci_a or not ci_b:
        print(f"Column '{col_name}' not found in one or both sheets.")
        return path

    vals_b = {ws_b.cell(r, ci_b).value for r in range(2, ws_b.max_row + 1)} - {None}
    highlighted = 0
    for r in range(2, ws_a.max_row + 1):
        cell = ws_a.cell(r, ci_a)
        if cell.value in vals_b:
            cell.fill = RED_FILL
            highlighted += 1

    out = resolve_output_xlsx(path, "_cross_highlighted")
    wb.save(out)
    print(f"{highlighted} cell(s) highlighted in red. Output: {out}")
    return out


def remove_exact_duplicates(path: str, columns: Optional[list] = None,
                            keep: str = "first") -> str:
    df, sheet = load_df(path)
    if columns and not require_columns(df, columns):
        return path
    kept = df.drop_duplicates(subset=columns, keep=keep)
    removed = len(df) - len(kept)
    out = save_kept_rows(path, df, kept, sheet)
    print(f"{removed} duplicate row(s) removed. Output: {out}")
    return out


def strip_whitespace(path: str, remove_all: bool = False) -> str:
    if remove_all:
        changed, out = change_text_cells(path, lambda v: v.replace(" ", ""))
    else:
        changed, out = change_text_cells(path, lambda v: " ".join(v.split()))
    print(f"{changed} cell(s) updated. Output: {out}")
    return out


MENU = """
============================
  ExcelModules - Python CLI
============================
 1  Alphabetical sort
 2  Brackets around 4-digit numbers
 3  Capitalise (Proper Case)
 4  Capitalise (first letter only)
 5  Highlight duplicates - multi-column
 6  Highlight duplicates - across two sheets
 7  Remove duplicate rows (keep first)
 8  Remove ALL duplicate rows
 9  Strip whitespace (trim)
 10 Strip whitespace (remove all spaces)
 0  Exit
"""


def parse_args():
    p = argparse.ArgumentParser(description="ExcelModules - Python CLI (xlsx/csv data cleaning).")
    p.add_argument("-i", "--input", default=None, help="Path to the .xlsx or .csv file (skips the prompt)")
    p.add_argument("-o", "--output", default=None,
                   help="Exact output path for the result (overrides the default _processed/_highlighted naming)")
    return p.parse_args()


def main():
    global OUTPUT_OVERRIDE
    args = parse_args()
    OUTPUT_OVERRIDE = args.output

    print(MENU)
    path = args.input or input("Enter path to your .xlsx or .csv file: ").strip().strip('"')
    if not Path(path).exists():
        print("File not found.")
        return

    choice = input("Choose an option: ").strip()

    if choice == "1":
        cols = input("Column names to sort by (comma-separated): ").split(",")
        asc  = input("Ascending? (y/n): ").lower() != "n"
        sort_columns(path, [c.strip() for c in cols], ascending=asc)
    elif choice == "2":
        brackets_around_4_digits(path)
    elif choice == "3":
        capitalise_proper(path)
    elif choice == "4":
        capitalise_first_only(path)
    elif choice == "5":
        cols = input("Column names to check (comma-separated): ").split(",")
        highlight_duplicates_multi_column(path, [c.strip() for c in cols])
    elif choice == "6":
        col   = input("Column name to compare: ").strip()
        sh_a  = input("Sheet A name: ").strip()
        sh_b  = input("Sheet B name: ").strip()
        highlight_duplicates_across_sheets(path, col, sh_a, sh_b)
    elif choice == "7":
        cols = input("Columns to deduplicate on (comma-separated, or Enter for all): ").strip()
        remove_exact_duplicates(path, columns=[c.strip() for c in cols.split(",")] if cols else None, keep="first")
    elif choice == "8":
        cols = input("Columns to deduplicate on (comma-separated, or Enter for all): ").strip()
        remove_exact_duplicates(path, columns=[c.strip() for c in cols.split(",")] if cols else None, keep=False)
    elif choice == "9":
        strip_whitespace(path)
    elif choice == "10":
        strip_whitespace(path, remove_all=True)
    elif choice == "0":
        sys.exit(0)
    else:
        print("Invalid choice.")


if __name__ == "__main__":
    main()
