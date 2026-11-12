import argparse
import csv as csvlib
import re
import shutil
import sys
from pathlib import Path
from typing import Optional

try:
    import pandas as pd
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
except ImportError:
    print("Missing dependencies. Run: pip install openpyxl pandas")
    sys.exit(1)


YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
RED_FILL    = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

OUTPUT_OVERRIDE: Optional[str] = None
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


def load_df(path: str, sheet: Optional[str] = None) -> "tuple[pd.DataFrame, str]":
    if is_csv(path):
        return pd.read_csv(path, dtype=str, **ONLY_BLANK_CELLS_ARE_MISSING), "Sheet1"
    wb = load_workbook(path)
    sheet_name = sheet or wb.sheetnames[0]
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

    df = df.sort_values(by=text_cols, ascending=ascending, ignore_index=True, key=alphabetical_key)
    out = save_output(path, df, sheet)
    print(f"Sorted by {text_cols}. Output: {out}")
    return out


_FOUR_DIGIT = re.compile(r"(?<![\(\d])(\d{4})(?![\)\d])")


def brackets_around_4_digits(path: str) -> str:
    df, sheet = load_df(path)
    changed = transform_text(df, lambda v: _FOUR_DIGIT.sub(r"(\1)", v))
    out = save_output(path, df, sheet)
    print(f"{changed} cell(s) updated. Output: {out}")
    return out


def capitalise_proper(path: str) -> str:
    df, sheet = load_df(path)
    changed = transform_text(df, str.title)
    out = save_output(path, df, sheet)
    print(f"{changed} cell(s) converted. Output: {out}")
    return out


def capitalise_first_only(path: str) -> str:
    df, sheet = load_df(path)
    changed = transform_text(df, lambda v: v[:1].upper() + v[1:])
    out = save_output(path, df, sheet)
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
    before = len(df)
    df = df.drop_duplicates(subset=columns, keep=keep, ignore_index=True)
    removed = before - len(df)
    out = save_output(path, df, sheet)
    print(f"{removed} duplicate row(s) removed. Output: {out}")
    return out


def strip_whitespace(path: str, remove_all: bool = False) -> str:
    df, sheet = load_df(path)
    if remove_all:
        changed = transform_text(df, lambda v: v.replace(" ", ""))
    else:
        changed = transform_text(df, lambda v: " ".join(v.split()))
    out = save_output(path, df, sheet)
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
