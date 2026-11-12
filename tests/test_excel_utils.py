import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

import excel_utils


@pytest.fixture(autouse=True)
def no_output_override(monkeypatch):
    monkeypatch.setattr(excel_utils, "OUTPUT_OVERRIDE", None)


def make_workbook(path, sheets):
    workbook = Workbook()
    workbook.remove(workbook.active)
    for name, rows in sheets.items():
        sheet = workbook.create_sheet(name)
        for row in rows:
            sheet.append(row)
    workbook.save(path)
    return str(path)


def make_csv(path, rows):
    pd.DataFrame(rows[1:], columns=rows[0]).to_csv(path, index=False)
    return str(path)


def sheet_rows(path, sheet):
    return [list(row) for row in load_workbook(path)[sheet].iter_rows(values_only=True)]


PEOPLE = [["Name", "Year"], ["  bob  smith ", "born 1990"], ["alice", "1985 (1985)"], ["carol", None]]
NOTES = [["Note"], ["keep me"]]


@pytest.fixture
def workbook_path(tmp_path):
    return make_workbook(tmp_path / "people.xlsx", {"People": PEOPLE, "Notes": NOTES})


def test_output_is_written_beside_the_input_and_the_input_is_untouched(workbook_path):
    out = excel_utils.capitalise_proper(workbook_path)
    assert out.endswith("people_processed.xlsx")
    assert sheet_rows(workbook_path, "People") == PEOPLE


def test_other_sheets_are_kept(workbook_path):
    out = excel_utils.capitalise_proper(workbook_path)
    assert sheet_rows(out, "Notes") == NOTES


def test_capitalise_proper(workbook_path, capsys):
    out = excel_utils.capitalise_proper(workbook_path)
    names = [row[0] for row in sheet_rows(out, "People")[1:]]
    assert names == ["  Bob  Smith ", "Alice", "Carol"]
    assert "4 cell(s) converted" in capsys.readouterr().out


def test_capitalise_first_only(workbook_path):
    out = excel_utils.capitalise_first_only(workbook_path)
    assert sheet_rows(out, "People")[1][0] == "  bob  smith "
    assert sheet_rows(out, "People")[2][0] == "Alice"


def test_brackets_around_4_digits_skips_already_bracketed(workbook_path):
    out = excel_utils.brackets_around_4_digits(workbook_path)
    years = [row[1] for row in sheet_rows(out, "People")[1:]]
    assert years == ["born (1990)", "(1985) (1985)", None]


def test_brackets_ignore_longer_numbers(tmp_path):
    path = make_csv(tmp_path / "ids.csv", [["Id"], ["12345"], ["ref 2024"]])
    out = excel_utils.brackets_around_4_digits(path)
    assert pd.read_csv(out, dtype=str)["Id"].tolist() == ["12345", "ref (2024)"]


def test_strip_whitespace_trims_and_collapses(workbook_path):
    out = excel_utils.strip_whitespace(workbook_path)
    assert sheet_rows(out, "People")[1][0] == "bob smith"


def test_strip_whitespace_remove_all(workbook_path):
    out = excel_utils.strip_whitespace(workbook_path, remove_all=True)
    assert sheet_rows(out, "People")[1][0] == "bobsmith"


def test_sort_columns(tmp_path):
    path = make_csv(tmp_path / "names.csv", [["Name", "Age"], ["carol", 3], ["alice", 1], ["bob", 2]])
    out = excel_utils.sort_columns(path, ["Name"])
    assert out.endswith("names_processed.csv")
    assert pd.read_csv(out)["Age"].tolist() == [1, 2, 3]


def test_sort_columns_descending(tmp_path):
    path = make_csv(tmp_path / "names.csv", [["Name"], ["alice"], ["carol"], ["bob"]])
    out = excel_utils.sort_columns(path, ["Name"], ascending=False)
    assert pd.read_csv(out)["Name"].tolist() == ["carol", "bob", "alice"]


def test_sort_columns_without_text_columns_leaves_the_file(tmp_path):
    path = make_csv(tmp_path / "numbers.csv", [["Age"], [3], [1]])
    assert excel_utils.sort_columns(path, ["Age", "Missing"]) == path


def test_remove_exact_duplicates(tmp_path):
    rows = [["A", "B"], ["x", 1], ["x", 1], ["x", 2], ["y", 1]]
    path = make_csv(tmp_path / "dupes.csv", rows)
    out = excel_utils.remove_exact_duplicates(path)
    assert len(pd.read_csv(out)) == 3


def test_remove_duplicates_by_column_dropping_all(tmp_path):
    rows = [["A", "B"], ["x", 1], ["x", 2], ["y", 1]]
    path = make_csv(tmp_path / "dupes.csv", rows)
    out = excel_utils.remove_exact_duplicates(path, columns=["A"], keep=False)
    assert pd.read_csv(out)["A"].tolist() == ["y"]


def test_remove_duplicates_with_unknown_column_does_nothing(tmp_path, capsys):
    path = make_csv(tmp_path / "dupes.csv", [["A"], ["x"], ["x"]])
    assert excel_utils.remove_exact_duplicates(path, columns=["Nope"]) == path
    assert "Unknown column(s)" in capsys.readouterr().out


def test_highlight_duplicates_multi_column(tmp_path):
    rows = [["A", "B"], ["x", 1], ["x", 2], ["x", 1], [None, None], [None, None]]
    path = make_workbook(tmp_path / "dupes.xlsx", {"Data": rows})
    out = excel_utils.highlight_duplicates_multi_column(path, ["A", "B"])
    sheet = load_workbook(out)["Data"]
    highlighted = [row for row in range(2, 7) if sheet.cell(row, 1).fill.start_color.rgb.endswith("FFFF00")]
    assert highlighted == [2, 4]


def test_highlight_duplicates_from_csv_writes_xlsx(tmp_path):
    path = make_csv(tmp_path / "dupes.csv", [["A"], ["x"], ["x"]])
    out = excel_utils.highlight_duplicates_multi_column(path, ["A"])
    assert out.endswith("dupes_highlighted.xlsx")


def test_highlight_duplicates_across_sheets(tmp_path):
    path = make_workbook(tmp_path / "two.xlsx", {
        "First": [["Email"], ["a@x"], ["b@x"], ["c@x"]],
        "Second": [["Email"], ["c@x"], ["a@x"]],
    })
    out = excel_utils.highlight_duplicates_across_sheets(path, "Email", "First", "Second")
    sheet = load_workbook(out)["First"]
    red = [sheet.cell(row, 1).value for row in range(2, 5) if sheet.cell(row, 1).fill.start_color.rgb.endswith("FF0000")]
    assert red == ["a@x", "c@x"]


def test_highlight_across_unknown_sheet(workbook_path, capsys):
    assert excel_utils.highlight_duplicates_across_sheets(workbook_path, "Name", "People", "Missing") == workbook_path
    assert "Unknown sheet(s)" in capsys.readouterr().out


def test_output_override_for_xlsx_highlight(tmp_path, monkeypatch, workbook_path):
    monkeypatch.setattr(excel_utils, "OUTPUT_OVERRIDE", str(tmp_path / "custom.out"))
    assert excel_utils.resolve_output_xlsx(workbook_path, "_x") == str(tmp_path / "custom.xlsx")


def test_csv_input_with_xlsx_override(tmp_path, monkeypatch):
    path = make_csv(tmp_path / "names.csv", [["Name"], ["bob"]])
    target = tmp_path / "result.xlsx"
    monkeypatch.setattr(excel_utils, "OUTPUT_OVERRIDE", str(target))
    excel_utils.capitalise_proper(path)
    assert sheet_rows(target, "Sheet1") == [["Name"], ["Bob"]]


def test_csv_cells_the_operation_does_not_touch_are_written_back_unchanged(tmp_path):
    path = tmp_path / "people.csv"
    path.write_text("Name,Zip,Count,Note\n  alice ,01234,1,NA\nbob,00501,,null\n", encoding="utf-8")
    out = excel_utils.strip_whitespace(str(path))
    with open(out, encoding="utf-8") as written:
        assert written.read().splitlines() == ["Name,Zip,Count,Note", "alice,01234,1,NA", "bob,00501,,null"]


def test_xlsx_text_that_looks_like_a_missing_value_is_kept(tmp_path):
    path = make_workbook(tmp_path / "notes.xlsx", {"Notes": [["Name", "Note"], [" alice ", "NA"], ["bob", "null"], ["carol", None]]})
    out = excel_utils.strip_whitespace(path)
    assert sheet_rows(out, "Notes")[1:] == [["alice", "NA"], ["bob", "null"], ["carol", None]]


def test_sort_ignores_case_and_copes_with_numbers_among_the_text(tmp_path):
    path = make_workbook(tmp_path / "mixed.xlsx", {"Names": [["Name", "Row"], ["bob", 1], [12, 2], ["Zed", 3], ["alice", 4], [None, 5]]})
    out = excel_utils.sort_columns(path, ["Name"])
    assert sheet_rows(out, "Names")[1:] == [[12, 2], ["alice", 4], ["bob", 1], ["Zed", 3], [None, 5]]


@pytest.mark.parametrize("text, expected", [
    ("it's john's 2nd o'neil", "It's John's 2nd O'neil"),
    ("MARY-JANE mcDONALD", "Mary-jane Mcdonald"),
    ("  two\tspaced\nlines ", "  Two\tSpaced\nLines "),
])
def test_proper_case_matches_excel_proper_case(text, expected):
    assert excel_utils.proper_case(text) == expected


def test_capitalise_proper_keeps_apostrophes_and_ordinals(tmp_path):
    path = make_csv(tmp_path / "names.csv", [["Name"], ["it's john's 2nd o'neil"]])
    out = excel_utils.capitalise_proper(path)
    assert pd.read_csv(out, dtype=str)["Name"].tolist() == ["It's John's 2nd O'neil"]
