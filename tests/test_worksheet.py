from openpyxl_rust.styles import Font
from openpyxl_rust.worksheet import Worksheet


def test_worksheet_title():
    ws = Worksheet(title="Sheet1")
    assert ws.title == "Sheet1"
    ws.title = "Data"
    assert ws.title == "Data"


def test_setitem_string():
    ws = Worksheet()
    ws["A1"] = "hello"
    assert ws["A1"].value == "hello"


def test_setitem_number():
    ws = Worksheet()
    ws["B2"] = 42
    assert ws["B2"].value == 42


def test_cell_method():
    ws = Worksheet()
    cell = ws.cell(row=1, column=1, value="test")
    assert cell.value == "test"
    assert ws["A1"].value == "test"


def test_cell_font():
    ws = Worksheet()
    ws["A1"] = "bold"
    ws["A1"].font = Font(bold=True)
    assert ws["A1"].font.bold is True


def test_column_dimensions():
    ws = Worksheet()
    ws.column_dimensions["A"].width = 20
    assert ws.column_dimensions["A"].width == 20


def test_row_dimensions():
    ws = Worksheet()
    ws.row_dimensions[1].height = 30
    assert ws.row_dimensions[1].height == 30


def test_freeze_panes():
    ws = Worksheet()
    ws.freeze_panes = "A2"
    assert ws.freeze_panes == "A2"


def test_merge_cells():
    ws = Worksheet()
    ws.merge_cells("A1:D1")
    assert ("A1", "D1") in ws.merged_cell_ranges


def test_setitem_none():
    ws = Worksheet()
    ws["A1"] = None
    assert ws["A1"].value is None
