from openpyxl_rust import Workbook


def test_empty_sheet_dimensions():
    wb = Workbook()
    ws = wb.active
    assert ws.min_row is None
    assert ws.max_row is None
    assert ws.min_column is None
    assert ws.max_column is None
    assert ws.dimensions == ""


def test_single_cell_dimensions():
    wb = Workbook()
    ws = wb.active
    ws["C5"] = "hello"
    assert ws.min_row == 5
    assert ws.max_row == 5
    assert ws.min_column == 3
    assert ws.max_column == 3
    assert ws.dimensions == "C5:C5"


def test_multiple_cells_dimensions():
    wb = Workbook()
    ws = wb.active
    ws["B2"] = 1
    ws["D10"] = 2
    ws["A1"] = 3
    assert ws.min_row == 1
    assert ws.max_row == 10
    assert ws.min_column == 1
    assert ws.max_column == 4
    assert ws.dimensions == "A1:D10"


def test_append_updates_dimensions():
    wb = Workbook()
    ws = wb.active
    ws.append([1, 2, 3])
    ws.append([4, 5, 6])
    assert ws.min_row == 1
    assert ws.max_row == 2
    assert ws.min_column == 1
    assert ws.max_column == 3
    assert ws.dimensions == "A1:C2"


def test_getitem_creates_cell_in_dimensions():
    wb = Workbook()
    ws = wb.active
    _ = ws["A1"]
    assert ws.min_row == 1
    assert ws.max_row == 1
