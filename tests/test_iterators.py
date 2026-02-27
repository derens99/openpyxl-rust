from openpyxl_rust import Cell, Workbook


def test_iter_rows_basic():
    wb = Workbook()
    ws = wb.active
    ws.append([1, 2, 3])
    ws.append([4, 5, 6])
    rows = list(ws.iter_rows())
    assert len(rows) == 2
    assert len(rows[0]) == 3
    assert all(isinstance(c, Cell) for c in rows[0])
    assert rows[0][0].value == 1
    assert rows[1][2].value == 6


def test_iter_rows_values_only():
    wb = Workbook()
    ws = wb.active
    ws.append(["a", "b"])
    ws.append(["c", "d"])
    rows = list(ws.iter_rows(values_only=True))
    assert rows == [("a", "b"), ("c", "d")]


def test_iter_rows_with_bounds():
    wb = Workbook()
    ws = wb.active
    ws.append([1, 2, 3])
    ws.append([4, 5, 6])
    ws.append([7, 8, 9])
    rows = list(ws.iter_rows(min_row=2, max_row=3, min_col=1, max_col=2, values_only=True))
    assert rows == [(4, 5), (7, 8)]


def test_iter_rows_empty_sheet():
    wb = Workbook()
    ws = wb.active
    rows = list(ws.iter_rows())
    assert rows == []


def test_iter_cols_basic():
    wb = Workbook()
    ws = wb.active
    ws.append([1, 2])
    ws.append([3, 4])
    cols = list(ws.iter_cols(values_only=True))
    assert cols == [(1, 3), (2, 4)]


def test_iter_cols_with_bounds():
    wb = Workbook()
    ws = wb.active
    ws.append([1, 2, 3])
    ws.append([4, 5, 6])
    cols = list(ws.iter_cols(min_col=2, max_col=3, values_only=True))
    assert cols == [(2, 5), (3, 6)]


def test_values_property():
    wb = Workbook()
    ws = wb.active
    ws.append([10, 20])
    ws.append([30, 40])
    vals = list(ws.values)
    assert vals == [(10, 20), (30, 40)]


def test_iter_rows_sparse_data():
    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="A1")
    ws.cell(row=1, column=3, value="C1")
    rows = list(ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=3, values_only=True))
    assert rows == [("A1", None, "C1")]
