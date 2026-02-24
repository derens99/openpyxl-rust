from openpyxl_rust.cell import Cell


def test_cell_stores_value():
    c = Cell(row=1, column=1)
    c.value = "hello"
    assert c.value == "hello"


def test_cell_stores_number():
    c = Cell(row=1, column=1, value=42.5)
    assert c.value == 42.5


def test_cell_coordinate():
    c = Cell(row=1, column=1)
    assert c.coordinate == "A1"


def test_cell_coordinate_multi_letter():
    c = Cell(row=1, column=27)
    assert c.coordinate == "AA1"


def test_cell_number_format():
    c = Cell(row=1, column=1, value=100)
    c.number_format = "$#,##0.00"
    assert c.number_format == "$#,##0.00"
