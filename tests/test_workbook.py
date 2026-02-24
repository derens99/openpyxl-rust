from openpyxl_rust.workbook import Workbook


def test_workbook_has_active_sheet():
    wb = Workbook()
    assert wb.active is not None
    assert wb.active.title == "Sheet"


def test_workbook_create_sheet():
    wb = Workbook()
    ws = wb.create_sheet("Data")
    assert ws.title == "Data"
    assert len(wb.sheetnames) == 2


def test_workbook_sheetnames():
    wb = Workbook()
    wb.create_sheet("A")
    wb.create_sheet("B")
    assert wb.sheetnames == ["Sheet", "A", "B"]


def test_workbook_getitem():
    wb = Workbook()
    wb.create_sheet("Data")
    ws = wb["Data"]
    assert ws.title == "Data"
