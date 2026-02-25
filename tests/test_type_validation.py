import pytest
from openpyxl_rust import Workbook


def test_unsupported_type_list():
    wb = Workbook()
    ws = wb.active
    with pytest.raises(TypeError, match="Unsupported cell value type: list"):
        ws["A1"] = [1, 2, 3]


def test_unsupported_type_dict():
    wb = Workbook()
    ws = wb.active
    with pytest.raises(TypeError, match="Unsupported cell value type: dict"):
        ws["A1"] = {"key": "value"}


def test_unsupported_type_set():
    wb = Workbook()
    ws = wb.active
    with pytest.raises(TypeError, match="Unsupported cell value type: set"):
        ws["A1"] = {1, 2, 3}


def test_unsupported_type_custom_object():
    class Foo:
        pass
    wb = Workbook()
    ws = wb.active
    with pytest.raises(TypeError, match="Unsupported cell value type: Foo"):
        ws["A1"] = Foo()


def test_supported_types_no_error():
    from datetime import datetime, date
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "string"
    ws["A2"] = 42
    ws["A3"] = 3.14
    ws["A4"] = True
    ws["A5"] = datetime(2024, 1, 1)
    ws["A6"] = date(2024, 1, 1)
    ws["A7"] = None
    ws["A8"] = "=SUM(A1:A2)"
