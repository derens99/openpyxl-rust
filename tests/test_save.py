# tests/test_save.py
import os
import tempfile
from io import BytesIO

from openpyxl_rust import Workbook
from openpyxl_rust.styles import Font


def test_save_to_file():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Hello"
    ws["B1"] = 42

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        wb.save(path)
        assert os.path.exists(path)
        assert os.path.getsize(path) > 0
    finally:
        os.unlink(path)


def test_save_to_buffer():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Hello"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    data = buf.read()
    # xlsx files start with PK (ZIP magic)
    assert data[:2] == b"PK"
    assert len(data) > 0


def test_save_with_font():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Bold"
    ws["A1"].font = Font(bold=True, size=14, name="Arial")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        wb.save(path)
        assert os.path.getsize(path) > 0
    finally:
        os.unlink(path)


def test_save_with_number_format():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = 1234.5
    ws["A1"].number_format = "$#,##0.00"

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        wb.save(path)
        assert os.path.getsize(path) > 0
    finally:
        os.unlink(path)


def test_save_multiple_sheets():
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "First"
    ws1["A1"] = "Sheet 1"

    ws2 = wb.create_sheet("Second")
    ws2["A1"] = "Sheet 2"

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        wb.save(path)
        assert os.path.getsize(path) > 0
    finally:
        os.unlink(path)


def test_save_with_freeze_panes():
    wb = Workbook()
    ws = wb.active
    ws.freeze_panes = "A2"
    ws["A1"] = "Header"

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        wb.save(path)
        assert os.path.getsize(path) > 0
    finally:
        os.unlink(path)


def test_save_with_column_width_and_row_height():
    wb = Workbook()
    ws = wb.active
    ws.column_dimensions["A"].width = 25
    ws.row_dimensions[1].height = 40
    ws["A1"] = "Wide and tall"

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        wb.save(path)
        assert os.path.getsize(path) > 0
    finally:
        os.unlink(path)


def test_save_with_merged_cells():
    wb = Workbook()
    ws = wb.active
    ws.merge_cells("A1:D1")
    ws["A1"] = "Merged"

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        wb.save(path)
        assert os.path.getsize(path) > 0
    finally:
        os.unlink(path)


def test_save_formula():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = 10
    ws["A2"] = 20
    ws["A3"] = "=SUM(A1:A2)"

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        wb.save(path)
        assert os.path.getsize(path) > 0
    finally:
        os.unlink(path)


def test_save_boolean():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = True
    ws["A2"] = False

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        wb.save(path)
        assert os.path.getsize(path) > 0
    finally:
        os.unlink(path)
