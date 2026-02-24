# openpyxl_rust Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a Rust-backed Python package that provides an openpyxl-compatible API for writing .xlsx files, dramatically faster than pure-Python openpyxl.

**Architecture:** Python classes (Workbook, Worksheet, Cell, Font) collect data in Python dicts. At save time, the data is passed to a PyO3 Rust function that uses `rust_xlsxwriter` to generate the .xlsx file. This "collect then flush" design minimizes FFI calls.

**Tech Stack:** Rust + rust_xlsxwriter, PyO3 + maturin, Python 3.9+, pytest

**System notes:** Python is invoked via `py` on this Windows system. Rust 1.93 is installed. maturin and openpyxl must be pip-installed.

---

### Task 1: Project Scaffolding

**Files:**
- Create: `Cargo.toml`
- Create: `pyproject.toml`
- Create: `src/lib.rs`
- Create: `python/openpyxl_rust/__init__.py`

**Step 1: Initialize git repo**

```bash
cd /c/Users/Deren/Desktop/Files/Coding/openpyxl-rust
git init
```

**Step 2: Create Cargo.toml**

```toml
[package]
name = "openpyxl-rust"
version = "0.1.0"
edition = "2021"

[lib]
name = "_openpyxl_rust"
crate-type = ["cdylib"]

[dependencies]
pyo3 = { version = "0.24", features = ["extension-module"] }
rust_xlsxwriter = "0.93"
```

**Step 3: Create pyproject.toml**

```toml
[build-system]
requires = ["maturin>=1.0,<2.0"]
build-backend = "maturin"

[project]
name = "openpyxl_rust"
requires-python = ">=3.9"
version = "0.1.0"
description = "Rust-backed openpyxl-compatible Excel writer"

[tool.maturin]
features = ["pyo3/extension-module"]
python-source = "python"
module-name = "openpyxl_rust._openpyxl_rust"
```

**Step 4: Create minimal src/lib.rs**

```rust
use pyo3::prelude::*;

/// Placeholder — will be replaced with real save logic.
#[pyfunction]
fn _save_workbook() -> PyResult<()> {
    Ok(())
}

#[pymodule]
fn _openpyxl_rust(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(_save_workbook, m)?)?;
    Ok(())
}
```

**Step 5: Create python/openpyxl_rust/__init__.py**

```python
from openpyxl_rust._openpyxl_rust import _save_workbook
```

**Step 6: Install maturin and build**

```bash
py -m pip install maturin
cd /c/Users/Deren/Desktop/Files/Coding/openpyxl-rust
maturin develop
```

**Step 7: Verify the module imports**

```bash
py -c "from openpyxl_rust._openpyxl_rust import _save_workbook; print('OK')"
```

Expected: `OK`

**Step 8: Commit**

```bash
git add Cargo.toml pyproject.toml src/lib.rs python/
git commit -m "feat: project scaffolding with PyO3 + maturin"
```

---

### Task 2: Cell class

**Files:**
- Create: `python/openpyxl_rust/cell.py`
- Create: `tests/test_cell.py`

**Step 1: Write failing test**

```python
# tests/test_cell.py
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
```

**Step 2: Run test to verify it fails**

```bash
py -m pytest tests/test_cell.py -v
```

Expected: FAIL — `ImportError`

**Step 3: Implement Cell class**

```python
# python/openpyxl_rust/cell.py


def _col_letter(col_idx):
    """Convert 1-based column index to Excel column letter(s). 1->A, 27->AA."""
    result = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


class Cell:
    def __init__(self, row=1, column=1, value=None):
        self.row = row
        self.column = column
        self.value = value
        self.font = None
        self.number_format = "General"

    @property
    def coordinate(self):
        return f"{_col_letter(self.column)}{self.row}"
```

**Step 4: Run test to verify it passes**

```bash
py -m pytest tests/test_cell.py -v
```

Expected: all PASS

**Step 5: Commit**

```bash
git add python/openpyxl_rust/cell.py tests/test_cell.py
git commit -m "feat: Cell class with value, coordinate, number_format"
```

---

### Task 3: Font class

**Files:**
- Create: `python/openpyxl_rust/styles/__init__.py`
- Create: `python/openpyxl_rust/styles/fonts.py`
- Create: `tests/test_styles.py`

**Step 1: Write failing test**

```python
# tests/test_styles.py
from openpyxl_rust.styles import Font


def test_font_defaults():
    f = Font()
    assert f.bold is False
    assert f.italic is False
    assert f.underline is None
    assert f.name == "Calibri"
    assert f.size == 11


def test_font_bold():
    f = Font(bold=True, size=14, name="Arial")
    assert f.bold is True
    assert f.size == 14
    assert f.name == "Arial"


def test_font_color():
    f = Font(color="FF0000")
    assert f.color == "FF0000"


def test_font_equality():
    f1 = Font(bold=True)
    f2 = Font(bold=True)
    assert f1 == f2


def test_font_inequality():
    f1 = Font(bold=True)
    f2 = Font(bold=False)
    assert f1 != f2
```

**Step 2: Run test to verify it fails**

```bash
py -m pytest tests/test_styles.py -v
```

Expected: FAIL — `ImportError`

**Step 3: Implement Font**

```python
# python/openpyxl_rust/styles/fonts.py
class Font:
    def __init__(self, name="Calibri", size=11, bold=False, italic=False,
                 underline=None, color=None):
        self.name = name
        self.size = size
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.color = color

    def __eq__(self, other):
        if not isinstance(other, Font):
            return NotImplemented
        return (self.name == other.name and self.size == other.size
                and self.bold == other.bold and self.italic == other.italic
                and self.underline == other.underline and self.color == other.color)

    def __repr__(self):
        return (f"Font(name={self.name!r}, size={self.size}, bold={self.bold}, "
                f"italic={self.italic}, underline={self.underline!r}, color={self.color!r})")
```

```python
# python/openpyxl_rust/styles/__init__.py
from openpyxl_rust.styles.fonts import Font

__all__ = ["Font"]
```

**Step 4: Run test to verify it passes**

```bash
py -m pytest tests/test_styles.py -v
```

Expected: all PASS

**Step 5: Commit**

```bash
git add python/openpyxl_rust/styles/ tests/test_styles.py
git commit -m "feat: Font class with name, size, bold, italic, underline, color"
```

---

### Task 4: Worksheet class

**Files:**
- Create: `python/openpyxl_rust/worksheet.py`
- Create: `tests/test_worksheet.py`

**Step 1: Write failing test**

```python
# tests/test_worksheet.py
from openpyxl_rust.worksheet import Worksheet
from openpyxl_rust.styles import Font


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
```

**Step 2: Run test to verify it fails**

```bash
py -m pytest tests/test_worksheet.py -v
```

Expected: FAIL — `ImportError`

**Step 3: Implement Worksheet**

```python
# python/openpyxl_rust/worksheet.py
import re
from openpyxl_rust.cell import Cell, _col_letter


def _parse_cell_ref(ref_str):
    """Parse 'A1' -> (row, col) as 1-based integers."""
    m = re.match(r"^([A-Z]+)(\d+)$", ref_str.upper())
    if not m:
        raise ValueError(f"Invalid cell reference: {ref_str}")
    letters, row_str = m.group(1), m.group(2)
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - 64)
    return int(row_str), col


class ColumnDimension:
    def __init__(self):
        self.width = None


class RowDimension:
    def __init__(self):
        self.height = None


class _ColumnDimensionsDict:
    def __init__(self):
        self._dims = {}

    def __getitem__(self, key):
        if key not in self._dims:
            self._dims[key] = ColumnDimension()
        return self._dims[key]

    def items(self):
        return self._dims.items()


class _RowDimensionsDict:
    def __init__(self):
        self._dims = {}

    def __getitem__(self, key):
        if key not in self._dims:
            self._dims[key] = RowDimension()
        return self._dims[key]

    def items(self):
        return self._dims.items()


class Worksheet:
    def __init__(self, title="Sheet"):
        self.title = title
        self._cells = {}  # (row, col) -> Cell
        self.column_dimensions = _ColumnDimensionsDict()
        self.row_dimensions = _RowDimensionsDict()
        self.freeze_panes = None
        self.merged_cell_ranges = []

    def __setitem__(self, key, value):
        row, col = _parse_cell_ref(key)
        if (row, col) in self._cells:
            self._cells[(row, col)].value = value
        else:
            self._cells[(row, col)] = Cell(row=row, column=col, value=value)

    def __getitem__(self, key):
        row, col = _parse_cell_ref(key)
        if (row, col) not in self._cells:
            self._cells[(row, col)] = Cell(row=row, column=col)
        return self._cells[(row, col)]

    def cell(self, row, column, value=None):
        if (row, column) in self._cells:
            cell = self._cells[(row, column)]
            if value is not None:
                cell.value = value
            return cell
        c = Cell(row=row, column=column, value=value)
        self._cells[(row, column)] = c
        return c

    def merge_cells(self, range_string):
        parts = range_string.split(":")
        if len(parts) != 2:
            raise ValueError(f"Invalid merge range: {range_string}")
        self.merged_cell_ranges.append((parts[0].upper(), parts[1].upper()))
```

**Step 4: Run test to verify it passes**

```bash
py -m pytest tests/test_worksheet.py -v
```

Expected: all PASS

**Step 5: Commit**

```bash
git add python/openpyxl_rust/worksheet.py tests/test_worksheet.py
git commit -m "feat: Worksheet class with cell access, dimensions, freeze, merge"
```

---

### Task 5: Workbook class (Python-side, no Rust save yet)

**Files:**
- Create: `python/openpyxl_rust/workbook.py`
- Create: `tests/test_workbook.py`
- Modify: `python/openpyxl_rust/__init__.py`

**Step 1: Write failing test**

```python
# tests/test_workbook.py
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
```

**Step 2: Run test to verify it fails**

```bash
py -m pytest tests/test_workbook.py -v
```

Expected: FAIL — `ImportError`

**Step 3: Implement Workbook**

```python
# python/openpyxl_rust/workbook.py
from openpyxl_rust.worksheet import Worksheet


class Workbook:
    def __init__(self):
        self._sheets = [Worksheet(title="Sheet")]

    @property
    def active(self):
        return self._sheets[0] if self._sheets else None

    def create_sheet(self, title=None):
        title = title or f"Sheet{len(self._sheets) + 1}"
        ws = Worksheet(title=title)
        self._sheets.append(ws)
        return ws

    @property
    def sheetnames(self):
        return [s.title for s in self._sheets]

    def __getitem__(self, name):
        for s in self._sheets:
            if s.title == name:
                return s
        raise KeyError(f"Worksheet '{name}' not found")

    def save(self, filename):
        raise NotImplementedError("Rust save not yet wired up")
```

**Step 4: Update __init__.py**

```python
# python/openpyxl_rust/__init__.py
from openpyxl_rust.workbook import Workbook
from openpyxl_rust.worksheet import Worksheet
from openpyxl_rust.cell import Cell

__all__ = ["Workbook", "Worksheet", "Cell"]
```

**Step 5: Run test to verify it passes**

```bash
py -m pytest tests/test_workbook.py -v
```

Expected: all PASS

**Step 6: Commit**

```bash
git add python/openpyxl_rust/workbook.py python/openpyxl_rust/__init__.py tests/test_workbook.py
git commit -m "feat: Workbook class with active sheet, create_sheet, getitem"
```

---

### Task 6: Rust save engine

**Files:**
- Modify: `src/lib.rs`

This is the core Rust implementation. The `_save_workbook` function receives a Python dict describing the entire workbook and uses `rust_xlsxwriter` to produce the file.

**Step 1: Write the Rust implementation**

```rust
// src/lib.rs
use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyBytes, PyDict, PyList};
use rust_xlsxwriter::{Format, Formula, Workbook, XlsxError};

fn make_format(font_dict: &Bound<'_, PyDict>) -> Result<Format, PyErr> {
    let mut fmt = Format::new();

    if let Ok(Some(bold)) = font_dict.get_item("bold") {
        if bold.extract::<bool>()? {
            fmt = fmt.set_bold();
        }
    }
    if let Ok(Some(italic)) = font_dict.get_item("italic") {
        if italic.extract::<bool>()? {
            fmt = fmt.set_italic();
        }
    }
    if let Ok(Some(underline)) = font_dict.get_item("underline") {
        let ul: String = underline.extract()?;
        if ul == "single" {
            fmt = fmt.set_underline(rust_xlsxwriter::FormatUnderline::Single);
        } else if ul == "double" {
            fmt = fmt.set_underline(rust_xlsxwriter::FormatUnderline::Double);
        }
    }
    if let Ok(Some(name)) = font_dict.get_item("name") {
        let n: String = name.extract()?;
        fmt = fmt.set_font_name(&n);
    }
    if let Ok(Some(size)) = font_dict.get_item("size") {
        let s: f64 = size.extract()?;
        fmt = fmt.set_font_size(s);
    }
    if let Ok(Some(color)) = font_dict.get_item("color") {
        let c: String = color.extract()?;
        if let Ok(rgb) = u32::from_str_radix(&c, 16) {
            fmt = fmt.set_font_color(rust_xlsxwriter::Color::from(rgb));
        }
    }

    Ok(fmt)
}

fn xlsx_err(e: XlsxError) -> PyErr {
    pyo3::exceptions::PyRuntimeError::new_err(e.to_string())
}

#[pyfunction]
fn _save_workbook(py: Python<'_>, data: &Bound<'_, PyDict>) -> PyResult<PyObject> {
    let mut workbook = Workbook::new();

    let sheets: &Bound<'_, PyList> = data.get_item("sheets")?.unwrap().downcast()?;

    for sheet_obj in sheets.iter() {
        let sheet_dict: &Bound<'_, PyDict> = sheet_obj.downcast()?;
        let title: String = sheet_dict
            .get_item("title")?
            .unwrap()
            .extract()?;

        let worksheet = workbook.add_worksheet();
        worksheet.set_name(&title).map_err(xlsx_err)?;

        // Write cells
        let cells: &Bound<'_, PyList> = sheet_dict.get_item("cells")?.unwrap().downcast()?;
        for cell_obj in cells.iter() {
            let cell: &Bound<'_, PyDict> = cell_obj.downcast()?;
            let row: u32 = cell.get_item("row")?.unwrap().extract()?;
            let col: u16 = cell.get_item("col")?.unwrap().extract()?;

            // Build format if font or number_format present
            let mut has_format = false;
            let mut fmt = Format::new();

            if let Ok(Some(font_obj)) = cell.get_item("font") {
                let font_dict: &Bound<'_, PyDict> = font_obj.downcast()?;
                fmt = make_format(font_dict)?;
                has_format = true;
            }

            if let Ok(Some(nf)) = cell.get_item("number_format") {
                let nf_str: String = nf.extract()?;
                if nf_str != "General" {
                    fmt = fmt.set_num_format(&nf_str);
                    has_format = true;
                }
            }

            let value_obj = cell.get_item("value")?.unwrap();

            if value_obj.is_none() {
                // blank cell — skip or write blank
                if has_format {
                    worksheet.write_blank(row, col, &fmt).map_err(xlsx_err)?;
                }
            } else if let Ok(v) = value_obj.extract::<bool>() {
                // Must check bool before i64/f64 because Python bool is subclass of int
                if has_format {
                    worksheet.write_boolean_with_format(row, col, v, &fmt).map_err(xlsx_err)?;
                } else {
                    worksheet.write_boolean(row, col, v).map_err(xlsx_err)?;
                }
            } else if let Ok(v) = value_obj.extract::<f64>() {
                if has_format {
                    worksheet.write_number_with_format(row, col, v, &fmt).map_err(xlsx_err)?;
                } else {
                    worksheet.write_number(row, col, v).map_err(xlsx_err)?;
                }
            } else if let Ok(v) = value_obj.extract::<String>() {
                if v.starts_with('=') {
                    let formula = Formula::new(&v);
                    if has_format {
                        worksheet.write_formula_with_format(row, col, formula, &fmt).map_err(xlsx_err)?;
                    } else {
                        worksheet.write_formula(row, col, formula).map_err(xlsx_err)?;
                    }
                } else if has_format {
                    worksheet.write_string_with_format(row, col, &v, &fmt).map_err(xlsx_err)?;
                } else {
                    worksheet.write_string(row, col, &v).map_err(xlsx_err)?;
                }
            }
        }

        // Column widths
        if let Ok(Some(col_widths_obj)) = sheet_dict.get_item("column_widths") {
            let col_widths: &Bound<'_, PyDict> = col_widths_obj.downcast()?;
            for (key, val) in col_widths.iter() {
                let col_idx: u16 = key.extract()?;
                let width: f64 = val.extract()?;
                worksheet.set_column_width(col_idx, width).map_err(xlsx_err)?;
            }
        }

        // Row heights
        if let Ok(Some(row_heights_obj)) = sheet_dict.get_item("row_heights") {
            let row_heights: &Bound<'_, PyDict> = row_heights_obj.downcast()?;
            for (key, val) in row_heights.iter() {
                let row_idx: u32 = key.extract()?;
                let height: f64 = val.extract()?;
                worksheet.set_row_height(row_idx, height).map_err(xlsx_err)?;
            }
        }

        // Freeze panes
        if let Ok(Some(freeze_obj)) = sheet_dict.get_item("freeze_panes") {
            let freeze: &Bound<'_, PyList> = freeze_obj.downcast()?;
            let row: u32 = freeze.get_item(0)?.extract()?;
            let col: u16 = freeze.get_item(1)?.extract()?;
            worksheet.set_freeze_panes(row, col).map_err(xlsx_err)?;
        }

        // Merged cells
        if let Ok(Some(merges_obj)) = sheet_dict.get_item("merged_cells") {
            let merges: &Bound<'_, PyList> = merges_obj.downcast()?;
            for merge_obj in merges.iter() {
                let merge: &Bound<'_, PyList> = merge_obj.downcast()?;
                let r1: u32 = merge.get_item(0)?.extract()?;
                let c1: u16 = merge.get_item(1)?.extract()?;
                let r2: u32 = merge.get_item(2)?.extract()?;
                let c2: u16 = merge.get_item(3)?.extract()?;
                worksheet.merge_range(r1, c1, r2, c2, "", &Format::new()).map_err(xlsx_err)?;
            }
        }
    }

    // Save to path or return bytes
    if let Ok(Some(path_obj)) = data.get_item("path") {
        let path: String = path_obj.extract()?;
        workbook.save(&path).map_err(xlsx_err)?;
        Ok(py.None())
    } else {
        let buf = workbook.save_to_buffer().map_err(xlsx_err)?;
        Ok(PyBytes::new(py, &buf).into())
    }
}

#[pymodule]
fn _openpyxl_rust(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(_save_workbook, m)?)?;
    Ok(())
}
```

**Step 2: Build and verify compilation**

```bash
cd /c/Users/Deren/Desktop/Files/Coding/openpyxl-rust && maturin develop --release
```

Expected: successful build

**Step 3: Commit**

```bash
git add src/lib.rs
git commit -m "feat: Rust save engine with cell types, fonts, dimensions, freeze, merge"
```

---

### Task 7: Wire Workbook.save() to Rust

**Files:**
- Modify: `python/openpyxl_rust/workbook.py`
- Modify: `python/openpyxl_rust/worksheet.py`
- Create: `tests/test_save.py`

**Step 1: Write failing test**

```python
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
```

**Step 2: Run test to verify it fails**

```bash
py -m pytest tests/test_save.py -v
```

Expected: FAIL — `NotImplementedError: Rust save not yet wired up`

**Step 3: Add _to_save_dict to Worksheet**

Add this method to the end of the `Worksheet` class in `python/openpyxl_rust/worksheet.py`:

```python
    def _to_save_dict(self):
        """Serialize worksheet data for the Rust save engine."""
        cells = []
        for (row, col), cell in self._cells.items():
            cell_data = {
                "row": row - 1,   # Rust uses 0-based
                "col": col - 1,   # Rust uses 0-based
                "value": cell.value,
            }
            if cell.font is not None:
                cell_data["font"] = {
                    "bold": cell.font.bold,
                    "italic": cell.font.italic,
                    "underline": cell.font.underline,
                    "name": cell.font.name,
                    "size": cell.font.size,
                    "color": cell.font.color,
                }
            if cell.number_format != "General":
                cell_data["number_format"] = cell.number_format
            cells.append(cell_data)

        # Column widths: convert letter key to 0-based index
        from openpyxl_rust.worksheet import _parse_cell_ref
        col_widths = {}
        for letter, dim in self.column_dimensions.items():
            if dim.width is not None:
                _, col_idx = _parse_cell_ref(f"{letter}1")
                col_widths[col_idx - 1] = dim.width

        # Row heights: convert 1-based to 0-based
        row_heights = {}
        for row_num, dim in self.row_dimensions.items():
            if dim.height is not None:
                row_heights[row_num - 1] = dim.height

        # Freeze panes
        freeze = None
        if self.freeze_panes:
            r, c = _parse_cell_ref(self.freeze_panes)
            freeze = [r - 1, c - 1]

        # Merged cells: convert refs to 0-based [r1, c1, r2, c2]
        merges = []
        for start_ref, end_ref in self.merged_cell_ranges:
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            merges.append([r1 - 1, c1 - 1, r2 - 1, c2 - 1])

        result = {
            "title": self.title,
            "cells": cells,
        }
        if col_widths:
            result["column_widths"] = col_widths
        if row_heights:
            result["row_heights"] = row_heights
        if freeze is not None:
            result["freeze_panes"] = freeze
        if merges:
            result["merged_cells"] = merges

        return result
```

**Step 4: Wire Workbook.save()**

Replace the `save` method in `python/openpyxl_rust/workbook.py`:

```python
    def save(self, filename):
        from openpyxl_rust._openpyxl_rust import _save_workbook
        from io import BytesIO

        sheets = [ws._to_save_dict() for ws in self._sheets]

        if isinstance(filename, (str, bytes, os.PathLike)):
            data = {"sheets": sheets, "path": str(filename)}
            _save_workbook(data)
        else:
            # Assume file-like object (BytesIO etc.)
            data = {"sheets": sheets}
            result_bytes = _save_workbook(data)
            filename.write(result_bytes)
```

Also add `import os` at the top of `workbook.py`.

**Step 5: Rebuild Rust module**

```bash
cd /c/Users/Deren/Desktop/Files/Coding/openpyxl-rust && maturin develop --release
```

**Step 6: Run tests**

```bash
py -m pytest tests/test_save.py -v
```

Expected: all PASS

**Step 7: Run all tests**

```bash
py -m pytest tests/ -v
```

Expected: all PASS

**Step 8: Commit**

```bash
git add python/openpyxl_rust/workbook.py python/openpyxl_rust/worksheet.py tests/test_save.py
git commit -m "feat: wire Workbook.save() to Rust engine"
```

---

### Task 8: Validation with openpyxl reading back

**Files:**
- Create: `tests/test_compat.py`

This test writes a file with openpyxl_rust, then reads it back with openpyxl to verify the output is valid.

**Step 1: Install openpyxl**

```bash
py -m pip install openpyxl
```

**Step 2: Write the test**

```python
# tests/test_compat.py
"""Read back files written by openpyxl_rust using openpyxl to verify correctness."""
import os
import tempfile

import openpyxl as real_openpyxl
from openpyxl_rust import Workbook
from openpyxl_rust.styles import Font


def _write_and_read(setup_fn):
    """Helper: create workbook, apply setup_fn, save, read back with openpyxl."""
    wb = Workbook()
    setup_fn(wb)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        wb.save(path)
        return real_openpyxl.load_workbook(path)
    finally:
        os.unlink(path)


def test_compat_string_and_number():
    def setup(wb):
        ws = wb.active
        ws["A1"] = "Hello"
        ws["B1"] = 42.5

    rb = _write_and_read(setup)
    ws = rb.active
    assert ws["A1"].value == "Hello"
    assert ws["B1"].value == 42.5


def test_compat_multiple_sheets():
    def setup(wb):
        wb.active.title = "First"
        wb.active["A1"] = "one"
        ws2 = wb.create_sheet("Second")
        ws2["A1"] = "two"

    rb = _write_and_read(setup)
    assert rb.sheetnames == ["First", "Second"]
    assert rb["First"]["A1"].value == "one"
    assert rb["Second"]["A1"].value == "two"


def test_compat_bold_font():
    def setup(wb):
        ws = wb.active
        ws["A1"] = "Bold"
        ws["A1"].font = Font(bold=True)

    rb = _write_and_read(setup)
    assert rb.active["A1"].value == "Bold"
    assert rb.active["A1"].font.bold is True


def test_compat_number_format():
    def setup(wb):
        ws = wb.active
        ws["A1"] = 1234.5
        ws["A1"].number_format = "#,##0.00"

    rb = _write_and_read(setup)
    assert rb.active["A1"].value == 1234.5
    assert rb.active["A1"].number_format == "#,##0.00"


def test_compat_boolean():
    def setup(wb):
        ws = wb.active
        ws["A1"] = True
        ws["A2"] = False

    rb = _write_and_read(setup)
    assert rb.active["A1"].value is True
    assert rb.active["A2"].value is False
```

**Step 3: Run the compat tests**

```bash
py -m pytest tests/test_compat.py -v
```

Expected: all PASS

**Step 4: Commit**

```bash
git add tests/test_compat.py
git commit -m "test: add openpyxl read-back compatibility tests"
```

---

### Task 9: Benchmark — openpyxl_rust vs openpyxl

**Files:**
- Create: `benchmarks/bench_vs_openpyxl.py`

**Step 1: Write the benchmark script**

```python
# benchmarks/bench_vs_openpyxl.py
"""
Performance comparison: openpyxl_rust vs openpyxl.
Runs identical operations with both libraries and prints a comparison table.
"""
import os
import tempfile
import time
from datetime import datetime


def bench_large_data_openpyxl(path, rows=100_000, cols=10):
    """100k rows x 10 cols of mixed types."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            if c % 3 == 0:
                ws.cell(row=r, column=c, value=f"str_{r}_{c}")
            elif c % 3 == 1:
                ws.cell(row=r, column=c, value=r * c * 1.1)
            else:
                ws.cell(row=r, column=c, value=r % 2 == 0)
    wb.save(path)


def bench_large_data_rust(path, rows=100_000, cols=10):
    """100k rows x 10 cols of mixed types."""
    from openpyxl_rust import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            if c % 3 == 0:
                ws.cell(row=r, column=c, value=f"str_{r}_{c}")
            elif c % 3 == 1:
                ws.cell(row=r, column=c, value=r * c * 1.1)
            else:
                ws.cell(row=r, column=c, value=r % 2 == 0)
    wb.save(path)


def bench_formatted_openpyxl(path, rows=10_000):
    """10k rows with bold headers, number formats, column widths."""
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ["Name", "Revenue", "Cost", "Profit", "Margin"]
    bold = Font(bold=True, size=12)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = bold
    for c in range(1, 6):
        ws.column_dimensions[chr(64 + c)].width = 15
    for r in range(2, rows + 2):
        ws.cell(row=r, column=1, value=f"Item {r}")
        ws.cell(row=r, column=2, value=r * 100.0)
        ws.cell(row=r, column=2).number_format = "$#,##0.00"
        ws.cell(row=r, column=3, value=r * 60.0)
        ws.cell(row=r, column=3).number_format = "$#,##0.00"
        ws.cell(row=r, column=4, value=r * 40.0)
        ws.cell(row=r, column=4).number_format = "$#,##0.00"
        ws.cell(row=r, column=5, value=0.4)
        ws.cell(row=r, column=5).number_format = "0.0%"
    wb.save(path)


def bench_formatted_rust(path, rows=10_000):
    """10k rows with bold headers, number formats, column widths."""
    from openpyxl_rust import Workbook
    from openpyxl_rust.styles import Font
    wb = Workbook()
    ws = wb.active
    headers = ["Name", "Revenue", "Cost", "Profit", "Margin"]
    bold = Font(bold=True, size=12)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = bold
    for c in range(1, 6):
        ws.column_dimensions[chr(64 + c)].width = 15
    for r in range(2, rows + 2):
        ws.cell(row=r, column=1, value=f"Item {r}")
        ws.cell(row=r, column=2, value=r * 100.0)
        ws.cell(row=r, column=2).number_format = "$#,##0.00"
        ws.cell(row=r, column=3, value=r * 60.0)
        ws.cell(row=r, column=3).number_format = "$#,##0.00"
        ws.cell(row=r, column=4, value=r * 40.0)
        ws.cell(row=r, column=4).number_format = "$#,##0.00"
        ws.cell(row=r, column=5, value=0.4)
        ws.cell(row=r, column=5).number_format = "0.0%"
    wb.save(path)


def bench_multisheet_openpyxl(path, sheets=5, rows=20_000):
    """5 sheets x 20k rows each."""
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for s in range(sheets):
        ws = wb.create_sheet(f"Sheet{s + 1}")
        for r in range(1, rows + 1):
            ws.cell(row=r, column=1, value=f"s{s}_r{r}")
            ws.cell(row=r, column=2, value=r * 1.5)
            ws.cell(row=r, column=3, value=r % 2 == 0)
    wb.save(path)


def bench_multisheet_rust(path, sheets=5, rows=20_000):
    """5 sheets x 20k rows each."""
    from openpyxl_rust import Workbook
    wb = Workbook()
    # Remove default sheet by just overwriting it
    wb._sheets = []
    for s in range(sheets):
        ws = wb.create_sheet(f"Sheet{s + 1}")
        for r in range(1, rows + 1):
            ws.cell(row=r, column=1, value=f"s{s}_r{r}")
            ws.cell(row=r, column=2, value=r * 1.5)
            ws.cell(row=r, column=3, value=r % 2 == 0)
    wb.save(path)


def run_bench(name, fn_openpyxl, fn_rust):
    """Run a single benchmark pair and return results."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path1 = f.name
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path2 = f.name

    try:
        # openpyxl
        start = time.perf_counter()
        fn_openpyxl(path1)
        t_openpyxl = time.perf_counter() - start
        size_openpyxl = os.path.getsize(path1)

        # openpyxl_rust
        start = time.perf_counter()
        fn_rust(path2)
        t_rust = time.perf_counter() - start
        size_rust = os.path.getsize(path2)

        speedup = t_openpyxl / t_rust if t_rust > 0 else float("inf")

        return {
            "name": name,
            "openpyxl_time": t_openpyxl,
            "rust_time": t_rust,
            "speedup": speedup,
            "openpyxl_size": size_openpyxl,
            "rust_size": size_rust,
        }
    finally:
        os.unlink(path1)
        os.unlink(path2)


def main():
    benchmarks = [
        ("Large data (100k rows x 10 cols)", bench_large_data_openpyxl, bench_large_data_rust),
        ("Formatted (10k rows, styles)", bench_formatted_openpyxl, bench_formatted_rust),
        ("Multi-sheet (5 x 20k rows)", bench_multisheet_openpyxl, bench_multisheet_rust),
    ]

    print("=" * 75)
    print("  openpyxl_rust vs openpyxl — Performance Benchmark")
    print("=" * 75)
    print()

    results = []
    for name, fn_o, fn_r in benchmarks:
        print(f"Running: {name}...")
        r = run_bench(name, fn_o, fn_r)
        results.append(r)

    print()
    print(f"{'Benchmark':<35} {'openpyxl':>10} {'ours':>10} {'speedup':>10}")
    print("-" * 67)
    for r in results:
        print(
            f"{r['name']:<35} {r['openpyxl_time']:>9.2f}s {r['rust_time']:>9.2f}s {r['speedup']:>9.1f}x"
        )

    print()
    print(f"{'Benchmark':<35} {'openpyxl':>12} {'ours':>12}")
    print("-" * 61)
    for r in results:
        print(
            f"{r['name']:<35} {r['openpyxl_size']/1024:>10.1f} KB {r['rust_size']/1024:>10.1f} KB"
        )

    print()
    avg_speedup = sum(r["speedup"] for r in results) / len(results)
    print(f"Average speedup: {avg_speedup:.1f}x faster")
    print()


if __name__ == "__main__":
    main()
```

**Step 2: Run the benchmark**

```bash
py benchmarks/bench_vs_openpyxl.py
```

Expected: table showing openpyxl_rust significantly faster than openpyxl.

**Step 3: Commit**

```bash
git add benchmarks/
git commit -m "bench: add openpyxl_rust vs openpyxl performance comparison"
```

---

### Task 10: Final integration test and cleanup

**Step 1: Run all tests**

```bash
py -m pytest tests/ -v
```

Expected: all PASS

**Step 2: Run the benchmark one final time**

```bash
py benchmarks/bench_vs_openpyxl.py
```

**Step 3: Final commit**

```bash
git add -A
git commit -m "chore: final cleanup"
```
