# Rust-Backed Cells Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Move cell storage from Python dicts to Rust, eliminating per-cell JSON serialization and the `_flush_formats` pass, targeting ~4-5x speedup over openpyxl.

**Architecture:** Python `Cell` becomes a thin proxy that writes values/formats directly to Rust via PyO3 calls. `CellFormat` is a flat Rust struct (no JSON). At save time, Rust converts `CellFormat` → `rust_xlsxwriter::Format` directly. The `_flush_formats()` method is eliminated for cell data — only sheet-level metadata (hyperlinks, comments, images, data validations, conditional formatting, protection, page setup) still gets flushed at save time.

**Tech Stack:** Rust (PyO3 0.24, rust_xlsxwriter 0.93), Python 3.12, maturin

---

### Task 1: Add `CellFormat` struct and conversion in Rust

**Files:**
- Modify: `src/lib.rs` (lines 14-46 — CellData/CellValue/SheetData, lines 96-277 — `build_format_from_json`)

**Step 1: Add `CellFormat` struct after `CellData` enum (line ~22)**

```rust
#[derive(Clone, Debug)]
struct CellFormat {
    font_bold: bool,
    font_italic: bool,
    font_name: Option<String>,
    font_size: Option<f64>,
    font_color: Option<String>,
    font_underline: Option<u8>,    // 0=none, 1=single, 2=double
    font_strikethrough: bool,
    font_vert_align: Option<u8>,   // 1=superscript, 2=subscript
    number_format: Option<String>,
    align_horizontal: Option<u8>,  // encoded as: 1=left,2=center,3=right,4=fill,5=justify,6=centerAcross,7=distributed
    align_vertical: Option<u8>,    // 1=top,2=center,3=bottom,4=justify,5=distributed
    align_wrap_text: bool,
    align_shrink_to_fit: bool,
    align_indent: u8,
    align_text_rotation: i16,
    fill_type: Option<u8>,         // 1=solid,2=darkGray,3=mediumGray,4=lightGray,5=gray125,6=gray0625
    fill_start_color: Option<String>,
    fill_end_color: Option<String>,
    border_left_style: Option<u8>,
    border_left_color: Option<String>,
    border_right_style: Option<u8>,
    border_right_color: Option<String>,
    border_top_style: Option<u8>,
    border_top_color: Option<String>,
    border_bottom_style: Option<u8>,
    border_bottom_color: Option<String>,
    border_diagonal_style: Option<u8>,
    border_diagonal_color: Option<String>,
    border_diagonal_up: bool,
    border_diagonal_down: bool,
}

impl Default for CellFormat {
    fn default() -> Self {
        CellFormat {
            font_bold: false, font_italic: false, font_name: None, font_size: None,
            font_color: None, font_underline: None, font_strikethrough: false,
            font_vert_align: None, number_format: None, align_horizontal: None,
            align_vertical: None, align_wrap_text: false, align_shrink_to_fit: false,
            align_indent: 0, align_text_rotation: 0, fill_type: None,
            fill_start_color: None, fill_end_color: None,
            border_left_style: None, border_left_color: None,
            border_right_style: None, border_right_color: None,
            border_top_style: None, border_top_color: None,
            border_bottom_style: None, border_bottom_color: None,
            border_diagonal_style: None, border_diagonal_color: None,
            border_diagonal_up: false, border_diagonal_down: false,
        }
    }
}
```

**Step 2: Update `CellValue` to use `CellFormat` instead of `format_json`**

```rust
#[derive(Clone, Debug)]
struct CellValue {
    value: CellData,
    format: CellFormat,
}
```

**Step 3: Add `cell_format_to_xlsx_format` function (replaces `build_format_from_json` for the new path)**

This function converts `CellFormat` directly to `rust_xlsxwriter::Format`. It replaces the JSON round-trip. Use the same numeric encoding for border styles as `parse_border_style_str` but via u8 lookup. Keep `build_format_from_json` for now since conditional formatting still uses JSON.

```rust
fn border_style_from_u8(v: u8) -> rust_xlsxwriter::FormatBorder {
    match v {
        1 => rust_xlsxwriter::FormatBorder::Thin,
        2 => rust_xlsxwriter::FormatBorder::Medium,
        3 => rust_xlsxwriter::FormatBorder::Thick,
        4 => rust_xlsxwriter::FormatBorder::Dashed,
        5 => rust_xlsxwriter::FormatBorder::Dotted,
        6 => rust_xlsxwriter::FormatBorder::Double,
        7 => rust_xlsxwriter::FormatBorder::Hair,
        8 => rust_xlsxwriter::FormatBorder::MediumDashed,
        9 => rust_xlsxwriter::FormatBorder::DashDot,
        10 => rust_xlsxwriter::FormatBorder::MediumDashDot,
        11 => rust_xlsxwriter::FormatBorder::DashDotDot,
        12 => rust_xlsxwriter::FormatBorder::MediumDashDotDot,
        13 => rust_xlsxwriter::FormatBorder::SlantDashDot,
        _ => rust_xlsxwriter::FormatBorder::Thin,
    }
}

fn fill_pattern_from_u8(v: u8) -> rust_xlsxwriter::FormatPattern {
    match v {
        1 => rust_xlsxwriter::FormatPattern::Solid,
        2 => rust_xlsxwriter::FormatPattern::DarkGray,
        3 => rust_xlsxwriter::FormatPattern::MediumGray,
        4 => rust_xlsxwriter::FormatPattern::LightGray,
        5 => rust_xlsxwriter::FormatPattern::Gray125,
        6 => rust_xlsxwriter::FormatPattern::Gray0625,
        _ => rust_xlsxwriter::FormatPattern::Solid,
    }
}

fn cell_format_to_xlsx_format(cf: &CellFormat) -> (Format, bool) {
    let mut fmt = Format::new();
    let mut has_any = false;

    // Font
    if cf.font_bold { fmt = fmt.set_bold(); has_any = true; }
    if cf.font_italic { fmt = fmt.set_italic(); has_any = true; }
    if let Some(ref name) = cf.font_name { fmt = fmt.set_font_name(name); has_any = true; }
    if let Some(size) = cf.font_size { fmt = fmt.set_font_size(size); has_any = true; }
    if let Some(ref color) = cf.font_color {
        if let Some(clr) = parse_color_str(color) { fmt = fmt.set_font_color(clr); has_any = true; }
    }
    if let Some(ul) = cf.font_underline {
        match ul {
            1 => { fmt = fmt.set_underline(rust_xlsxwriter::FormatUnderline::Single); has_any = true; }
            2 => { fmt = fmt.set_underline(rust_xlsxwriter::FormatUnderline::Double); has_any = true; }
            _ => {}
        }
    }
    if cf.font_strikethrough { fmt = fmt.set_font_strikethrough(); has_any = true; }
    if let Some(va) = cf.font_vert_align {
        match va {
            1 => { fmt = fmt.set_font_script(rust_xlsxwriter::FormatScript::Superscript); has_any = true; }
            2 => { fmt = fmt.set_font_script(rust_xlsxwriter::FormatScript::Subscript); has_any = true; }
            _ => {}
        }
    }

    // Number format
    if let Some(ref nf) = cf.number_format {
        if nf != "General" { fmt = fmt.set_num_format(nf); has_any = true; }
    }

    // Alignment
    if let Some(h) = cf.align_horizontal {
        let a = match h {
            1 => rust_xlsxwriter::FormatAlign::Left,
            2 => rust_xlsxwriter::FormatAlign::Center,
            3 => rust_xlsxwriter::FormatAlign::Right,
            4 => rust_xlsxwriter::FormatAlign::Fill,
            5 => rust_xlsxwriter::FormatAlign::Justify,
            6 => rust_xlsxwriter::FormatAlign::CenterAcross,
            7 => rust_xlsxwriter::FormatAlign::Distributed,
            _ => rust_xlsxwriter::FormatAlign::General,
        };
        fmt = fmt.set_align(a); has_any = true;
    }
    if let Some(v) = cf.align_vertical {
        let a = match v {
            1 => rust_xlsxwriter::FormatAlign::Top,
            2 => rust_xlsxwriter::FormatAlign::VerticalCenter,
            3 => rust_xlsxwriter::FormatAlign::Bottom,
            4 => rust_xlsxwriter::FormatAlign::VerticalJustify,
            5 => rust_xlsxwriter::FormatAlign::VerticalDistributed,
            _ => rust_xlsxwriter::FormatAlign::Bottom,
        };
        fmt = fmt.set_align(a); has_any = true;
    }
    if cf.align_wrap_text { fmt = fmt.set_text_wrap(); has_any = true; }
    if cf.align_shrink_to_fit { fmt = fmt.set_shrink(); has_any = true; }
    if cf.align_indent > 0 { fmt = fmt.set_indent(cf.align_indent); has_any = true; }
    if cf.align_text_rotation != 0 { fmt = fmt.set_rotation(cf.align_text_rotation); has_any = true; }

    // Fill
    if let Some(ft) = cf.fill_type {
        fmt = fmt.set_pattern(fill_pattern_from_u8(ft)); has_any = true;
    }
    if let Some(ref sc) = cf.fill_start_color {
        if let Some(clr) = parse_color_str(sc) { fmt = fmt.set_background_color(clr); has_any = true; }
    }
    if let Some(ref ec) = cf.fill_end_color {
        if let Some(clr) = parse_color_str(ec) { fmt = fmt.set_foreground_color(clr); has_any = true; }
    }

    // Borders
    if let Some(s) = cf.border_left_style {
        fmt = fmt.set_border_left(border_style_from_u8(s)); has_any = true;
    }
    if let Some(ref c) = cf.border_left_color {
        if let Some(clr) = parse_color_str(c) { fmt = fmt.set_border_left_color(clr); }
    }
    if let Some(s) = cf.border_right_style {
        fmt = fmt.set_border_right(border_style_from_u8(s)); has_any = true;
    }
    if let Some(ref c) = cf.border_right_color {
        if let Some(clr) = parse_color_str(c) { fmt = fmt.set_border_right_color(clr); }
    }
    if let Some(s) = cf.border_top_style {
        fmt = fmt.set_border_top(border_style_from_u8(s)); has_any = true;
    }
    if let Some(ref c) = cf.border_top_color {
        if let Some(clr) = parse_color_str(c) { fmt = fmt.set_border_top_color(clr); }
    }
    if let Some(s) = cf.border_bottom_style {
        fmt = fmt.set_border_bottom(border_style_from_u8(s)); has_any = true;
    }
    if let Some(ref c) = cf.border_bottom_color {
        if let Some(clr) = parse_color_str(c) { fmt = fmt.set_border_bottom_color(clr); }
    }
    if let Some(s) = cf.border_diagonal_style {
        fmt = fmt.set_border_diagonal(border_style_from_u8(s)); has_any = true;
        if let Some(ref c) = cf.border_diagonal_color {
            if let Some(clr) = parse_color_str(c) { fmt = fmt.set_border_diagonal_color(clr); }
        }
        // BUG FIX: only set diagonal type when at least one direction is true
        if cf.border_diagonal_up || cf.border_diagonal_down {
            let diag_type = match (cf.border_diagonal_up, cf.border_diagonal_down) {
                (true, true) => rust_xlsxwriter::FormatDiagonalBorder::BorderUpDown,
                (true, false) => rust_xlsxwriter::FormatDiagonalBorder::BorderUp,
                (false, true) => rust_xlsxwriter::FormatDiagonalBorder::BorderDown,
                _ => unreachable!(),
            };
            fmt = fmt.set_border_diagonal_type(diag_type);
        }
    }

    (fmt, has_any)
}
```

**Step 4: Update `save()` to use `cell_format_to_xlsx_format` instead of `build_format_from_json`**

Replace lines 605-666 (the cell-writing loop) with:

```rust
for (&(row, col), cv) in &sd.cells {
    let (fmt, has_format) = cell_format_to_xlsx_format(&cv.format);

    match &cv.value {
        CellData::String(s) => {
            if has_format {
                worksheet.write_string_with_format(row, col, s, &fmt).map_err(xlsx_err)?;
            } else {
                worksheet.write_string(row, col, s).map_err(xlsx_err)?;
            }
        }
        CellData::Number(n) => {
            if has_format {
                worksheet.write_number_with_format(row, col, *n, &fmt).map_err(xlsx_err)?;
            } else {
                worksheet.write_number(row, col, *n).map_err(xlsx_err)?;
            }
        }
        CellData::Boolean(b) => {
            if has_format {
                worksheet.write_boolean_with_format(row, col, *b, &fmt).map_err(xlsx_err)?;
            } else {
                worksheet.write_boolean(row, col, *b).map_err(xlsx_err)?;
            }
        }
        CellData::Formula(f) => {
            let formula = Formula::new(f);
            if has_format {
                worksheet.write_formula_with_format(row, col, formula, &fmt).map_err(xlsx_err)?;
            } else {
                worksheet.write_formula(row, col, formula).map_err(xlsx_err)?;
            }
        }
        CellData::DateTime(serial) => {
            if has_format {
                worksheet.write_number_with_format(row, col, *serial, &fmt).map_err(xlsx_err)?;
            } else {
                worksheet.write_number(row, col, *serial).map_err(xlsx_err)?;
            }
        }
        CellData::Empty => {
            if has_format {
                worksheet.write_blank(row, col, &fmt).map_err(xlsx_err)?;
            }
        }
    }
}
```

**Step 5: Simplify `CellData::DateTime` to just `DateTime(f64)`**

Remove the dead `is_date_only` field. Change the enum variant from `DateTime { serial: f64, is_date_only: bool }` to `DateTime(f64)`.

**Step 6: Update all existing `RustWorkbook` methods that create/modify `CellValue`**

Every method that does `sd.cells.insert(key, CellValue { value: ..., format_json: ... })` must change to `CellValue { value: ..., format: CellFormat::default() }`. And `format_json` access becomes `format` access.

This affects: `set_cell_string`, `set_cell_number`, `set_cell_boolean`, `set_cell_datetime`, `set_cell_empty`, `set_rows_batch`.

**Step 7: Add new PyO3 methods for direct format setting**

```rust
/// Set a cell value with automatic type detection (replaces per-type Python dispatch)
fn set_cell_value(&mut self, py: Python<'_>, sheet: usize, row: u32, col: u16, value: &Bound<'_, pyo3::PyAny>) -> PyResult<()> {
    let sd = self.sheets.get_mut(sheet)
        .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
    let key = (row, col);

    let cell_data = if value.is_none() {
        CellData::Empty
    } else if let Ok(b) = value.extract::<bool>() {
        CellData::Boolean(b)
    } else if let Ok(n) = value.extract::<f64>() {
        CellData::Number(n)
    } else if let Ok(s) = value.extract::<String>() {
        if s.starts_with('=') { CellData::Formula(s) } else { CellData::String(s) }
    } else {
        return Err(pyo3::exceptions::PyTypeError::new_err(
            format!("Unsupported cell value type: {}", value.get_type().name()?)
        ));
    };

    if let Some(cv) = sd.cells.get_mut(&key) {
        cv.value = cell_data;
    } else {
        sd.cells.insert(key, CellValue { value: cell_data, format: CellFormat::default() });
    }
    Ok(())
}

fn set_cell_font(&mut self, sheet: usize, row: u32, col: u16,
                 bold: bool, italic: bool, name: Option<String>, size: Option<f64>,
                 color: Option<String>, underline: Option<u8>, strikethrough: bool,
                 vert_align: Option<u8>) -> PyResult<()> {
    let sd = self.sheets.get_mut(sheet)
        .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
    let key = (row, col);
    let cv = sd.cells.entry(key).or_insert_with(|| CellValue { value: CellData::Empty, format: CellFormat::default() });
    cv.format.font_bold = bold;
    cv.format.font_italic = italic;
    cv.format.font_name = name;
    cv.format.font_size = size;
    cv.format.font_color = color;
    cv.format.font_underline = underline;
    cv.format.font_strikethrough = strikethrough;
    cv.format.font_vert_align = vert_align;
    Ok(())
}

fn set_cell_alignment(&mut self, sheet: usize, row: u32, col: u16,
                      horizontal: Option<u8>, vertical: Option<u8>,
                      wrap_text: bool, shrink_to_fit: bool,
                      indent: u8, text_rotation: i16) -> PyResult<()> {
    let sd = self.sheets.get_mut(sheet)
        .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
    let key = (row, col);
    let cv = sd.cells.entry(key).or_insert_with(|| CellValue { value: CellData::Empty, format: CellFormat::default() });
    cv.format.align_horizontal = horizontal;
    cv.format.align_vertical = vertical;
    cv.format.align_wrap_text = wrap_text;
    cv.format.align_shrink_to_fit = shrink_to_fit;
    cv.format.align_indent = indent;
    cv.format.align_text_rotation = text_rotation;
    Ok(())
}

fn set_cell_fill(&mut self, sheet: usize, row: u32, col: u16,
                 fill_type: Option<u8>, start_color: Option<String>,
                 end_color: Option<String>) -> PyResult<()> {
    let sd = self.sheets.get_mut(sheet)
        .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
    let key = (row, col);
    let cv = sd.cells.entry(key).or_insert_with(|| CellValue { value: CellData::Empty, format: CellFormat::default() });
    cv.format.fill_type = fill_type;
    cv.format.fill_start_color = start_color;
    cv.format.fill_end_color = end_color;
    Ok(())
}

fn set_cell_border(&mut self, sheet: usize, row: u32, col: u16,
                   left_style: Option<u8>, left_color: Option<String>,
                   right_style: Option<u8>, right_color: Option<String>,
                   top_style: Option<u8>, top_color: Option<String>,
                   bottom_style: Option<u8>, bottom_color: Option<String>,
                   diag_style: Option<u8>, diag_color: Option<String>,
                   diag_up: bool, diag_down: bool) -> PyResult<()> {
    let sd = self.sheets.get_mut(sheet)
        .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
    let key = (row, col);
    let cv = sd.cells.entry(key).or_insert_with(|| CellValue { value: CellData::Empty, format: CellFormat::default() });
    cv.format.border_left_style = left_style;
    cv.format.border_left_color = left_color;
    cv.format.border_right_style = right_style;
    cv.format.border_right_color = right_color;
    cv.format.border_top_style = top_style;
    cv.format.border_top_color = top_color;
    cv.format.border_bottom_style = bottom_style;
    cv.format.border_bottom_color = bottom_color;
    cv.format.border_diagonal_style = diag_style;
    cv.format.border_diagonal_color = diag_color;
    cv.format.border_diagonal_up = diag_up;
    cv.format.border_diagonal_down = diag_down;
    Ok(())
}

fn set_cell_number_format(&mut self, sheet: usize, row: u32, col: u16, format: String) -> PyResult<()> {
    let sd = self.sheets.get_mut(sheet)
        .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
    let key = (row, col);
    let cv = sd.cells.entry(key).or_insert_with(|| CellValue { value: CellData::Empty, format: CellFormat::default() });
    cv.format.number_format = Some(format);
    Ok(())
}
```

**Step 8: Remove `set_cell_format`, `set_cell_formats_batch`**

These JSON-based methods are no longer needed for cell formatting. Keep them temporarily if you want a fallback, but they should be removed once the Python side is updated.

**Step 9: Build and verify compilation**

Run: `.venv/Scripts/python -m maturin develop --release`
Expected: Compiles with no errors (warnings about unused `set_cell_format`/`set_cell_formats_batch` are ok for now)

**Step 10: Commit**

```bash
git add src/lib.rs
git commit -m "feat: add CellFormat struct and direct format setters in Rust

Replace JSON-based format serialization with structured CellFormat.
Add set_cell_value, set_cell_font, set_cell_alignment, set_cell_fill,
set_cell_border, set_cell_number_format PyO3 methods.
Fix diagonal border (false,false) defaulting to BorderUp.
Remove dead is_date_only field from CellData::DateTime."
```

---

### Task 2: Update Python `Cell` to proxy through Rust

**Files:**
- Modify: `python/openpyxl_rust/cell.py`

**Step 1: Add encoding helpers at the top of `cell.py`**

```python
# Border style string -> u8 encoding (matches Rust border_style_from_u8)
_BORDER_STYLE_MAP = {
    "thin": 1, "medium": 2, "thick": 3, "dashed": 4, "dotted": 5,
    "double": 6, "hair": 7, "mediumDashed": 8, "dashDot": 9,
    "mediumDashDot": 10, "dashDotDot": 11, "mediumDashDotDot": 12,
    "slantDashDot": 13,
}

_FILL_TYPE_MAP = {
    "solid": 1, "darkGray": 2, "mediumGray": 3, "lightGray": 4,
    "gray125": 5, "gray0625": 6,
}

_HALIGN_MAP = {
    "left": 1, "center": 2, "right": 3, "fill": 4, "justify": 5,
    "centerContinuous": 6, "center_continuous": 6, "distributed": 7,
}

_VALIGN_MAP = {
    "top": 1, "center": 2, "bottom": 3, "justify": 4, "distributed": 5,
}

def _underline_to_u8(val):
    if val is None: return None
    if val == "single": return 1
    if val == "double": return 2
    return None

def _vert_align_to_u8(val):
    if val is None: return None
    if val == "superscript": return 1
    if val == "subscript": return 2
    return None
```

**Step 2: Rewrite `Cell` class with `__slots__` and Rust proxy**

```python
class Cell:
    __slots__ = ('_wb', '_sheet_idx', 'row', 'column',
                 '_font', '_alignment', '_border', '_fill',
                 '_number_format', '_hyperlink', '_comment', '_value')

    # Type constants matching openpyxl
    TYPE_STRING = 's'
    TYPE_FORMULA = 'f'
    TYPE_NUMERIC = 'n'
    TYPE_BOOL = 'b'
    TYPE_NULL = 'n'
    TYPE_INLINE = 's'
    TYPE_ERROR = 'e'
    TYPE_FORMULA_CACHE_STRING = 's'

    def __init__(self, wb=None, sheet_idx=None, row=1, column=1, value=None):
        self._wb = wb
        self._sheet_idx = sheet_idx
        self.row = row
        self.column = column
        self._font = None
        self._alignment = None
        self._border = None
        self._fill = None
        self._number_format = "General"
        self._hyperlink = None
        self._comment = None
        self._value = None
        if value is not None:
            self.value = value

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, val):
        from datetime import datetime, date, time
        self._value = val
        if self._wb is None:
            return
        r, c = self.row - 1, self.column - 1
        # datetime before date (datetime is subclass of date)
        if isinstance(val, datetime):
            serial = _date_to_excel_serial(val.year, val.month, val.day)
            serial += (val.hour * 3600 + val.minute * 60 + val.second + val.microsecond / 1_000_000) / 86400.0
            self._wb.set_cell_datetime(self._sheet_idx, r, c, serial)
            if self._number_format == "General":
                self.number_format = "yyyy-mm-dd hh:mm:ss"
        elif isinstance(val, date):
            serial = _date_to_excel_serial(val.year, val.month, val.day)
            self._wb.set_cell_datetime(self._sheet_idx, r, c, serial)
            if self._number_format == "General":
                self.number_format = "yyyy-mm-dd"
        elif isinstance(val, time):
            serial = (val.hour * 3600 + val.minute * 60 + val.second + val.microsecond / 1_000_000) / 86400.0
            self._wb.set_cell_datetime(self._sheet_idx, r, c, serial)
            if self._number_format == "General":
                self.number_format = "hh:mm:ss"
        elif isinstance(val, bool):
            self._wb.set_cell_boolean(self._sheet_idx, r, c, val)
        elif isinstance(val, (int, float)):
            self._wb.set_cell_number(self._sheet_idx, r, c, float(val))
        elif isinstance(val, str):
            self._wb.set_cell_string(self._sheet_idx, r, c, val)
        elif val is None:
            self._wb.set_cell_empty(self._sheet_idx, r, c)

    @property
    def font(self):
        return self._font

    @font.setter
    def font(self, f):
        self._font = f
        if self._wb is None or f is None:
            return
        self._wb.set_cell_font(
            self._sheet_idx, self.row - 1, self.column - 1,
            f.bold, f.italic, f.name, f.size, f.color,
            _underline_to_u8(f.underline), f.strikethrough,
            _vert_align_to_u8(f.vertAlign))

    @property
    def alignment(self):
        return self._alignment

    @alignment.setter
    def alignment(self, a):
        self._alignment = a
        if self._wb is None or a is None:
            return
        self._wb.set_cell_alignment(
            self._sheet_idx, self.row - 1, self.column - 1,
            _HALIGN_MAP.get(a.horizontal) if a.horizontal else None,
            _VALIGN_MAP.get(a.vertical) if a.vertical else None,
            a.wrap_text, a.shrink_to_fit, a.indent, a.text_rotation)

    @property
    def border(self):
        return self._border

    @border.setter
    def border(self, b):
        self._border = b
        if self._wb is None or b is None:
            return
        self._wb.set_cell_border(
            self._sheet_idx, self.row - 1, self.column - 1,
            _BORDER_STYLE_MAP.get(b.left.style) if b.left.style else None, b.left.color,
            _BORDER_STYLE_MAP.get(b.right.style) if b.right.style else None, b.right.color,
            _BORDER_STYLE_MAP.get(b.top.style) if b.top.style else None, b.top.color,
            _BORDER_STYLE_MAP.get(b.bottom.style) if b.bottom.style else None, b.bottom.color,
            _BORDER_STYLE_MAP.get(b.diagonal.style) if b.diagonal.style else None, b.diagonal.color,
            b.diagonalUp, b.diagonalDown)

    @property
    def fill(self):
        return self._fill

    @fill.setter
    def fill(self, f):
        self._fill = f
        if self._wb is None or f is None:
            return
        self._wb.set_cell_fill(
            self._sheet_idx, self.row - 1, self.column - 1,
            _FILL_TYPE_MAP.get(f.fill_type) if f.fill_type else None,
            f.start_color, f.end_color)

    @property
    def number_format(self):
        return self._number_format

    @number_format.setter
    def number_format(self, nf):
        self._number_format = nf
        if self._wb is None:
            return
        if nf != "General":
            self._wb.set_cell_number_format(self._sheet_idx, self.row - 1, self.column - 1, nf)

    @property
    def hyperlink(self):
        return self._hyperlink

    @hyperlink.setter
    def hyperlink(self, val):
        self._hyperlink = val

    @property
    def comment(self):
        return self._comment

    @comment.setter
    def comment(self, val):
        self._comment = val

    @property
    def coordinate(self):
        return f"{_col_letter(self.column)}{self.row}"

    @property
    def data_type(self):
        from datetime import datetime, date, time
        v = self._value
        if v is None: return 'n'
        if isinstance(v, bool): return 'b'
        if isinstance(v, (int, float)): return 'n'
        if isinstance(v, (datetime, date, time)): return 'd'
        if isinstance(v, str):
            return 'f' if v.startswith('=') else 's'
        return 'n'
```

Note: `_date_to_excel_serial` needs to be importable from `cell.py` — it lives in `worksheet.py`. Move it to `cell.py` (or a shared utils) in this step.

**Step 3: Commit**

```bash
git add python/openpyxl_rust/cell.py
git commit -m "feat: rewrite Cell as Rust-backed proxy with __slots__

Cell now pushes values and formats directly to Rust on set.
Eliminates deferred _flush_formats for cell data.
Fix: include microsecond precision in datetime serial conversion."
```

---

### Task 3: Update `Worksheet` to use new Cell proxy

**Files:**
- Modify: `python/openpyxl_rust/worksheet.py`

**Step 1: Remove `_date_to_excel_serial` from worksheet.py (moved to cell.py in Task 2)**

**Step 2: Remove `_set_rust_value` method entirely**

This is replaced by `Cell.value` setter.

**Step 3: Update `__init__` — rename `_cells` to `_cell_cache`**

```python
def __init__(self, title="Sheet", workbook=None, sheet_idx=None):
    # ... keep all existing fields ...
    self._cell_cache = {}  # (row, col) -> Cell proxy
    # ... rest stays same ...
```

**Step 4: Update `cell()` to create proxy Cells with Rust reference**

```python
def cell(self, row, column, value=None):
    key = (row, column)
    if key not in self._cell_cache:
        self._cell_cache[key] = Cell(
            wb=self._workbook._rust_wb if self._workbook else None,
            sheet_idx=self._sheet_idx,
            row=row, column=column)
    c = self._cell_cache[key]
    if value is not None:
        c.value = value
    return c
```

**Step 5: Update `__setitem__` and `__getitem__`**

`__setitem__` uses `cell()` instead of creating raw Cells:
```python
def __setitem__(self, key, value):
    row, col = _parse_cell_ref(key)
    self.cell(row=row, column=col, value=value)
```

`__getitem__` uses `cell()`:
```python
def __getitem__(self, key):
    if isinstance(key, int):
        min_col = self.min_column or 1
        max_col = self.max_column or 1
        return tuple(self.cell(row=key, column=col) for col in range(min_col, max_col + 1))
    if ':' in key:
        start_ref, end_ref = key.split(':')
        r1, c1 = _parse_cell_ref(start_ref)
        r2, c2 = _parse_cell_ref(end_ref)
        return tuple(
            tuple(self.cell(row=row, column=col) for col in range(c1, c2 + 1))
            for row in range(r1, r2 + 1)
        )
    row, col = _parse_cell_ref(key)
    return self.cell(row=row, column=col)
```

**Step 6: Update `append()` to create proxy Cells**

```python
def append(self, iterable):
    row = self._next_row()
    for col_idx, value in enumerate(iterable, start=1):
        self.cell(row=row, column=col_idx, value=value)
    self._current_row = row
```

**Step 7: Update `append_rows()` — keep the Rust batch path, store lightweight cache entries**

```python
def append_rows(self, rows_data):
    from datetime import datetime, date, time
    from openpyxl_rust.cell import _date_to_excel_serial
    start_row = self._next_row()
    start_row_0based = start_row - 1

    if self._workbook is not None and self._sheet_idx is not None:
        rows_list = []
        for row in rows_data:
            converted_row = []
            for value in row:
                if isinstance(value, datetime):
                    serial = _date_to_excel_serial(value.year, value.month, value.day)
                    serial += (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86400.0
                    converted_row.append(serial)
                elif isinstance(value, date):
                    serial = _date_to_excel_serial(value.year, value.month, value.day)
                    converted_row.append(serial)
                elif isinstance(value, time):
                    serial = (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86400.0
                    converted_row.append(serial)
                else:
                    converted_row.append(value)
            rows_list.append(converted_row)
        self._workbook._rust_wb.set_rows_batch(self._sheet_idx, start_row_0based, rows_list)

    # Store lightweight cache entries for Python-side access
    # Also handle datetime number formats
    for row_offset, row_values in enumerate(rows_data):
        row = start_row + row_offset
        for col_idx, value in enumerate(row_values, start=1):
            c = Cell(wb=None, sheet_idx=None, row=row, column=col_idx)
            c._value = value
            # Set datetime number format in Rust
            if isinstance(value, datetime):
                if self._workbook:
                    self._workbook._rust_wb.set_cell_number_format(
                        self._sheet_idx, row - 1, col_idx - 1, "yyyy-mm-dd hh:mm:ss")
            elif isinstance(value, date):
                if self._workbook:
                    self._workbook._rust_wb.set_cell_number_format(
                        self._sheet_idx, row - 1, col_idx - 1, "yyyy-mm-dd")
            elif isinstance(value, time):
                if self._workbook:
                    self._workbook._rust_wb.set_cell_number_format(
                        self._sheet_idx, row - 1, col_idx - 1, "hh:mm:ss")
            self._cell_cache[(row, col_idx)] = c
        self._current_row = row
```

**Step 8: Update `min_row`, `max_row`, `min_column`, `max_column`, `dimensions` to use `_cell_cache`**

Replace all `self._cells` references with `self._cell_cache`.

**Step 9: Update `iter_rows`, `iter_cols`, `values` to use `_cell_cache`**

Replace `self._cells` with `self._cell_cache` and `self._cells.get((row, col))` with `self._cell_cache.get((row, col))`.

**Step 10: Update `_next_row`**

Replace `self._cells` with `self._cell_cache`.

**Step 11: Update `insert_rows`, `delete_rows`, `insert_cols`, `delete_cols`**

Replace `self._cells` with `self._cell_cache`. The `_resync_rust` method needs updating too — it now re-pushes values from the cache Cells:

```python
def _resync_rust(self):
    if self._workbook is None or self._sheet_idx is None:
        return
    wb = self._workbook._rust_wb
    idx = self._sheet_idx
    wb.clear_cells(idx)
    wb.clear_merge_ranges(idx)
    for (row, col), cell in self._cell_cache.items():
        if cell._value is not None:
            # Re-push value to Rust
            old_wb = cell._wb
            cell._wb = wb
            cell._sheet_idx = idx
            cell.value = cell._value  # triggers Rust push
            # Re-push format if set
            if cell._font is not None:
                cell.font = cell._font
            if cell._alignment is not None:
                cell.alignment = cell._alignment
            if cell._border is not None:
                cell.border = cell._border
            if cell._fill is not None:
                cell.fill = cell._fill
            if cell._number_format != "General":
                cell.number_format = cell._number_format
    for (start_ref, end_ref) in self.merged_cell_ranges:
        r1, c1 = _parse_cell_ref(start_ref)
        r2, c2 = _parse_cell_ref(end_ref)
        wb.add_merge_range(idx, r1 - 1, c1 - 1, r2 - 1, c2 - 1)
```

**Step 12: Rewrite `_flush_formats` to only handle sheet-level metadata**

Rename to `_flush_metadata`. Remove all cell format logic. Keep only: column widths, row heights, freeze panes, hyperlinks, comments, autofilter, protection, page setup, data validations, conditional formatting, images.

```python
def _flush_metadata(self):
    """Push sheet-level metadata to Rust. Called right before save.
    Cell values and formats are already in Rust via proxy setters."""
    if self._workbook is None or self._sheet_idx is None:
        return
    wb = self._workbook._rust_wb
    idx = self._sheet_idx

    # Column widths
    for letter, dim in self.column_dimensions.items():
        if dim.width is not None:
            _, col_idx = _parse_cell_ref(f"{letter}1")
            wb.set_column_width(idx, col_idx - 1, dim.width)

    # Row heights
    for row_num, dim in self.row_dimensions.items():
        if dim.height is not None:
            wb.set_row_height(idx, row_num - 1, dim.height)

    # Freeze panes
    if self.freeze_panes:
        r, c = _parse_cell_ref(self.freeze_panes)
        wb.set_freeze_panes(idx, r - 1, c - 1)

    # Hyperlinks
    for (row, col), cell in self._cell_cache.items():
        if cell._hyperlink is not None:
            url = cell._hyperlink
            if isinstance(url, str) and url.startswith("#"):
                url = "internal:" + url[1:]
            wb.add_hyperlink(idx, row - 1, col - 1, url, None, None)

    # Comments/Notes
    for (row, col), cell in self._cell_cache.items():
        if cell._comment is not None:
            author = cell._comment.author if cell._comment.author else None
            wb.add_note(idx, row - 1, col - 1, cell._comment.text, author)

    # Autofilter, protection, page setup, data validations,
    # conditional formatting, images — KEEP AS-IS (copy from existing _flush_formats)
    # ... (these all still use JSON for complex structures, which is fine since
    #      they are called once per sheet, not per cell)
```

**Step 13: Update `workbook.py` save() to call `_flush_metadata` instead of `_flush_formats`**

```python
def save(self, filename):
    for ws in self._sheets:
        ws._flush_metadata()
    # ... rest stays same
```

**Step 14: Commit**

```bash
git add python/openpyxl_rust/cell.py python/openpyxl_rust/worksheet.py python/openpyxl_rust/workbook.py
git commit -m "feat: wire Python Cell/Worksheet to Rust-backed storage

Cell is now a thin proxy with __slots__. Values and formats push to Rust
immediately on set. _flush_formats replaced with _flush_metadata (sheet-level only).
Eliminates per-cell JSON serialization at save time."
```

---

### Task 4: Update `set_cell_datetime` Rust method (simplify)

**Files:**
- Modify: `src/lib.rs`

**Step 1: Change `set_cell_datetime` signature — remove `is_date_only` parameter**

```rust
fn set_cell_datetime(&mut self, sheet: usize, row: u32, col: u16, serial: f64) -> PyResult<()> {
    let sd = self.sheets.get_mut(sheet)
        .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
    let key = (row, col);
    if let Some(cv) = sd.cells.get_mut(&key) {
        cv.value = CellData::DateTime(serial);
    } else {
        sd.cells.insert(key, CellValue { value: CellData::DateTime(serial), format: CellFormat::default() });
    }
    Ok(())
}
```

**Step 2: Fix sheet protection ordering**

Swap lines 766-769 so `protect_with_options` is called first, then `protect_with_password`:

```rust
worksheet.protect_with_options(&opts);
if let Some(pw) = password {
    worksheet.protect_with_password(pw);
}
```

**Step 3: Remove old `set_cell_format` and `set_cell_formats_batch` methods**

Delete these methods — they used the old `format_json` field which no longer exists.

**Step 4: Build and verify**

Run: `.venv/Scripts/python -m maturin develop --release`
Expected: Clean compilation, no warnings about unused fields.

**Step 5: Commit**

```bash
git add src/lib.rs
git commit -m "fix: simplify set_cell_datetime, fix protection order, remove JSON format methods"
```

---

### Task 5: Update `__init__.py` load_workbook to use new Cell constructor

**Files:**
- Modify: `python/openpyxl_rust/__init__.py`

**Step 1: The `load_workbook` function creates cells via `ws.cell()` which now uses the proxy**

No code changes needed — `ws.cell(row=..., column=..., value=...)` already routes through the updated `Worksheet.cell()` which creates a proxy Cell. Just verify it works.

**Step 2: Commit (if any changes were needed)**

---

### Task 6: Build, run full test suite, fix failures

**Step 1: Build**

Run: `.venv/Scripts/python -m maturin develop --release`

**Step 2: Run tests**

Run: `.venv/Scripts/python -m pytest tests/ -v`

**Step 3: Fix any failures**

Common expected issues:
- Tests accessing `cell.value` might see different behavior since `Cell._value` stores the original Python object while Rust stores the converted value
- Tests that directly reference `ws._cells` need to use `ws._cell_cache`
- Some tests may create bare `Cell()` without workbook reference — ensure the `wb=None` path still works

**Step 4: Commit fixes**

```bash
git add -A
git commit -m "fix: resolve test failures from Rust-backed cell migration"
```

---

### Task 7: Run benchmarks and verify speedup

**Step 1: Run benchmark**

Run: `.venv/Scripts/python benchmarks/bench_vs_openpyxl.py`
Expected: ~4-5x speedup vs openpyxl (down from ~3s to ~1.4-1.6s for 100k rows)

**Step 2: Run profiling breakdown**

Run the same profiling script from the analysis to verify time distribution has shifted.

**Step 3: Update README performance table if numbers changed significantly**

**Step 4: Commit**

```bash
git add benchmarks/ README.md
git commit -m "docs: update benchmark numbers after Rust-backed cell optimization"
```

---

### Task 8: Final verification

**Step 1: Run full test suite one more time**

Run: `.venv/Scripts/python -m pytest tests/ -v`
Expected: All tests pass (except potentially `test_image_multiple` which was already failing before)

**Step 2: Verify openpyxl compatibility with a round-trip test**

```python
# Quick smoke test: write with openpyxl_rust, read back with openpyxl
from openpyxl_rust import Workbook
from openpyxl_rust.styles import Font, Alignment, Border, Side, PatternFill
import openpyxl

wb = Workbook()
ws = wb.active
ws["A1"] = "hello"
ws["A1"].font = Font(bold=True, size=14, color="FF0000")
ws["B1"] = 42.5
ws["B1"].number_format = "$#,##0.00"
ws["C1"] = True
ws.cell(row=2, column=1, value="world")
ws.cell(row=2, column=1).alignment = Alignment(horizontal="center")
ws.cell(row=2, column=2, value=100)
ws.cell(row=2, column=2).fill = PatternFill(fill_type="solid", start_color="FFFF00")
ws.cell(row=2, column=3, value=200)
ws.cell(row=2, column=3).border = Border(left=Side(style="thin"), right=Side(style="medium", color="0000FF"))
wb.save("_test_compat.xlsx")

# Read back with openpyxl
wb2 = openpyxl.load_workbook("_test_compat.xlsx")
ws2 = wb2.active
assert ws2["A1"].value == "hello"
assert ws2["A1"].font.bold == True
assert ws2["B1"].value == 42.5
assert ws2["C1"].value == True
assert ws2["A2"].value == "world"
print("Round-trip compatibility OK")
```

**Step 3: Clean up temp files, final commit**

```bash
git add -A
git commit -m "feat: v0.6.0 Rust-backed cells for ~4-5x speedup"
```
