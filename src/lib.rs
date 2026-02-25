// src/lib.rs
use pyo3::prelude::*;
use pyo3::types::{PyAnyMethods, PyBytes, PyDict, PyList, PyTuple};
use rust_xlsxwriter::{Format, Formula, Workbook, XlsxError};
use calamine::{open_workbook, Reader, Xlsx, Data};

use std::collections::HashMap;
use std::io::Cursor;

// =====================================================================
// Cell data types for RustWorkbook storage
// =====================================================================

#[derive(Clone, Debug)]
enum CellData {
    String(String),
    Number(f64),
    Boolean(bool),
    Formula(String),
    DateTime { serial: f64, is_date_only: bool },
    Empty,
}

#[derive(Clone, Debug)]
struct CellValue {
    value: CellData,
    format_json: Option<String>,
}

#[derive(Clone, Debug)]
struct SheetData {
    title: String,
    cells: HashMap<(u32, u16), CellValue>,        // (row, col) -> CellValue (0-based)
    column_widths: HashMap<u16, f64>,               // col (0-based) -> width
    row_heights: HashMap<u32, f64>,                 // row (0-based) -> height
    freeze_panes: Option<(u32, u16)>,               // (row, col) 0-based
    merged_ranges: Vec<(u32, u16, u32, u16)>,       // (r1, c1, r2, c2) 0-based
    hyperlinks: Vec<(u32, u16, String, Option<String>, Option<String>)>, // (row, col, url, text, tooltip)
    notes: Vec<(u32, u16, String, Option<String>)>,                     // (row, col, text, author)
    autofilter: Option<(u32, u16, u32, u16)>,                          // (r1, c1, r2, c2) 0-based
    protection_json: Option<String>,
    page_setup_json: Option<String>,
    images: Vec<(u32, u16, Vec<u8>, Option<f64>, Option<f64>)>,  // (row, col, image_data, scale_width, scale_height)
    data_validations: Vec<String>,
    conditional_formats: Vec<String>,
}

impl SheetData {
    fn new(title: String) -> Self {
        SheetData {
            title,
            cells: HashMap::new(),
            column_widths: HashMap::new(),
            row_heights: HashMap::new(),
            freeze_panes: None,
            merged_ranges: Vec::new(),
            hyperlinks: Vec::new(),
            notes: Vec::new(),
            autofilter: None,
            protection_json: None,
            page_setup_json: None,
            images: Vec::new(),
            data_validations: Vec::new(),
            conditional_formats: Vec::new(),
        }
    }
}

// =====================================================================
// build_format_from_json: parse JSON string into rust_xlsxwriter::Format
// =====================================================================

fn parse_border_style_str(s: &str) -> rust_xlsxwriter::FormatBorder {
    match s {
        "thin" => rust_xlsxwriter::FormatBorder::Thin,
        "medium" => rust_xlsxwriter::FormatBorder::Medium,
        "thick" => rust_xlsxwriter::FormatBorder::Thick,
        "dashed" => rust_xlsxwriter::FormatBorder::Dashed,
        "dotted" => rust_xlsxwriter::FormatBorder::Dotted,
        "double" => rust_xlsxwriter::FormatBorder::Double,
        "hair" => rust_xlsxwriter::FormatBorder::Hair,
        "mediumDashed" => rust_xlsxwriter::FormatBorder::MediumDashed,
        "dashDot" => rust_xlsxwriter::FormatBorder::DashDot,
        "mediumDashDot" => rust_xlsxwriter::FormatBorder::MediumDashDot,
        "dashDotDot" => rust_xlsxwriter::FormatBorder::DashDotDot,
        "mediumDashDotDot" => rust_xlsxwriter::FormatBorder::MediumDashDotDot,
        "slantDashDot" => rust_xlsxwriter::FormatBorder::SlantDashDot,
        _ => rust_xlsxwriter::FormatBorder::Thin,
    }
}

fn parse_color_str(c: &str) -> Option<rust_xlsxwriter::Color> {
    u32::from_str_radix(c, 16).ok().map(rust_xlsxwriter::Color::from)
}

fn build_format_from_json(json_str: &str) -> Result<Format, String> {
    let val: serde_json::Value = serde_json::from_str(json_str)
        .map_err(|e| format!("JSON parse error: {}", e))?;
    let obj = val.as_object().ok_or("Expected JSON object")?;
    let mut fmt = Format::new();

    // Font
    if let Some(font) = obj.get("font").and_then(|v| v.as_object()) {
        if let Some(bold) = font.get("bold").and_then(|v| v.as_bool()) {
            if bold { fmt = fmt.set_bold(); }
        }
        if let Some(italic) = font.get("italic").and_then(|v| v.as_bool()) {
            if italic { fmt = fmt.set_italic(); }
        }
        if let Some(ul) = font.get("underline").and_then(|v| v.as_str()) {
            if ul == "single" {
                fmt = fmt.set_underline(rust_xlsxwriter::FormatUnderline::Single);
            } else if ul == "double" {
                fmt = fmt.set_underline(rust_xlsxwriter::FormatUnderline::Double);
            }
        }
        if let Some(name) = font.get("name").and_then(|v| v.as_str()) {
            fmt = fmt.set_font_name(name);
        }
        if let Some(size) = font.get("size").and_then(|v| v.as_f64()) {
            fmt = fmt.set_font_size(size);
        }
        if let Some(color) = font.get("color").and_then(|v| v.as_str()) {
            if let Some(clr) = parse_color_str(color) {
                fmt = fmt.set_font_color(clr);
            }
        }
        if let Some(st) = font.get("strikethrough").and_then(|v| v.as_bool()) {
            if st { fmt = fmt.set_font_strikethrough(); }
        }
        if let Some(va) = font.get("vertAlign").and_then(|v| v.as_str()) {
            match va {
                "superscript" => { fmt = fmt.set_font_script(rust_xlsxwriter::FormatScript::Superscript); }
                "subscript" => { fmt = fmt.set_font_script(rust_xlsxwriter::FormatScript::Subscript); }
                _ => {}
            }
        }
    }

    // Alignment
    if let Some(align) = obj.get("alignment").and_then(|v| v.as_object()) {
        if let Some(h) = align.get("horizontal").and_then(|v| v.as_str()) {
            let a = match h {
                "center" => rust_xlsxwriter::FormatAlign::Center,
                "right" => rust_xlsxwriter::FormatAlign::Right,
                "left" => rust_xlsxwriter::FormatAlign::Left,
                "fill" => rust_xlsxwriter::FormatAlign::Fill,
                "justify" => rust_xlsxwriter::FormatAlign::Justify,
                "centerContinuous" | "center_continuous" => rust_xlsxwriter::FormatAlign::CenterAcross,
                "distributed" => rust_xlsxwriter::FormatAlign::Distributed,
                _ => rust_xlsxwriter::FormatAlign::General,
            };
            fmt = fmt.set_align(a);
        }
        if let Some(v) = align.get("vertical").and_then(|v| v.as_str()) {
            let a = match v {
                "center" => rust_xlsxwriter::FormatAlign::VerticalCenter,
                "top" => rust_xlsxwriter::FormatAlign::Top,
                "bottom" => rust_xlsxwriter::FormatAlign::Bottom,
                "justify" => rust_xlsxwriter::FormatAlign::VerticalJustify,
                "distributed" => rust_xlsxwriter::FormatAlign::VerticalDistributed,
                _ => rust_xlsxwriter::FormatAlign::Bottom,
            };
            fmt = fmt.set_align(a);
        }
        if let Some(wt) = align.get("wrap_text").and_then(|v| v.as_bool()) {
            if wt { fmt = fmt.set_text_wrap(); }
        }
        if let Some(sf) = align.get("shrink_to_fit").and_then(|v| v.as_bool()) {
            if sf { fmt = fmt.set_shrink(); }
        }
        if let Some(indent) = align.get("indent").and_then(|v| v.as_u64()) {
            if indent > 0 { fmt = fmt.set_indent(indent as u8); }
        }
        if let Some(rot) = align.get("text_rotation").and_then(|v| v.as_i64()) {
            if rot != 0 { fmt = fmt.set_rotation(rot as i16); }
        }
    }

    // Border
    if let Some(border) = obj.get("border").and_then(|v| v.as_object()) {
        if let Some(left) = border.get("left").and_then(|v| v.as_object()) {
            if let Some(style) = left.get("style").and_then(|v| v.as_str()) {
                fmt = fmt.set_border_left(parse_border_style_str(style));
            }
            if let Some(color) = left.get("color").and_then(|v| v.as_str()) {
                if let Some(clr) = parse_color_str(color) {
                    fmt = fmt.set_border_left_color(clr);
                }
            }
        }
        if let Some(right) = border.get("right").and_then(|v| v.as_object()) {
            if let Some(style) = right.get("style").and_then(|v| v.as_str()) {
                fmt = fmt.set_border_right(parse_border_style_str(style));
            }
            if let Some(color) = right.get("color").and_then(|v| v.as_str()) {
                if let Some(clr) = parse_color_str(color) {
                    fmt = fmt.set_border_right_color(clr);
                }
            }
        }
        if let Some(top) = border.get("top").and_then(|v| v.as_object()) {
            if let Some(style) = top.get("style").and_then(|v| v.as_str()) {
                fmt = fmt.set_border_top(parse_border_style_str(style));
            }
            if let Some(color) = top.get("color").and_then(|v| v.as_str()) {
                if let Some(clr) = parse_color_str(color) {
                    fmt = fmt.set_border_top_color(clr);
                }
            }
        }
        if let Some(bottom) = border.get("bottom").and_then(|v| v.as_object()) {
            if let Some(style) = bottom.get("style").and_then(|v| v.as_str()) {
                fmt = fmt.set_border_bottom(parse_border_style_str(style));
            }
            if let Some(color) = bottom.get("color").and_then(|v| v.as_str()) {
                if let Some(clr) = parse_color_str(color) {
                    fmt = fmt.set_border_bottom_color(clr);
                }
            }
        }
        if let Some(diag) = border.get("diagonal").and_then(|v| v.as_object()) {
            if let Some(style) = diag.get("style").and_then(|v| v.as_str()) {
                fmt = fmt.set_border_diagonal(parse_border_style_str(style));
            }
            if let Some(color) = diag.get("color").and_then(|v| v.as_str()) {
                if let Some(clr) = parse_color_str(color) {
                    fmt = fmt.set_border_diagonal_color(clr);
                }
            }
            let diag_up = diag.get("diagonalUp").and_then(|v| v.as_bool()).unwrap_or(false);
            let diag_down = diag.get("diagonalDown").and_then(|v| v.as_bool()).unwrap_or(false);
            let diag_type = match (diag_up, diag_down) {
                (true, true) => rust_xlsxwriter::FormatDiagonalBorder::BorderUpDown,
                (true, false) => rust_xlsxwriter::FormatDiagonalBorder::BorderUp,
                (false, true) => rust_xlsxwriter::FormatDiagonalBorder::BorderDown,
                (false, false) => rust_xlsxwriter::FormatDiagonalBorder::BorderUp,
            };
            fmt = fmt.set_border_diagonal_type(diag_type);
        }
    }

    // Fill
    if let Some(fill) = obj.get("fill").and_then(|v| v.as_object()) {
        if let Some(ft) = fill.get("fill_type").and_then(|v| v.as_str()) {
            let pattern = match ft {
                "solid" => rust_xlsxwriter::FormatPattern::Solid,
                "darkGray" => rust_xlsxwriter::FormatPattern::DarkGray,
                "mediumGray" => rust_xlsxwriter::FormatPattern::MediumGray,
                "lightGray" => rust_xlsxwriter::FormatPattern::LightGray,
                "gray125" => rust_xlsxwriter::FormatPattern::Gray125,
                "gray0625" => rust_xlsxwriter::FormatPattern::Gray0625,
                _ => rust_xlsxwriter::FormatPattern::Solid,
            };
            fmt = fmt.set_pattern(pattern);
        }
        if let Some(sc) = fill.get("start_color").and_then(|v| v.as_str()) {
            if let Some(clr) = parse_color_str(sc) {
                fmt = fmt.set_background_color(clr);
            }
        }
        if let Some(ec) = fill.get("end_color").and_then(|v| v.as_str()) {
            if let Some(clr) = parse_color_str(ec) {
                fmt = fmt.set_foreground_color(clr);
            }
        }
    }

    // Number format
    if let Some(nf) = obj.get("number_format").and_then(|v| v.as_str()) {
        if nf != "General" {
            fmt = fmt.set_num_format(nf);
        }
    }

    Ok(fmt)
}

// =====================================================================
// RustWorkbook PyO3 class
// =====================================================================

#[pyclass]
struct RustWorkbook {
    sheets: Vec<SheetData>,
    defined_names: Vec<(String, String)>,
}

#[pymethods]
impl RustWorkbook {
    #[new]
    fn new() -> Self {
        RustWorkbook {
            sheets: vec![SheetData::new("Sheet".to_string())],
            defined_names: Vec::new(),
        }
    }

    fn add_sheet(&mut self, title: String) -> usize {
        let idx = self.sheets.len();
        self.sheets.push(SheetData::new(title));
        idx
    }

    fn remove_sheet(&mut self, sheet_idx: usize) -> PyResult<()> {
        if sheet_idx >= self.sheets.len() {
            return Err(pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"));
        }
        self.sheets.remove(sheet_idx);
        Ok(())
    }

    fn set_sheet_title(&mut self, sheet_idx: usize, title: String) -> PyResult<()> {
        if sheet_idx >= self.sheets.len() {
            return Err(pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"));
        }
        self.sheets[sheet_idx].title = title;
        Ok(())
    }

    fn get_sheet_title(&self, sheet_idx: usize) -> PyResult<String> {
        if sheet_idx >= self.sheets.len() {
            return Err(pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"));
        }
        Ok(self.sheets[sheet_idx].title.clone())
    }

    fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    fn set_cell_string(&mut self, sheet: usize, row: u32, col: u16, value: String) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let cell_data = if value.starts_with('=') {
            CellData::Formula(value)
        } else {
            CellData::String(value)
        };
        let key = (row, col);
        if let Some(cv) = sd.cells.get_mut(&key) {
            cv.value = cell_data;
        } else {
            sd.cells.insert(key, CellValue { value: cell_data, format_json: None });
        }
        Ok(())
    }

    fn set_cell_number(&mut self, sheet: usize, row: u32, col: u16, value: f64) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        if let Some(cv) = sd.cells.get_mut(&key) {
            cv.value = CellData::Number(value);
        } else {
            sd.cells.insert(key, CellValue { value: CellData::Number(value), format_json: None });
        }
        Ok(())
    }

    fn set_cell_boolean(&mut self, sheet: usize, row: u32, col: u16, value: bool) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        if let Some(cv) = sd.cells.get_mut(&key) {
            cv.value = CellData::Boolean(value);
        } else {
            sd.cells.insert(key, CellValue { value: CellData::Boolean(value), format_json: None });
        }
        Ok(())
    }

    fn set_cell_datetime(&mut self, sheet: usize, row: u32, col: u16, serial: f64, is_date_only: bool) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        let cell_data = CellData::DateTime { serial, is_date_only };
        if let Some(cv) = sd.cells.get_mut(&key) {
            cv.value = cell_data;
        } else {
            sd.cells.insert(key, CellValue { value: cell_data, format_json: None });
        }
        Ok(())
    }

    fn set_cell_empty(&mut self, sheet: usize, row: u32, col: u16) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        if let Some(cv) = sd.cells.get_mut(&key) {
            cv.value = CellData::Empty;
        } else {
            sd.cells.insert(key, CellValue { value: CellData::Empty, format_json: None });
        }
        Ok(())
    }

    fn set_cell_format(&mut self, sheet: usize, row: u32, col: u16, format_json: String) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        if let Some(cv) = sd.cells.get_mut(&key) {
            cv.format_json = Some(format_json);
        } else {
            // Cell doesn't have a value yet, create it as Empty with format
            sd.cells.insert(key, CellValue { value: CellData::Empty, format_json: Some(format_json) });
        }
        Ok(())
    }

    /// Batch-set cell formats. Each item in the list is a tuple (row, col, format_json).
    fn set_cell_formats_batch(&mut self, sheet: usize, formats: &Bound<'_, PyList>) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;

        for item in formats.iter() {
            let tuple: &Bound<'_, PyTuple> = item.downcast()
                .map_err(|_| pyo3::exceptions::PyTypeError::new_err("Each format entry must be a tuple"))?;
            let row: u32 = tuple.get_item(0)?.extract()?;
            let col: u16 = tuple.get_item(1)?.extract()?;
            let format_json: String = tuple.get_item(2)?.extract()?;

            let key = (row, col);
            if let Some(cv) = sd.cells.get_mut(&key) {
                cv.format_json = Some(format_json);
            } else {
                sd.cells.insert(key, CellValue { value: CellData::Empty, format_json: Some(format_json) });
            }
        }
        Ok(())
    }

    fn get_cell_value(&self, py: Python<'_>, sheet: usize, row: u32, col: u16) -> PyResult<PyObject> {
        let sd = self.sheets.get(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        match sd.cells.get(&key) {
            Some(cv) => match &cv.value {
                CellData::String(s) => Ok(s.as_str().into_pyobject(py).unwrap().into_any().unbind()),
                CellData::Number(n) => Ok((*n).into_pyobject(py).unwrap().into_any().unbind()),
                CellData::Boolean(b) => {
                    let py_bool = (*b).into_pyobject(py).unwrap();
                    let owned = py_bool.to_owned();
                    Ok(owned.into_any().unbind())
                },
                CellData::Formula(f) => Ok(f.as_str().into_pyobject(py).unwrap().into_any().unbind()),
                CellData::DateTime { serial, .. } => Ok((*serial).into_pyobject(py).unwrap().into_any().unbind()),
                CellData::Empty => Ok(py.None()),
            },
            None => Ok(py.None()),
        }
    }

    fn set_column_width(&mut self, sheet: usize, col: u16, width: f64) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.column_widths.insert(col, width);
        Ok(())
    }

    fn set_row_height(&mut self, sheet: usize, row: u32, height: f64) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.row_heights.insert(row, height);
        Ok(())
    }

    fn set_freeze_panes(&mut self, sheet: usize, row: u32, col: u16) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.freeze_panes = Some((row, col));
        Ok(())
    }

    fn add_merge_range(&mut self, sheet: usize, r1: u32, c1: u16, r2: u32, c2: u16) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.merged_ranges.push((r1, c1, r2, c2));
        Ok(())
    }

    fn add_hyperlink(&mut self, sheet: usize, row: u32, col: u16, url: String, text: Option<String>, tooltip: Option<String>) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.hyperlinks.push((row, col, url, text, tooltip));
        Ok(())
    }

    fn add_note(&mut self, sheet: usize, row: u32, col: u16, text: String, author: Option<String>) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.notes.push((row, col, text, author));
        Ok(())
    }

    fn set_autofilter(&mut self, sheet: usize, r1: u32, c1: u16, r2: u32, c2: u16) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.autofilter = Some((r1, c1, r2, c2));
        Ok(())
    }

    fn set_protection(&mut self, sheet: usize, json: String) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.protection_json = Some(json);
        Ok(())
    }

    fn set_page_setup(&mut self, sheet: usize, json: String) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.page_setup_json = Some(json);
        Ok(())
    }

    fn add_image(&mut self, sheet: usize, row: u32, col: u16, data: Vec<u8>, scale_width: Option<f64>, scale_height: Option<f64>) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.images.push((row, col, data, scale_width, scale_height));
        Ok(())
    }

    fn add_data_validation(&mut self, sheet: usize, json: String) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.data_validations.push(json);
        Ok(())
    }

    fn add_conditional_format(&mut self, sheet: usize, json: String) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.conditional_formats.push(json);
        Ok(())
    }

    fn clear_cells(&mut self, sheet: usize) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.cells.clear();
        Ok(())
    }

    fn clear_merge_ranges(&mut self, sheet: usize) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.merged_ranges.clear();
        Ok(())
    }

    fn add_defined_name(&mut self, name: String, formula: String) -> PyResult<()> {
        self.defined_names.push((name, formula));
        Ok(())
    }

    fn set_rows_batch(&mut self, sheet: usize, start_row: u32, rows: &Bound<'_, PyList>) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;

        for (row_offset, row_obj) in rows.iter().enumerate() {
            let row_list: &Bound<'_, PyList> = row_obj.downcast()
                .map_err(|_| pyo3::exceptions::PyTypeError::new_err("Each row must be a list"))?;
            let row = start_row + row_offset as u32;

            for (col_idx, cell_obj) in row_list.iter().enumerate() {
                let col = col_idx as u16;
                let key = (row, col);

                let cell_data = if cell_obj.is_none() {
                    CellData::Empty
                } else if let Ok(b) = cell_obj.extract::<bool>() {
                    CellData::Boolean(b)
                } else if let Ok(n) = cell_obj.extract::<f64>() {
                    CellData::Number(n)
                } else if let Ok(s) = cell_obj.extract::<String>() {
                    if s.starts_with('=') {
                        CellData::Formula(s)
                    } else {
                        CellData::String(s)
                    }
                } else {
                    // Skip unsupported types
                    continue;
                };

                if let Some(cv) = sd.cells.get_mut(&key) {
                    cv.value = cell_data;
                } else {
                    sd.cells.insert(key, CellValue { value: cell_data, format_json: None });
                }
            }
        }
        Ok(())
    }

    fn save(&self, py: Python<'_>, path: Option<&str>) -> PyResult<PyObject> {
        let mut workbook = Workbook::new();

        for sd in &self.sheets {
            let worksheet = workbook.add_worksheet();
            worksheet.set_name(&sd.title).map_err(xlsx_err)?;

            // Write cells
            for (&(row, col), cv) in &sd.cells {
                // Build format
                let mut has_format = false;
                let mut fmt = Format::new();

                if let Some(ref json_str) = cv.format_json {
                    match build_format_from_json(json_str) {
                        Ok(f) => {
                            fmt = f;
                            has_format = true;
                        }
                        Err(e) => {
                            return Err(pyo3::exceptions::PyRuntimeError::new_err(
                                format!("Format error: {}", e)
                            ));
                        }
                    }
                }

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
                    CellData::DateTime { serial, .. } => {
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

            // Column widths
            for (&col, &width) in &sd.column_widths {
                worksheet.set_column_width(col, width).map_err(xlsx_err)?;
            }

            // Row heights
            for (&row, &height) in &sd.row_heights {
                worksheet.set_row_height(row, height).map_err(xlsx_err)?;
            }

            // Freeze panes
            if let Some((row, col)) = sd.freeze_panes {
                worksheet.set_freeze_panes(row, col).map_err(xlsx_err)?;
            }

            // Merged cells
            for &(r1, c1, r2, c2) in &sd.merged_ranges {
                worksheet.merge_range(r1, c1, r2, c2, "", &Format::new()).map_err(xlsx_err)?;
            }

            // Hyperlinks
            for (row, col, url, text, tooltip) in &sd.hyperlinks {
                let mut link = rust_xlsxwriter::Url::new(url);
                if let Some(t) = text { link = link.set_text(t); }
                if let Some(tip) = tooltip { link = link.set_tip(tip); }
                worksheet.write_url(*row, *col, &link).map_err(xlsx_err)?;
            }

            // Notes/Comments
            for (row, col, text, author) in &sd.notes {
                let mut note = rust_xlsxwriter::Note::new(text);
                if let Some(a) = author { note = note.set_author(a); }
                worksheet.insert_note(*row, *col, &note).map_err(xlsx_err)?;
            }

            // Autofilter
            if let Some((r1, c1, r2, c2)) = sd.autofilter {
                worksheet.autofilter(r1, c1, r2, c2).map_err(xlsx_err)?;
            }

            // Protection
            if let Some(ref json_str) = sd.protection_json {
                let prot: serde_json::Value = serde_json::from_str(json_str)
                    .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("Protection JSON error: {}", e)))?;
                let obj = prot.as_object().unwrap();

                let password = obj.get("password").and_then(|v| v.as_str());

                let mut opts = rust_xlsxwriter::ProtectionOptions::default();
                // openpyxl: True = protected (can't do), rust_xlsxwriter: true = CAN do
                // For select_locked/unlocked: openpyxl False = can select, rust_xlsxwriter true = can select (same)
                if let Some(v) = obj.get("select_locked_cells").and_then(|v| v.as_bool()) {
                    opts.select_locked_cells = !v;
                }
                if let Some(v) = obj.get("select_unlocked_cells").and_then(|v| v.as_bool()) {
                    opts.select_unlocked_cells = !v;
                }
                if let Some(v) = obj.get("format_cells").and_then(|v| v.as_bool()) {
                    opts.format_cells = !v;
                }
                if let Some(v) = obj.get("format_columns").and_then(|v| v.as_bool()) {
                    opts.format_columns = !v;
                }
                if let Some(v) = obj.get("format_rows").and_then(|v| v.as_bool()) {
                    opts.format_rows = !v;
                }
                if let Some(v) = obj.get("insert_columns").and_then(|v| v.as_bool()) {
                    opts.insert_columns = !v;
                }
                if let Some(v) = obj.get("insert_rows").and_then(|v| v.as_bool()) {
                    opts.insert_rows = !v;
                }
                if let Some(v) = obj.get("insert_hyperlinks").and_then(|v| v.as_bool()) {
                    opts.insert_links = !v;
                }
                if let Some(v) = obj.get("delete_columns").and_then(|v| v.as_bool()) {
                    opts.delete_columns = !v;
                }
                if let Some(v) = obj.get("delete_rows").and_then(|v| v.as_bool()) {
                    opts.delete_rows = !v;
                }
                if let Some(v) = obj.get("sort").and_then(|v| v.as_bool()) {
                    opts.sort = !v;
                }
                if let Some(v) = obj.get("autofilter").and_then(|v| v.as_bool()) {
                    opts.use_autofilter = !v;
                }
                if let Some(v) = obj.get("pivot_tables").and_then(|v| v.as_bool()) {
                    opts.use_pivot_tables = !v;
                }
                if let Some(v) = obj.get("objects").and_then(|v| v.as_bool()) {
                    opts.edit_objects = !v;
                }
                if let Some(v) = obj.get("scenarios").and_then(|v| v.as_bool()) {
                    opts.edit_scenarios = !v;
                }

                if let Some(pw) = password {
                    worksheet.protect_with_password(pw);
                }
                worksheet.protect_with_options(&opts);
            }

            // Page setup
            if let Some(ref json_str) = sd.page_setup_json {
                let page: serde_json::Value = serde_json::from_str(json_str)
                    .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("Page setup JSON error: {}", e)))?;
                let obj = page.as_object().unwrap();

                // Orientation
                if let Some(orient) = obj.get("orientation").and_then(|v| v.as_str()) {
                    match orient {
                        "landscape" => { worksheet.set_landscape(); }
                        "portrait" => { worksheet.set_portrait(); }
                        _ => {}
                    }
                }

                // Paper size
                if let Some(ps) = obj.get("paper_size").and_then(|v| v.as_u64()) {
                    worksheet.set_paper_size(ps as u8);
                }

                // Scale
                if let Some(scale) = obj.get("scale").and_then(|v| v.as_u64()) {
                    worksheet.set_print_scale(scale as u16);
                }

                // Fit to pages
                if obj.contains_key("fit_to_width") || obj.contains_key("fit_to_height") {
                    let w = obj.get("fit_to_width").and_then(|v| v.as_u64()).unwrap_or(0) as u16;
                    let h = obj.get("fit_to_height").and_then(|v| v.as_u64()).unwrap_or(0) as u16;
                    worksheet.set_print_fit_to_pages(w, h);
                }

                // Margins
                if let Some(margins) = obj.get("margins").and_then(|v| v.as_object()) {
                    let left = margins.get("left").and_then(|v| v.as_f64()).unwrap_or(0.75);
                    let right = margins.get("right").and_then(|v| v.as_f64()).unwrap_or(0.75);
                    let top = margins.get("top").and_then(|v| v.as_f64()).unwrap_or(1.0);
                    let bottom = margins.get("bottom").and_then(|v| v.as_f64()).unwrap_or(1.0);
                    let header = margins.get("header").and_then(|v| v.as_f64()).unwrap_or(0.5);
                    let footer = margins.get("footer").and_then(|v| v.as_f64()).unwrap_or(0.5);
                    worksheet.set_margins(left, right, top, bottom, header, footer);
                }

                // Header/Footer
                if let Some(header_str) = obj.get("header").and_then(|v| v.as_str()) {
                    worksheet.set_header(header_str);
                }
                if let Some(footer_str) = obj.get("footer").and_then(|v| v.as_str()) {
                    worksheet.set_footer(footer_str);
                }

                // Print area: "A1:F10" -> parse to 0-based coords
                if let Some(print_area) = obj.get("print_area").and_then(|v| v.as_str()) {
                    if let Some((r1, c1, r2, c2)) = parse_cell_range(print_area) {
                        worksheet.set_print_area(r1, c1, r2, c2).map_err(xlsx_err)?;
                    }
                }

                // Print title rows: "1:3" -> parse to 0-based
                if let Some(rows_str) = obj.get("print_title_rows").and_then(|v| v.as_str()) {
                    if let Some((first, last)) = parse_row_range(rows_str) {
                        worksheet.set_repeat_rows(first, last).map_err(xlsx_err)?;
                    }
                }

                // Print title cols: "A:B" -> parse to 0-based
                if let Some(cols_str) = obj.get("print_title_cols").and_then(|v| v.as_str()) {
                    if let Some((first, last)) = parse_col_range(cols_str) {
                        worksheet.set_repeat_columns(first, last).map_err(xlsx_err)?;
                    }
                }

                // Center horizontally/vertically
                if let Some(ch) = obj.get("center_horizontally").and_then(|v| v.as_bool()) {
                    worksheet.set_print_center_horizontally(ch);
                }
                if let Some(cv) = obj.get("center_vertically").and_then(|v| v.as_bool()) {
                    worksheet.set_print_center_vertically(cv);
                }

                // Gridlines
                if let Some(gl) = obj.get("gridlines").and_then(|v| v.as_bool()) {
                    worksheet.set_print_gridlines(gl);
                }

                // Headings (row/column headers)
                if let Some(h) = obj.get("headings").and_then(|v| v.as_bool()) {
                    if h {
                        worksheet.set_print_headings(true);
                    }
                }
            }

            // Images
            for (row, col, data, scale_w, scale_h) in &sd.images {
                let mut img = rust_xlsxwriter::Image::new_from_buffer(data)
                    .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("Image error: {}", e)))?;
                if let Some(w) = scale_w { img = img.set_scale_width(*w); }
                if let Some(h) = scale_h { img = img.set_scale_height(*h); }
                worksheet.insert_image(*row, *col, &img).map_err(xlsx_err)?;
            }

            // Data Validations
            for json_str in &sd.data_validations {
                let val: serde_json::Value = serde_json::from_str(json_str)
                    .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("DataValidation JSON error: {}", e)))?;
                let obj = val.as_object().ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("DataValidation JSON must be an object"))?;

                let vtype = obj.get("type").and_then(|v| v.as_str()).unwrap_or("");
                let formula1 = obj.get("formula1").and_then(|v| v.as_str()).unwrap_or("");
                let formula2 = obj.get("formula2").and_then(|v| v.as_str());
                let op_str = obj.get("operator").and_then(|v| v.as_str()).unwrap_or("between");

                let mut dv = rust_xlsxwriter::DataValidation::new();

                match vtype {
                    "list" => {
                        // formula1 is like '"Dog,Cat,Bat"' (with quotes) or a cell range
                        let f1 = formula1.trim_matches('"');
                        if f1.contains('!') || f1.starts_with('$') || (f1.contains(':') && !f1.contains(',')) {
                            // Cell range reference
                            dv = dv.allow_list_formula(rust_xlsxwriter::Formula::new(formula1));
                        } else {
                            // Inline list
                            let items: Vec<&str> = f1.split(',').collect();
                            dv = dv.allow_list_strings(&items).map_err(xlsx_err)?;
                        }
                    }
                    "whole" => {
                        if op_str == "between" || op_str == "notBetween" {
                            let v1: i32 = formula1.parse().unwrap_or(0);
                            let v2: i32 = formula2.unwrap_or("0").parse().unwrap_or(0);
                            let rule = if op_str == "between" {
                                rust_xlsxwriter::DataValidationRule::Between(v1, v2)
                            } else {
                                rust_xlsxwriter::DataValidationRule::NotBetween(v1, v2)
                            };
                            dv = dv.allow_whole_number(rule);
                        } else {
                            let formula = rust_xlsxwriter::Formula::new(formula1);
                            let rule = match op_str {
                                "equal" => rust_xlsxwriter::DataValidationRule::EqualTo(formula),
                                "notEqual" => rust_xlsxwriter::DataValidationRule::NotEqualTo(formula),
                                "greaterThan" => rust_xlsxwriter::DataValidationRule::GreaterThan(formula),
                                "greaterThanOrEqual" => rust_xlsxwriter::DataValidationRule::GreaterThanOrEqualTo(formula),
                                "lessThan" => rust_xlsxwriter::DataValidationRule::LessThan(formula),
                                "lessThanOrEqual" => rust_xlsxwriter::DataValidationRule::LessThanOrEqualTo(formula),
                                _ => rust_xlsxwriter::DataValidationRule::EqualTo(formula),
                            };
                            dv = dv.allow_whole_number_formula(rule);
                        }
                    }
                    "decimal" => {
                        if op_str == "between" || op_str == "notBetween" {
                            let v1: f64 = formula1.parse().unwrap_or(0.0);
                            let v2: f64 = formula2.unwrap_or("0").parse().unwrap_or(0.0);
                            let rule = if op_str == "between" {
                                rust_xlsxwriter::DataValidationRule::Between(v1, v2)
                            } else {
                                rust_xlsxwriter::DataValidationRule::NotBetween(v1, v2)
                            };
                            dv = dv.allow_decimal_number(rule);
                        } else {
                            let formula = rust_xlsxwriter::Formula::new(formula1);
                            let rule = match op_str {
                                "equal" => rust_xlsxwriter::DataValidationRule::EqualTo(formula),
                                "notEqual" => rust_xlsxwriter::DataValidationRule::NotEqualTo(formula),
                                "greaterThan" => rust_xlsxwriter::DataValidationRule::GreaterThan(formula),
                                "greaterThanOrEqual" => rust_xlsxwriter::DataValidationRule::GreaterThanOrEqualTo(formula),
                                "lessThan" => rust_xlsxwriter::DataValidationRule::LessThan(formula),
                                "lessThanOrEqual" => rust_xlsxwriter::DataValidationRule::LessThanOrEqualTo(formula),
                                _ => rust_xlsxwriter::DataValidationRule::EqualTo(formula),
                            };
                            dv = dv.allow_decimal_number_formula(rule);
                        }
                    }
                    "textLength" => {
                        if op_str == "between" || op_str == "notBetween" {
                            let v1: u32 = formula1.parse().unwrap_or(0);
                            let v2: u32 = formula2.unwrap_or("0").parse().unwrap_or(0);
                            let rule = if op_str == "between" {
                                rust_xlsxwriter::DataValidationRule::Between(v1, v2)
                            } else {
                                rust_xlsxwriter::DataValidationRule::NotBetween(v1, v2)
                            };
                            dv = dv.allow_text_length(rule);
                        } else {
                            let formula = rust_xlsxwriter::Formula::new(formula1);
                            let rule = match op_str {
                                "equal" => rust_xlsxwriter::DataValidationRule::EqualTo(formula),
                                "notEqual" => rust_xlsxwriter::DataValidationRule::NotEqualTo(formula),
                                "greaterThan" => rust_xlsxwriter::DataValidationRule::GreaterThan(formula),
                                "greaterThanOrEqual" => rust_xlsxwriter::DataValidationRule::GreaterThanOrEqualTo(formula),
                                "lessThan" => rust_xlsxwriter::DataValidationRule::LessThan(formula),
                                "lessThanOrEqual" => rust_xlsxwriter::DataValidationRule::LessThanOrEqualTo(formula),
                                _ => rust_xlsxwriter::DataValidationRule::EqualTo(formula),
                            };
                            dv = dv.allow_text_length_formula(rule);
                        }
                    }
                    "custom" => {
                        dv = dv.allow_custom(rust_xlsxwriter::Formula::new(formula1));
                    }
                    "date" => {
                        // Use formula variant for date validations
                        let formula = rust_xlsxwriter::Formula::new(formula1);
                        if op_str == "between" || op_str == "notBetween" {
                            let formula2_val = rust_xlsxwriter::Formula::new(formula2.unwrap_or(""));
                            let rule = if op_str == "between" {
                                rust_xlsxwriter::DataValidationRule::Between(formula, formula2_val)
                            } else {
                                rust_xlsxwriter::DataValidationRule::NotBetween(formula, formula2_val)
                            };
                            dv = dv.allow_date_formula(rule);
                        } else {
                            let rule = match op_str {
                                "equal" => rust_xlsxwriter::DataValidationRule::EqualTo(formula),
                                "notEqual" => rust_xlsxwriter::DataValidationRule::NotEqualTo(formula),
                                "greaterThan" => rust_xlsxwriter::DataValidationRule::GreaterThan(formula),
                                "greaterThanOrEqual" => rust_xlsxwriter::DataValidationRule::GreaterThanOrEqualTo(formula),
                                "lessThan" => rust_xlsxwriter::DataValidationRule::LessThan(formula),
                                "lessThanOrEqual" => rust_xlsxwriter::DataValidationRule::LessThanOrEqualTo(formula),
                                _ => rust_xlsxwriter::DataValidationRule::EqualTo(formula),
                            };
                            dv = dv.allow_date_formula(rule);
                        }
                    }
                    "time" => {
                        let formula = rust_xlsxwriter::Formula::new(formula1);
                        if op_str == "between" || op_str == "notBetween" {
                            let formula2_val = rust_xlsxwriter::Formula::new(formula2.unwrap_or(""));
                            let rule = if op_str == "between" {
                                rust_xlsxwriter::DataValidationRule::Between(formula, formula2_val)
                            } else {
                                rust_xlsxwriter::DataValidationRule::NotBetween(formula, formula2_val)
                            };
                            dv = dv.allow_time_formula(rule);
                        } else {
                            let rule = match op_str {
                                "equal" => rust_xlsxwriter::DataValidationRule::EqualTo(formula),
                                "notEqual" => rust_xlsxwriter::DataValidationRule::NotEqualTo(formula),
                                "greaterThan" => rust_xlsxwriter::DataValidationRule::GreaterThan(formula),
                                "greaterThanOrEqual" => rust_xlsxwriter::DataValidationRule::GreaterThanOrEqualTo(formula),
                                "lessThan" => rust_xlsxwriter::DataValidationRule::LessThan(formula),
                                "lessThanOrEqual" => rust_xlsxwriter::DataValidationRule::LessThanOrEqualTo(formula),
                                _ => rust_xlsxwriter::DataValidationRule::EqualTo(formula),
                            };
                            dv = dv.allow_time_formula(rule);
                        }
                    }
                    _ => {}
                }

                // Set options
                if let Some(v) = obj.get("allow_blank").and_then(|v| v.as_bool()) {
                    dv = dv.ignore_blank(v);
                }
                if let Some(v) = obj.get("show_dropdown").and_then(|v| v.as_bool()) {
                    // openpyxl: showDropDown=True means HIDE dropdown
                    // rust_xlsxwriter: show_dropdown(true) means SHOW dropdown
                    dv = dv.show_dropdown(!v);
                }
                if let Some(v) = obj.get("show_input_message").and_then(|v| v.as_bool()) {
                    dv = dv.show_input_message(v);
                }
                if let Some(v) = obj.get("show_error_message").and_then(|v| v.as_bool()) {
                    dv = dv.show_error_message(v);
                }
                if let Some(t) = obj.get("input_title").and_then(|v| v.as_str()) {
                    if !t.is_empty() {
                        dv = dv.set_input_title(t).map_err(xlsx_err)?;
                    }
                }
                if let Some(m) = obj.get("input_message").and_then(|v| v.as_str()) {
                    if !m.is_empty() {
                        dv = dv.set_input_message(m).map_err(xlsx_err)?;
                    }
                }
                if let Some(t) = obj.get("error_title").and_then(|v| v.as_str()) {
                    if !t.is_empty() {
                        dv = dv.set_error_title(t).map_err(xlsx_err)?;
                    }
                }
                if let Some(m) = obj.get("error_message").and_then(|v| v.as_str()) {
                    if !m.is_empty() {
                        dv = dv.set_error_message(m).map_err(xlsx_err)?;
                    }
                }
                if let Some(s) = obj.get("error_style").and_then(|v| v.as_str()) {
                    let style = match s {
                        "warning" => rust_xlsxwriter::DataValidationErrorStyle::Warning,
                        "information" => rust_xlsxwriter::DataValidationErrorStyle::Information,
                        _ => rust_xlsxwriter::DataValidationErrorStyle::Stop,
                    };
                    dv = dv.set_error_style(style);
                }

                // Apply to ranges
                if let Some(rs) = obj.get("ranges").and_then(|v| v.as_array()) {
                    for range in rs {
                        let arr = range.as_array().ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("DataValidation range must be an array"))?;
                        if arr.len() < 4 { continue; }
                        let r1 = arr[0].as_u64().unwrap_or(0) as u32;
                        let c1 = arr[1].as_u64().unwrap_or(0) as u16;
                        let r2 = arr[2].as_u64().unwrap_or(0) as u32;
                        let c2 = arr[3].as_u64().unwrap_or(0) as u16;
                        worksheet.add_data_validation(r1, c1, r2, c2, &dv).map_err(xlsx_err)?;
                    }
                }
            }

            // Conditional Formats
            for cf_json_str in &sd.conditional_formats {
                let cf_val: serde_json::Value = serde_json::from_str(cf_json_str)
                    .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("ConditionalFormat JSON error: {}", e)))?;
                let cf_obj = cf_val.as_object().ok_or_else(|| pyo3::exceptions::PyRuntimeError::new_err("ConditionalFormat JSON must be an object"))?;

                let rule_type = cf_obj.get("rule_type").and_then(|v| v.as_str()).unwrap_or("");
                let range_str = cf_obj.get("range").and_then(|v| v.as_str()).unwrap_or("");

                // Parse range
                let (r1, c1, r2, c2) = match parse_cell_range(range_str) {
                    Some(coords) => coords,
                    None => continue,
                };

                match rule_type {
                    "2_color_scale" => {
                        let mut cf = rust_xlsxwriter::ConditionalFormat2ColorScale::new();
                        if let Some(color) = cf_obj.get("start_color").and_then(|v| v.as_str()) {
                            if let Some(clr) = parse_color_str(color) {
                                cf = cf.set_minimum_color(clr);
                            }
                        }
                        if let Some(color) = cf_obj.get("end_color").and_then(|v| v.as_str()) {
                            if let Some(clr) = parse_color_str(color) {
                                cf = cf.set_maximum_color(clr);
                            }
                        }
                        // Set min type/value if specified
                        if let Some(min_type) = cf_obj.get("start_type").and_then(|v| v.as_str()) {
                            let cf_type = parse_cf_type(min_type);
                            let value = cf_obj.get("start_value").and_then(|v| v.as_f64()).unwrap_or(0.0);
                            cf = cf.set_minimum(cf_type, value);
                        }
                        if let Some(max_type) = cf_obj.get("end_type").and_then(|v| v.as_str()) {
                            let cf_type = parse_cf_type(max_type);
                            let value = cf_obj.get("end_value").and_then(|v| v.as_f64()).unwrap_or(0.0);
                            cf = cf.set_maximum(cf_type, value);
                        }
                        worksheet.add_conditional_format(r1, c1, r2, c2, &cf).map_err(xlsx_err)?;
                    }
                    "3_color_scale" => {
                        let mut cf = rust_xlsxwriter::ConditionalFormat3ColorScale::new();
                        if let Some(color) = cf_obj.get("start_color").and_then(|v| v.as_str()) {
                            if let Some(clr) = parse_color_str(color) {
                                cf = cf.set_minimum_color(clr);
                            }
                        }
                        if let Some(color) = cf_obj.get("mid_color").and_then(|v| v.as_str()) {
                            if let Some(clr) = parse_color_str(color) {
                                cf = cf.set_midpoint_color(clr);
                            }
                        }
                        if let Some(color) = cf_obj.get("end_color").and_then(|v| v.as_str()) {
                            if let Some(clr) = parse_color_str(color) {
                                cf = cf.set_maximum_color(clr);
                            }
                        }
                        if let Some(min_type) = cf_obj.get("start_type").and_then(|v| v.as_str()) {
                            let cf_type = parse_cf_type(min_type);
                            let value = cf_obj.get("start_value").and_then(|v| v.as_f64()).unwrap_or(0.0);
                            cf = cf.set_minimum(cf_type, value);
                        }
                        if let Some(mid_type) = cf_obj.get("mid_type").and_then(|v| v.as_str()) {
                            let cf_type = parse_cf_type(mid_type);
                            let value = cf_obj.get("mid_value").and_then(|v| v.as_f64()).unwrap_or(50.0);
                            cf = cf.set_midpoint(cf_type, value);
                        }
                        if let Some(max_type) = cf_obj.get("end_type").and_then(|v| v.as_str()) {
                            let cf_type = parse_cf_type(max_type);
                            let value = cf_obj.get("end_value").and_then(|v| v.as_f64()).unwrap_or(0.0);
                            cf = cf.set_maximum(cf_type, value);
                        }
                        worksheet.add_conditional_format(r1, c1, r2, c2, &cf).map_err(xlsx_err)?;
                    }
                    "data_bar" => {
                        let mut cf = rust_xlsxwriter::ConditionalFormatDataBar::new();
                        if let Some(color) = cf_obj.get("color").and_then(|v| v.as_str()) {
                            if let Some(clr) = parse_color_str(color) {
                                cf = cf.set_fill_color(clr);
                            }
                        }
                        if let Some(bar_only) = cf_obj.get("bar_only").and_then(|v| v.as_bool()) {
                            if bar_only {
                                cf = cf.set_bar_only(true);
                            }
                        }
                        worksheet.add_conditional_format(r1, c1, r2, c2, &cf).map_err(xlsx_err)?;
                    }
                    "icon_set" => {
                        let mut cf = rust_xlsxwriter::ConditionalFormatIconSet::new();
                        if let Some(icon_style) = cf_obj.get("icon_style").and_then(|v| v.as_str()) {
                            let icon_type = match icon_style {
                                "3Arrows" => rust_xlsxwriter::ConditionalFormatIconType::ThreeArrows,
                                "3ArrowsGray" => rust_xlsxwriter::ConditionalFormatIconType::ThreeArrowsGray,
                                "3Flags" => rust_xlsxwriter::ConditionalFormatIconType::ThreeFlags,
                                "3TrafficLights1" => rust_xlsxwriter::ConditionalFormatIconType::ThreeTrafficLights,
                                "3TrafficLights2" => rust_xlsxwriter::ConditionalFormatIconType::ThreeTrafficLightsWithRim,
                                "3Signs" => rust_xlsxwriter::ConditionalFormatIconType::ThreeSigns,
                                "3Symbols" => rust_xlsxwriter::ConditionalFormatIconType::ThreeSymbolsCircled,
                                "3Symbols2" => rust_xlsxwriter::ConditionalFormatIconType::ThreeSymbols,
                                "3Stars" => rust_xlsxwriter::ConditionalFormatIconType::ThreeStars,
                                "3Triangles" => rust_xlsxwriter::ConditionalFormatIconType::ThreeTriangles,
                                "4Arrows" => rust_xlsxwriter::ConditionalFormatIconType::FourArrows,
                                "4ArrowsGray" => rust_xlsxwriter::ConditionalFormatIconType::FourArrowsGray,
                                "4RedToBlack" => rust_xlsxwriter::ConditionalFormatIconType::FourRedToBlack,
                                "4Rating" => rust_xlsxwriter::ConditionalFormatIconType::FourHistograms,
                                "4TrafficLights" => rust_xlsxwriter::ConditionalFormatIconType::FourTrafficLights,
                                "5Arrows" => rust_xlsxwriter::ConditionalFormatIconType::FiveArrows,
                                "5ArrowsGray" => rust_xlsxwriter::ConditionalFormatIconType::FiveArrowsGray,
                                "5Rating" => rust_xlsxwriter::ConditionalFormatIconType::FiveHistograms,
                                "5Quarters" => rust_xlsxwriter::ConditionalFormatIconType::FiveQuadrants,
                                "5Boxes" => rust_xlsxwriter::ConditionalFormatIconType::FiveBoxes,
                                _ => rust_xlsxwriter::ConditionalFormatIconType::ThreeTrafficLights,
                            };
                            cf = cf.set_icon_type(icon_type);
                        }
                        if let Some(reverse) = cf_obj.get("reverse").and_then(|v| v.as_bool()) {
                            if reverse {
                                cf = cf.reverse_icons(true);
                            }
                        }
                        if let Some(show_icons_only) = cf_obj.get("show_icons_only").and_then(|v| v.as_bool()) {
                            if show_icons_only {
                                cf = cf.show_icons_only(true);
                            }
                        }
                        worksheet.add_conditional_format(r1, c1, r2, c2, &cf).map_err(xlsx_err)?;
                    }
                    "cell_is" => {
                        let mut cf = rust_xlsxwriter::ConditionalFormatCell::new();
                        let operator = cf_obj.get("operator").and_then(|v| v.as_str()).unwrap_or("equal");
                        let formula_arr = cf_obj.get("formula").and_then(|v| v.as_array());

                        match operator {
                            "between" | "notBetween" => {
                                if let Some(arr) = formula_arr {
                                    let val1_str = arr.first().and_then(|v| v.as_str()).unwrap_or("0");
                                    let val2_str = arr.get(1).and_then(|v| v.as_str()).unwrap_or("0");
                                    // Try to parse as numbers first, otherwise use as formula strings
                                    let val1: f64 = val1_str.parse().unwrap_or(0.0);
                                    let val2: f64 = val2_str.parse().unwrap_or(0.0);
                                    let rule = if operator == "between" {
                                        rust_xlsxwriter::ConditionalFormatCellRule::Between(val1, val2)
                                    } else {
                                        rust_xlsxwriter::ConditionalFormatCellRule::NotBetween(val1, val2)
                                    };
                                    cf = cf.set_rule(rule);
                                }
                            }
                            _ => {
                                if let Some(arr) = formula_arr {
                                    if let Some(val_str) = arr.first().and_then(|v| v.as_str()) {
                                        // Try to parse as number, otherwise use as string value
                                        if let Ok(num_val) = val_str.parse::<f64>() {
                                            let rule = match operator {
                                                "lessThan" => rust_xlsxwriter::ConditionalFormatCellRule::LessThan(num_val),
                                                "lessThanOrEqual" => rust_xlsxwriter::ConditionalFormatCellRule::LessThanOrEqualTo(num_val),
                                                "greaterThan" => rust_xlsxwriter::ConditionalFormatCellRule::GreaterThan(num_val),
                                                "greaterThanOrEqual" => rust_xlsxwriter::ConditionalFormatCellRule::GreaterThanOrEqualTo(num_val),
                                                "equal" => rust_xlsxwriter::ConditionalFormatCellRule::EqualTo(num_val),
                                                "notEqual" => rust_xlsxwriter::ConditionalFormatCellRule::NotEqualTo(num_val),
                                                _ => rust_xlsxwriter::ConditionalFormatCellRule::EqualTo(num_val),
                                            };
                                            cf = cf.set_rule(rule);
                                        } else {
                                            // Use as string/formula value
                                            let rule = match operator {
                                                "lessThan" => rust_xlsxwriter::ConditionalFormatCellRule::LessThan(val_str.to_string()),
                                                "lessThanOrEqual" => rust_xlsxwriter::ConditionalFormatCellRule::LessThanOrEqualTo(val_str.to_string()),
                                                "greaterThan" => rust_xlsxwriter::ConditionalFormatCellRule::GreaterThan(val_str.to_string()),
                                                "greaterThanOrEqual" => rust_xlsxwriter::ConditionalFormatCellRule::GreaterThanOrEqualTo(val_str.to_string()),
                                                "equal" => rust_xlsxwriter::ConditionalFormatCellRule::EqualTo(val_str.to_string()),
                                                "notEqual" => rust_xlsxwriter::ConditionalFormatCellRule::NotEqualTo(val_str.to_string()),
                                                _ => rust_xlsxwriter::ConditionalFormatCellRule::EqualTo(val_str.to_string()),
                                            };
                                            cf = cf.set_rule(rule);
                                        }
                                    }
                                }
                            }
                        }

                        // Set format if present
                        if let Some(format_obj) = cf_obj.get("format") {
                            let format_str = serde_json::to_string(format_obj)
                                .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("Format serialize error: {}", e)))?;
                            let fmt = build_format_from_json(&format_str)
                                .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(e))?;
                            cf = cf.set_format(fmt);
                        }

                        if let Some(stop) = cf_obj.get("stop_if_true").and_then(|v| v.as_bool()) {
                            if stop {
                                cf = cf.set_stop_if_true(true);
                            }
                        }

                        worksheet.add_conditional_format(r1, c1, r2, c2, &cf).map_err(xlsx_err)?;
                    }
                    "formula" => {
                        let formula_str = cf_obj.get("formula").and_then(|v| v.as_str()).unwrap_or("");
                        let mut cf = rust_xlsxwriter::ConditionalFormatFormula::new();
                        cf = cf.set_rule(formula_str);

                        // Set format if present
                        if let Some(format_obj) = cf_obj.get("format") {
                            let format_str = serde_json::to_string(format_obj)
                                .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("Format serialize error: {}", e)))?;
                            let fmt = build_format_from_json(&format_str)
                                .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(e))?;
                            cf = cf.set_format(fmt);
                        }

                        if let Some(stop) = cf_obj.get("stop_if_true").and_then(|v| v.as_bool()) {
                            if stop {
                                cf = cf.set_stop_if_true(true);
                            }
                        }

                        worksheet.add_conditional_format(r1, c1, r2, c2, &cf).map_err(xlsx_err)?;
                    }
                    _ => {} // Unknown rule type, skip
                }
            }
        }

        // Named ranges
        for (name, formula) in &self.defined_names {
            workbook.define_name(name, formula).map_err(xlsx_err)?;
        }

        // Save to path or return bytes
        match path {
            Some(p) => {
                workbook.save(p).map_err(xlsx_err)?;
                Ok(py.None())
            }
            None => {
                let buf = workbook.save_to_buffer().map_err(xlsx_err)?;
                Ok(PyBytes::new(py, &buf).into())
            }
        }
    }
}

fn xlsx_err(e: XlsxError) -> PyErr {
    pyo3::exceptions::PyRuntimeError::new_err(e.to_string())
}

/// Map a conditional format type string to ConditionalFormatType.
fn parse_cf_type(s: &str) -> rust_xlsxwriter::ConditionalFormatType {
    match s {
        "min" => rust_xlsxwriter::ConditionalFormatType::Lowest,
        "max" => rust_xlsxwriter::ConditionalFormatType::Highest,
        "num" | "number" => rust_xlsxwriter::ConditionalFormatType::Number,
        "percent" => rust_xlsxwriter::ConditionalFormatType::Percent,
        "percentile" => rust_xlsxwriter::ConditionalFormatType::Percentile,
        "formula" => rust_xlsxwriter::ConditionalFormatType::Formula,
        _ => rust_xlsxwriter::ConditionalFormatType::Automatic,
    }
}

/// Parse column letters (e.g. "A") to 0-based column index.
fn col_letters_to_index(letters: &str) -> Option<u16> {
    let mut col: u16 = 0;
    for ch in letters.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        col = col * 26 + (ch.to_ascii_uppercase() as u16 - b'A' as u16 + 1);
    }
    if col == 0 { return None; }
    Some(col - 1) // convert to 0-based
}

/// Parse a cell reference like "A1" to (row_0based, col_0based).
fn parse_cell_ref(s: &str) -> Option<(u32, u16)> {
    let s = s.trim();
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in s.chars() {
        if ch.is_ascii_alphabetic() {
            if !digits.is_empty() { return None; }
            letters.push(ch);
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else {
            return None;
        }
    }
    let col = col_letters_to_index(&letters)?;
    let row: u32 = digits.parse().ok()?;
    if row == 0 { return None; }
    Some((row - 1, col))
}

/// Parse a cell range like "A1:F10" to (r1, c1, r2, c2) all 0-based.
fn parse_cell_range(s: &str) -> Option<(u32, u16, u32, u16)> {
    let parts: Vec<&str> = s.split(':').collect();
    if parts.len() != 2 { return None; }
    let (r1, c1) = parse_cell_ref(parts[0])?;
    let (r2, c2) = parse_cell_ref(parts[1])?;
    Some((r1, c1, r2, c2))
}

/// Parse a row range like "1:3" to (first_row_0based, last_row_0based).
fn parse_row_range(s: &str) -> Option<(u32, u32)> {
    let parts: Vec<&str> = s.split(':').collect();
    if parts.len() != 2 { return None; }
    let first: u32 = parts[0].trim().parse().ok()?;
    let last: u32 = parts[1].trim().parse().ok()?;
    if first == 0 || last == 0 { return None; }
    Some((first - 1, last - 1))
}

/// Parse a column range like "A:B" to (first_col_0based, last_col_0based).
fn parse_col_range(s: &str) -> Option<(u16, u16)> {
    let parts: Vec<&str> = s.split(':').collect();
    if parts.len() != 2 { return None; }
    let first = col_letters_to_index(parts[0].trim())?;
    let last = col_letters_to_index(parts[1].trim())?;
    Some((first, last))
}

/// Shared logic: convert a calamine Xlsx reader into a Python dict of sheet data.
fn _convert_workbook_to_py<R: std::io::Read + std::io::Seek>(
    py: Python<'_>,
    mut workbook: Xlsx<R>,
) -> PyResult<PyObject> {
    let sheet_names: Vec<String> = workbook.sheet_names().to_vec();

    let result = PyDict::new(py);
    let names_vec: Vec<PyObject> = sheet_names.iter().map(|s: &String| s.as_str().into_pyobject(py).unwrap().into_any().unbind()).collect();
    let names_list = PyList::new(py, &names_vec)?;
    result.set_item("sheet_names", names_list)?;

    let sheets_dict = PyDict::new(py);

    for name in &sheet_names {
        let range = workbook.worksheet_range(name)
            .map_err(|e: calamine::XlsxError| pyo3::exceptions::PyRuntimeError::new_err(e.to_string()))?;

        let (num_rows, num_cols) = range.get_size();
        let empty_vec: Vec<PyObject> = Vec::new();
        let rows_list = PyList::new(py, &empty_vec)?;

        for r in 0..num_rows {
            let row_list = PyList::new(py, &empty_vec)?;
            for c in 0..num_cols {
                let cell = range.get((r, c));
                let py_val: PyObject = match cell {
                    Some(Data::String(s)) => s.as_str().into_pyobject(py).unwrap().into_any().unbind(),
                    Some(Data::Float(f)) => {
                        let fv = *f;
                        if fv == (fv as i64) as f64 && fv.is_finite() {
                            (fv as i64).into_pyobject(py).unwrap().into_any().unbind()
                        } else {
                            fv.into_pyobject(py).unwrap().into_any().unbind()
                        }
                    }
                    Some(Data::Int(i)) => (*i).into_pyobject(py).unwrap().into_any().unbind(),
                    Some(Data::Bool(b)) => {
                        let py_bool = (*b).into_pyobject(py).unwrap();
                        let bound = py_bool.to_owned();
                        bound.into_any().unbind()
                    }
                    Some(Data::DateTime(dt)) => {
                        let s = dt.to_string();
                        s.into_pyobject(py).unwrap().into_any().unbind()
                    }
                    Some(Data::DateTimeIso(s)) => s.as_str().into_pyobject(py).unwrap().into_any().unbind(),
                    Some(Data::DurationIso(s)) => s.as_str().into_pyobject(py).unwrap().into_any().unbind(),
                    Some(Data::Error(e)) => format!("#ERROR: {:?}", e).into_pyobject(py).unwrap().into_any().unbind(),
                    Some(Data::Empty) | None => py.None(),
                };
                row_list.append(py_val)?;
            }
            rows_list.append(row_list)?;
        }

        sheets_dict.set_item(name.as_str(), rows_list)?;
    }

    result.set_item("sheets", sheets_dict)?;
    Ok(result.into())
}

#[pyfunction]
fn _load_workbook(py: Python<'_>, path: &str) -> PyResult<PyObject> {
    let workbook: Xlsx<_> = open_workbook(path)
        .map_err(|e: calamine::XlsxError| pyo3::exceptions::PyIOError::new_err(e.to_string()))?;
    _convert_workbook_to_py(py, workbook)
}

#[pyfunction]
fn _load_workbook_bytes(py: Python<'_>, data: &[u8]) -> PyResult<PyObject> {
    let cursor = Cursor::new(data);
    let workbook: Xlsx<_> = Xlsx::new(cursor)
        .map_err(|e: calamine::XlsxError| pyo3::exceptions::PyIOError::new_err(e.to_string()))?;
    _convert_workbook_to_py(py, workbook)
}

#[pymodule]
fn _openpyxl_rust(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(_load_workbook, m)?)?;
    m.add_function(wrap_pyfunction!(_load_workbook_bytes, m)?)?;
    m.add_class::<RustWorkbook>()?;
    Ok(())
}
