use crate::types::*;
use pyo3::prelude::*;
use pyo3::types::PyList;

/// Convert Excel serial number to (year, month, day) using inverse Julian Day algorithm.
fn serial_to_date_parts(serial: f64) -> (i32, u32, u32) {
    let mut s = serial.floor() as i64;
    if s > 59 {
        s -= 1;
    } // Undo Lotus 1-2-3 bug
    let j = s + 2_415_020; // Convert to Julian Day Number
    let a = j + 32_044;
    let b = (4 * a + 3) / 146_097;
    let c = a - (146_097 * b) / 4;
    let d = (4 * c + 3) / 1_461;
    let e = c - (1_461 * d) / 4;
    let m = (5 * e + 2) / 153;
    let day = e - (153 * m + 2) / 5 + 1;
    let month = m + 3 - 12 * (m / 10);
    let year = 100 * b + d - 4800 + m / 10;
    (year as i32, month as u32, day as u32)
}

/// Convert fractional day to (hours, minutes, seconds, microseconds).
fn serial_to_time_parts(serial: f64) -> (u32, u32, u32, u32) {
    let frac = serial.fract().abs();
    let total_us = (frac * 86_400_000_000.0).round() as u64;
    let hours = (total_us / 3_600_000_000) as u32;
    let rem = total_us % 3_600_000_000;
    let minutes = (rem / 60_000_000) as u32;
    let rem2 = rem % 60_000_000;
    let seconds = (rem2 / 1_000_000) as u32;
    let microseconds = (rem2 % 1_000_000) as u32;
    (hours, minutes, seconds, microseconds)
}

/// Convert a DateTime serial + kind to a Python datetime/date/time object.
fn datetime_to_py(py: Python<'_>, serial: f64, kind: u8) -> PyResult<PyObject> {
    let dt_mod = py.import("datetime")?;
    match kind {
        0 => {
            // date
            let (y, m, d) = serial_to_date_parts(serial);
            Ok(dt_mod.getattr("date")?.call1((y, m, d))?.unbind())
        }
        1 => {
            // time
            let (h, min, s, us) = serial_to_time_parts(serial);
            Ok(dt_mod.getattr("time")?.call1((h, min, s, us))?.unbind())
        }
        _ => {
            // datetime
            let (y, m, d) = serial_to_date_parts(serial);
            let (h, min, s, us) = serial_to_time_parts(serial);
            Ok(dt_mod
                .getattr("datetime")?
                .call1((y, m, d, h, min, s, us))?
                .unbind())
        }
    }
}

#[pyclass]
pub(crate) struct RustWorkbook {
    pub(crate) sheets: Vec<SheetData>,
    pub(crate) defined_names: Vec<(String, String)>,
    pub(crate) doc_properties_json: Option<String>,
}

#[pymethods]
impl RustWorkbook {
    #[new]
    fn new() -> Self {
        RustWorkbook {
            sheets: vec![SheetData::new("Sheet".to_string())],
            defined_names: Vec::new(),
            doc_properties_json: None,
        }
    }

    fn add_sheet(&mut self, title: String) -> usize {
        let idx = self.sheets.len();
        self.sheets.push(SheetData::new(title));
        idx
    }

    fn remove_sheet(&mut self, sheet_idx: usize) -> PyResult<()> {
        if sheet_idx >= self.sheets.len() {
            return Err(pyo3::exceptions::PyIndexError::new_err(
                "Sheet index out of range",
            ));
        }
        self.sheets.remove(sheet_idx);
        Ok(())
    }

    fn set_sheet_title(&mut self, sheet_idx: usize, title: String) -> PyResult<()> {
        if sheet_idx >= self.sheets.len() {
            return Err(pyo3::exceptions::PyIndexError::new_err(
                "Sheet index out of range",
            ));
        }
        self.sheets[sheet_idx].title = title;
        Ok(())
    }

    fn get_sheet_title(&self, sheet_idx: usize) -> PyResult<String> {
        if sheet_idx >= self.sheets.len() {
            return Err(pyo3::exceptions::PyIndexError::new_err(
                "Sheet index out of range",
            ));
        }
        Ok(self.sheets[sheet_idx].title.clone())
    }

    fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    fn set_cell_string(&mut self, sheet: usize, row: u32, col: u16, value: String) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
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
            sd.cells.insert(
                key,
                CellValue {
                    value: cell_data,
                    format: CellFormat::default(),
                },
            );
        }
        sd.track_cell(row, col);
        Ok(())
    }

    fn set_cell_number(&mut self, sheet: usize, row: u32, col: u16, value: f64) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        if let Some(cv) = sd.cells.get_mut(&key) {
            cv.value = CellData::Number(value);
        } else {
            sd.cells.insert(
                key,
                CellValue {
                    value: CellData::Number(value),
                    format: CellFormat::default(),
                },
            );
        }
        sd.track_cell(row, col);
        Ok(())
    }

    fn set_cell_boolean(&mut self, sheet: usize, row: u32, col: u16, value: bool) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        if let Some(cv) = sd.cells.get_mut(&key) {
            cv.value = CellData::Boolean(value);
        } else {
            sd.cells.insert(
                key,
                CellValue {
                    value: CellData::Boolean(value),
                    format: CellFormat::default(),
                },
            );
        }
        sd.track_cell(row, col);
        Ok(())
    }

    fn set_cell_datetime(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        serial: f64,
        kind: u8,
    ) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        if let Some(cv) = sd.cells.get_mut(&key) {
            cv.value = CellData::DateTime(serial, kind);
        } else {
            sd.cells.insert(
                key,
                CellValue {
                    value: CellData::DateTime(serial, kind),
                    format: CellFormat::default(),
                },
            );
        }
        sd.track_cell(row, col);
        Ok(())
    }

    fn set_cell_empty(&mut self, sheet: usize, row: u32, col: u16) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        if let Some(cv) = sd.cells.get_mut(&key) {
            cv.value = CellData::Empty;
        } else {
            sd.cells.insert(
                key,
                CellValue {
                    value: CellData::Empty,
                    format: CellFormat::default(),
                },
            );
        }
        sd.track_cell(row, col);
        Ok(())
    }

    fn get_cell_value(
        &self,
        py: Python<'_>,
        sheet: usize,
        row: u32,
        col: u16,
    ) -> PyResult<PyObject> {
        let sd = self
            .sheets
            .get(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        match sd.cells.get(&key) {
            Some(cv) => match &cv.value {
                CellData::String(s) => {
                    Ok(s.as_str().into_pyobject(py).unwrap().into_any().unbind())
                }
                CellData::Number(n) => Ok((*n).into_pyobject(py).unwrap().into_any().unbind()),
                CellData::Boolean(b) => {
                    let py_bool = (*b).into_pyobject(py).unwrap();
                    let owned = py_bool.to_owned();
                    Ok(owned.into_any().unbind())
                }
                CellData::Formula(f) => {
                    Ok(f.as_str().into_pyobject(py).unwrap().into_any().unbind())
                }
                CellData::DateTime(serial, kind) => datetime_to_py(py, *serial, *kind),
                CellData::Empty => Ok(py.None()),
            },
            None => Ok(py.None()),
        }
    }

    fn set_column_width(&mut self, sheet: usize, col: u16, width: f64) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.column_widths.insert(col, width);
        Ok(())
    }

    fn set_row_height(&mut self, sheet: usize, row: u32, height: f64) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.row_heights.insert(row, height);
        Ok(())
    }

    fn set_freeze_panes(&mut self, sheet: usize, row: u32, col: u16) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.freeze_panes = Some((row, col));
        Ok(())
    }

    fn add_merge_range(
        &mut self,
        sheet: usize,
        r1: u32,
        c1: u16,
        r2: u32,
        c2: u16,
    ) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.merged_ranges.push((r1, c1, r2, c2));
        Ok(())
    }

    fn add_hyperlink(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        url: String,
        text: Option<String>,
        tooltip: Option<String>,
    ) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.hyperlinks.push((row, col, url, text, tooltip));
        Ok(())
    }

    fn add_note(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        text: String,
        author: Option<String>,
    ) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.notes.push((row, col, text, author));
        Ok(())
    }

    fn set_autofilter(&mut self, sheet: usize, r1: u32, c1: u16, r2: u32, c2: u16) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.autofilter = Some((r1, c1, r2, c2));
        Ok(())
    }

    fn set_protection(&mut self, sheet: usize, json: String) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.protection_json = Some(json);
        Ok(())
    }

    fn set_page_setup(&mut self, sheet: usize, json: String) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.page_setup_json = Some(json);
        Ok(())
    }

    fn add_image(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        data: Vec<u8>,
        scale_width: Option<f64>,
        scale_height: Option<f64>,
    ) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.images.push((row, col, data, scale_width, scale_height));
        Ok(())
    }

    fn add_data_validation(&mut self, sheet: usize, json: String) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.data_validations.push(json);
        Ok(())
    }

    fn add_conditional_format(&mut self, sheet: usize, json: String) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.conditional_formats.push(json);
        Ok(())
    }

    fn add_table(&mut self, sheet: usize, json: String) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.tables.push(json);
        Ok(())
    }

    fn add_chart(&mut self, sheet: usize, json: String) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.charts.push(json);
        Ok(())
    }

    fn clear_cells(&mut self, sheet: usize) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.cells.clear();
        Ok(())
    }

    fn clear_merge_ranges(&mut self, sheet: usize) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.merged_ranges.clear();
        Ok(())
    }

    fn add_defined_name(&mut self, name: String, formula: String) -> PyResult<()> {
        self.defined_names.push((name, formula));
        Ok(())
    }

    fn set_rows_batch(
        &mut self,
        sheet: usize,
        start_row: u32,
        rows: &Bound<'_, PyList>,
    ) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;

        for (row_offset, row_obj) in rows.iter().enumerate() {
            let row_list: &Bound<'_, PyList> = row_obj
                .downcast()
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
                    sd.cells.insert(
                        key,
                        CellValue {
                            value: cell_data,
                            format: CellFormat::default(),
                        },
                    );
                }
                sd.track_cell(row, col);
            }
        }
        Ok(())
    }

    fn set_cell_value(
        &mut self,
        _py: Python<'_>,
        sheet: usize,
        row: u32,
        col: u16,
        value: &Bound<'_, pyo3::PyAny>,
    ) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        let cell_data = if value.is_none() {
            CellData::Empty
        } else if let Ok(b) = value.extract::<bool>() {
            CellData::Boolean(b)
        } else if let Ok(n) = value.extract::<f64>() {
            CellData::Number(n)
        } else if let Ok(s) = value.extract::<String>() {
            if s.starts_with('=') {
                CellData::Formula(s)
            } else {
                CellData::String(s)
            }
        } else {
            CellData::Empty
        };
        let cv = sd.cells.entry(key).or_insert_with(|| CellValue {
            value: CellData::Empty,
            format: CellFormat::default(),
        });
        cv.value = cell_data;
        sd.track_cell(row, col);
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn set_cell_font(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        bold: bool,
        italic: bool,
        name: Option<String>,
        size: Option<f64>,
        color: Option<String>,
        underline: Option<u8>,
        strikethrough: bool,
        vert_align: Option<u8>,
    ) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        let cv = sd.cells.entry(key).or_insert_with(|| CellValue {
            value: CellData::Empty,
            format: CellFormat::default(),
        });
        cv.format.font_bold = bold;
        cv.format.font_italic = italic;
        cv.format.font_name = name;
        cv.format.font_size = size;
        cv.format.font_color = color;
        cv.format.font_underline = underline;
        cv.format.font_strikethrough = strikethrough;
        cv.format.font_vert_align = vert_align;
        sd.track_cell(row, col);
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn set_cell_alignment(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        horizontal: Option<u8>,
        vertical: Option<u8>,
        wrap_text: bool,
        shrink_to_fit: bool,
        indent: u8,
        text_rotation: i16,
    ) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        let cv = sd.cells.entry(key).or_insert_with(|| CellValue {
            value: CellData::Empty,
            format: CellFormat::default(),
        });
        cv.format.align_horizontal = horizontal;
        cv.format.align_vertical = vertical;
        cv.format.align_wrap_text = wrap_text;
        cv.format.align_shrink_to_fit = shrink_to_fit;
        cv.format.align_indent = indent;
        cv.format.align_text_rotation = text_rotation;
        sd.track_cell(row, col);
        Ok(())
    }

    fn set_cell_fill(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        fill_type: Option<u8>,
        start_color: Option<String>,
        end_color: Option<String>,
    ) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        let cv = sd.cells.entry(key).or_insert_with(|| CellValue {
            value: CellData::Empty,
            format: CellFormat::default(),
        });
        cv.format.fill_type = fill_type;
        cv.format.fill_start_color = start_color;
        cv.format.fill_end_color = end_color;
        sd.track_cell(row, col);
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn set_cell_border(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        left_style: Option<u8>,
        left_color: Option<String>,
        right_style: Option<u8>,
        right_color: Option<String>,
        top_style: Option<u8>,
        top_color: Option<String>,
        bottom_style: Option<u8>,
        bottom_color: Option<String>,
        diag_style: Option<u8>,
        diag_color: Option<String>,
        diag_up: bool,
        diag_down: bool,
    ) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        let cv = sd.cells.entry(key).or_insert_with(|| CellValue {
            value: CellData::Empty,
            format: CellFormat::default(),
        });
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
        sd.track_cell(row, col);
        Ok(())
    }

    fn set_cell_number_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        format: String,
    ) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        let cv = sd.cells.entry(key).or_insert_with(|| CellValue {
            value: CellData::Empty,
            format: CellFormat::default(),
        });
        cv.format.number_format = Some(format);
        sd.track_cell(row, col);
        Ok(())
    }

    fn get_dimensions(
        &self,
        sheet: usize,
    ) -> PyResult<(Option<u32>, Option<u16>, Option<u32>, Option<u16>)> {
        #![allow(clippy::type_complexity)]
        let sd = self
            .sheets
            .get(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        Ok((sd.min_row, sd.min_col, sd.max_row, sd.max_col))
    }

    fn touch_cell(&mut self, sheet: usize, row: u32, col: u16) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.track_cell(row, col);
        let key = (row, col);
        sd.cells.entry(key).or_insert_with(|| CellValue {
            value: CellData::Empty,
            format: CellFormat::default(),
        });
        Ok(())
    }

    fn get_next_append_row(&self, sheet: usize) -> PyResult<u32> {
        let sd = self
            .sheets
            .get(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let max = sd.max_row.map(|r| r + 1).unwrap_or(0);
        Ok(std::cmp::max(sd.append_row, max))
    }

    fn set_next_append_row(&mut self, sheet: usize, row: u32) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.append_row = row;
        Ok(())
    }

    fn get_rows_batch(
        &self,
        py: Python<'_>,
        sheet: usize,
        min_row: u32,
        min_col: u16,
        max_row: u32,
        max_col: u16,
    ) -> PyResult<PyObject> {
        let sd = self
            .sheets
            .get(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let result = pyo3::types::PyList::empty(py);
        for r in min_row..=max_row {
            let row_list = pyo3::types::PyList::empty(py);
            for c in min_col..=max_col {
                let val = match sd.cells.get(&(r, c)) {
                    Some(cv) => match &cv.value {
                        CellData::String(s) => {
                            s.as_str().into_pyobject(py).unwrap().into_any().unbind()
                        }
                        CellData::Number(n) => (*n).into_pyobject(py).unwrap().into_any().unbind(),
                        CellData::Boolean(b) => (*b)
                            .into_pyobject(py)
                            .unwrap()
                            .to_owned()
                            .into_any()
                            .unbind(),
                        CellData::Formula(f) => {
                            f.as_str().into_pyobject(py).unwrap().into_any().unbind()
                        }
                        CellData::DateTime(s, k) => datetime_to_py(py, *s, *k)?,
                        CellData::Empty => py.None(),
                    },
                    None => py.None(),
                };
                row_list.append(val)?;
            }
            result.append(row_list)?;
        }
        Ok(result.into())
    }

    fn rust_insert_rows(&mut self, sheet: usize, idx: u32, amount: u32) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let keys: Vec<(u32, u16)> = sd.cells.keys().cloned().collect();
        let mut new_cells = std::collections::HashMap::new();
        for key in keys {
            let (r, c) = key;
            let cv = sd.cells.remove(&key).unwrap();
            if r >= idx {
                new_cells.insert((r + amount, c), cv);
            } else {
                new_cells.insert((r, c), cv);
            }
        }
        sd.cells = new_cells;
        for range in &mut sd.merged_ranges {
            if range.0 >= idx {
                range.0 += amount;
            }
            if range.2 >= idx {
                range.2 += amount;
            }
        }
        for h in &mut sd.hyperlinks {
            if h.0 >= idx {
                h.0 += amount;
            }
        }
        for n in &mut sd.notes {
            if n.0 >= idx {
                n.0 += amount;
            }
        }
        if sd.append_row >= idx {
            sd.append_row += amount;
        }
        sd.recompute_dimensions();
        Ok(())
    }

    fn rust_delete_rows(&mut self, sheet: usize, idx: u32, amount: u32) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let keys: Vec<(u32, u16)> = sd.cells.keys().cloned().collect();
        let mut new_cells = std::collections::HashMap::new();
        for key in keys {
            let (r, c) = key;
            let cv = sd.cells.remove(&key).unwrap();
            if r >= idx && r < idx + amount {
                // deleted
            } else if r >= idx + amount {
                new_cells.insert((r - amount, c), cv);
            } else {
                new_cells.insert((r, c), cv);
            }
        }
        sd.cells = new_cells;
        sd.merged_ranges
            .retain(|range| !(range.0 >= idx && range.2 < idx + amount));
        for range in &mut sd.merged_ranges {
            if range.0 >= idx + amount {
                range.0 -= amount;
            }
            if range.2 >= idx + amount {
                range.2 -= amount;
            }
        }
        sd.hyperlinks
            .retain(|h| !(h.0 >= idx && h.0 < idx + amount));
        for h in &mut sd.hyperlinks {
            if h.0 >= idx + amount {
                h.0 -= amount;
            }
        }
        sd.notes.retain(|n| !(n.0 >= idx && n.0 < idx + amount));
        for n in &mut sd.notes {
            if n.0 >= idx + amount {
                n.0 -= amount;
            }
        }
        if sd.append_row >= idx + amount {
            sd.append_row -= amount;
        } else if sd.append_row >= idx {
            sd.append_row = idx;
        }
        sd.recompute_dimensions();
        Ok(())
    }

    fn rust_insert_cols(&mut self, sheet: usize, idx: u16, amount: u16) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let keys: Vec<(u32, u16)> = sd.cells.keys().cloned().collect();
        let mut new_cells = std::collections::HashMap::new();
        for key in keys {
            let (r, c) = key;
            let cv = sd.cells.remove(&key).unwrap();
            if c >= idx {
                new_cells.insert((r, c + amount), cv);
            } else {
                new_cells.insert((r, c), cv);
            }
        }
        sd.cells = new_cells;
        for range in &mut sd.merged_ranges {
            if range.1 >= idx {
                range.1 += amount;
            }
            if range.3 >= idx {
                range.3 += amount;
            }
        }
        for h in &mut sd.hyperlinks {
            if h.1 >= idx {
                h.1 += amount;
            }
        }
        for n in &mut sd.notes {
            if n.1 >= idx {
                n.1 += amount;
            }
        }
        sd.recompute_dimensions();
        Ok(())
    }

    fn rust_delete_cols(&mut self, sheet: usize, idx: u16, amount: u16) -> PyResult<()> {
        let sd = self
            .sheets
            .get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let keys: Vec<(u32, u16)> = sd.cells.keys().cloned().collect();
        let mut new_cells = std::collections::HashMap::new();
        for key in keys {
            let (r, c) = key;
            let cv = sd.cells.remove(&key).unwrap();
            if c >= idx && c < idx + amount { /* deleted */
            } else if c >= idx + amount {
                new_cells.insert((r, c - amount), cv);
            } else {
                new_cells.insert((r, c), cv);
            }
        }
        sd.cells = new_cells;
        sd.merged_ranges
            .retain(|range| !(range.1 >= idx && range.3 < idx + amount));
        for range in &mut sd.merged_ranges {
            if range.1 >= idx + amount {
                range.1 -= amount;
            }
            if range.3 >= idx + amount {
                range.3 -= amount;
            }
        }
        sd.hyperlinks
            .retain(|h| !(h.1 >= idx && h.1 < idx + amount));
        for h in &mut sd.hyperlinks {
            if h.1 >= idx + amount {
                h.1 -= amount;
            }
        }
        sd.notes.retain(|n| !(n.1 >= idx && n.1 < idx + amount));
        for n in &mut sd.notes {
            if n.1 >= idx + amount {
                n.1 -= amount;
            }
        }
        sd.recompute_dimensions();
        Ok(())
    }

    fn set_sheet_visibility(&mut self, sheet: usize, state: u8) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.visibility = state;
        Ok(())
    }

    fn set_row_hidden(&mut self, sheet: usize, row: u32) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.hidden_rows.push(row);
        Ok(())
    }

    fn set_col_hidden(&mut self, sheet: usize, col: u16) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.hidden_cols.push(col);
        Ok(())
    }

    fn set_zoom(&mut self, sheet: usize, zoom: u16) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.zoom = Some(zoom);
        Ok(())
    }

    fn set_show_gridlines(&mut self, sheet: usize, show: bool) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.show_gridlines = Some(show);
        Ok(())
    }

    fn set_autofit(&mut self, sheet: usize, enabled: bool) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.autofit = enabled;
        Ok(())
    }

    fn set_row_breaks(&mut self, sheet: usize, breaks: Vec<u32>) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.row_breaks = breaks;
        Ok(())
    }

    fn set_col_breaks(&mut self, sheet: usize, breaks: Vec<u16>) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.col_breaks = breaks;
        Ok(())
    }

    fn set_row_outline_level(&mut self, sheet: usize, row: u32, level: u8) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.row_outline_levels.push((row, level));
        Ok(())
    }

    fn set_col_outline_level(&mut self, sheet: usize, col: u16, level: u8) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        sd.col_outline_levels.push((col, level));
        Ok(())
    }

    fn set_doc_properties(&mut self, json: String) -> PyResult<()> {
        self.doc_properties_json = Some(json);
        Ok(())
    }

    fn save(&self, py: Python<'_>, path: Option<&str>) -> PyResult<PyObject> {
        crate::save::save_workbook(&self.sheets, &self.defined_names, self.doc_properties_json.as_deref(), path, py)
    }
}
