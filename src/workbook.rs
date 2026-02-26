use pyo3::prelude::*;
use pyo3::types::PyList;
use crate::types::*;

#[pyclass]
pub(crate) struct RustWorkbook {
    pub(crate) sheets: Vec<SheetData>,
    pub(crate) defined_names: Vec<(String, String)>,
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
            sd.cells.insert(key, CellValue { value: cell_data, format: CellFormat::default() });
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
            sd.cells.insert(key, CellValue { value: CellData::Number(value), format: CellFormat::default() });
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
            sd.cells.insert(key, CellValue { value: CellData::Boolean(value), format: CellFormat::default() });
        }
        Ok(())
    }

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

    fn set_cell_empty(&mut self, sheet: usize, row: u32, col: u16) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        if let Some(cv) = sd.cells.get_mut(&key) {
            cv.value = CellData::Empty;
        } else {
            sd.cells.insert(key, CellValue { value: CellData::Empty, format: CellFormat::default() });
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
                CellData::DateTime(serial) => Ok((*serial).into_pyobject(py).unwrap().into_any().unbind()),
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
                    sd.cells.insert(key, CellValue { value: cell_data, format: CellFormat::default() });
                }
            }
        }
        Ok(())
    }

    fn set_cell_value(&mut self, _py: Python<'_>, sheet: usize, row: u32, col: u16, value: &Bound<'_, pyo3::PyAny>) -> PyResult<()> {
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
            if s.starts_with('=') {
                CellData::Formula(s)
            } else {
                CellData::String(s)
            }
        } else {
            CellData::Empty
        };
        let cv = sd.cells.entry(key).or_insert_with(|| CellValue { value: CellData::Empty, format: CellFormat::default() });
        cv.value = cell_data;
        Ok(())
    }

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

    fn set_cell_fill(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        fill_type: Option<u8>,
        start_color: Option<String>,
        end_color: Option<String>,
    ) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        let cv = sd.cells.entry(key).or_insert_with(|| CellValue { value: CellData::Empty, format: CellFormat::default() });
        cv.format.fill_type = fill_type;
        cv.format.fill_start_color = start_color;
        cv.format.fill_end_color = end_color;
        Ok(())
    }

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

    fn set_cell_number_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        format: String,
    ) -> PyResult<()> {
        let sd = self.sheets.get_mut(sheet)
            .ok_or_else(|| pyo3::exceptions::PyIndexError::new_err("Sheet index out of range"))?;
        let key = (row, col);
        let cv = sd.cells.entry(key).or_insert_with(|| CellValue { value: CellData::Empty, format: CellFormat::default() });
        cv.format.number_format = Some(format);
        Ok(())
    }



    fn save(&self, py: Python<'_>, path: Option<&str>) -> PyResult<PyObject> {
        crate::save::save_workbook(&self.sheets, &self.defined_names, path, py)
    }
}
