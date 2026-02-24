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

    let sheets_obj = data.get_item("sheets")?.unwrap();
    let sheets: &Bound<'_, PyList> = sheets_obj.downcast()?;

    for sheet_obj in sheets.iter() {
        let sheet_dict: &Bound<'_, PyDict> = sheet_obj.downcast()?;
        let title: String = sheet_dict
            .get_item("title")?
            .unwrap()
            .extract()?;

        let worksheet = workbook.add_worksheet();
        worksheet.set_name(&title).map_err(xlsx_err)?;

        // Write cells
        let cells_obj = sheet_dict.get_item("cells")?.unwrap();
        let cells: &Bound<'_, PyList> = cells_obj.downcast()?;
        for cell_obj in cells.iter() {
            let cell: &Bound<'_, PyDict> = cell_obj.downcast()?;
            let row: u32 = cell.get_item("row")?.unwrap().extract()?;
            let col: u16 = cell.get_item("col")?.unwrap().extract()?;

            // Build format if font or number_format present
            let mut has_format = false;
            let mut fmt = Format::new();

            if let Ok(Some(font_obj)) = cell.get_item("font") {
                if !font_obj.is_none() {
                    let font_dict: &Bound<'_, PyDict> = font_obj.downcast()?;
                    fmt = make_format(font_dict)?;
                    has_format = true;
                }
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
                // blank cell
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
