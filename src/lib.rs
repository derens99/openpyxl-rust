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

fn date_to_excel_serial(year: i32, month: i32, day: i32) -> f64 {
    let mut y = year;
    let mut m = month;
    if m <= 2 {
        y -= 1;
        m += 12;
    }
    let a = y / 100;
    let b = 2 - a + a / 4;
    let jd = (365.25 * (y + 4716) as f64) as i64
        + (30.6001 * (m + 1) as f64) as i64
        + day as i64
        + b as i64
        - 1524;
    let excel_epoch_jd: i64 = 2415019;
    let mut serial = (jd - excel_epoch_jd) as f64;
    if serial > 59.0 {
        serial += 1.0;
    }
    serial
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

            // Alignment
            if let Ok(Some(align_obj)) = cell.get_item("alignment") {
                let align: &Bound<'_, PyDict> = align_obj.downcast()?;
                if let Ok(Some(h)) = align.get_item("horizontal") {
                    let hs: String = h.extract()?;
                    let a = match hs.as_str() {
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
                if let Ok(Some(v)) = align.get_item("vertical") {
                    let vs: String = v.extract()?;
                    let a = match vs.as_str() {
                        "center" => rust_xlsxwriter::FormatAlign::VerticalCenter,
                        "top" => rust_xlsxwriter::FormatAlign::Top,
                        "bottom" => rust_xlsxwriter::FormatAlign::Bottom,
                        "justify" => rust_xlsxwriter::FormatAlign::VerticalJustify,
                        "distributed" => rust_xlsxwriter::FormatAlign::VerticalDistributed,
                        _ => rust_xlsxwriter::FormatAlign::Bottom,
                    };
                    fmt = fmt.set_align(a);
                }
                if let Ok(Some(wt)) = align.get_item("wrap_text") {
                    if wt.extract::<bool>()? {
                        fmt = fmt.set_text_wrap();
                    }
                }
                if let Ok(Some(sf)) = align.get_item("shrink_to_fit") {
                    if sf.extract::<bool>()? {
                        fmt = fmt.set_shrink();
                    }
                }
                if let Ok(Some(ind)) = align.get_item("indent") {
                    let i: u8 = ind.extract()?;
                    if i > 0 {
                        fmt = fmt.set_indent(i);
                    }
                }
                if let Ok(Some(rot)) = align.get_item("text_rotation") {
                    let r: i16 = rot.extract()?;
                    if r != 0 {
                        fmt = fmt.set_rotation(r);
                    }
                }
                has_format = true;
            }

            // Border
            if let Ok(Some(border_obj)) = cell.get_item("border") {
                let border: &Bound<'_, PyDict> = border_obj.downcast()?;

                fn parse_border_style(s: &str) -> rust_xlsxwriter::FormatBorder {
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

                fn parse_color(c: &str) -> Option<rust_xlsxwriter::Color> {
                    u32::from_str_radix(c, 16).ok().map(rust_xlsxwriter::Color::from)
                }

                if let Ok(Some(left_obj)) = border.get_item("left") {
                    let left: &Bound<'_, PyDict> = left_obj.downcast()?;
                    if let Ok(Some(style)) = left.get_item("style") {
                        fmt = fmt.set_border_left(parse_border_style(&style.extract::<String>()?));
                    }
                    if let Ok(Some(color)) = left.get_item("color") {
                        if let Some(clr) = parse_color(&color.extract::<String>()?) {
                            fmt = fmt.set_border_left_color(clr);
                        }
                    }
                }
                if let Ok(Some(right_obj)) = border.get_item("right") {
                    let right: &Bound<'_, PyDict> = right_obj.downcast()?;
                    if let Ok(Some(style)) = right.get_item("style") {
                        fmt = fmt.set_border_right(parse_border_style(&style.extract::<String>()?));
                    }
                    if let Ok(Some(color)) = right.get_item("color") {
                        if let Some(clr) = parse_color(&color.extract::<String>()?) {
                            fmt = fmt.set_border_right_color(clr);
                        }
                    }
                }
                if let Ok(Some(top_obj)) = border.get_item("top") {
                    let top: &Bound<'_, PyDict> = top_obj.downcast()?;
                    if let Ok(Some(style)) = top.get_item("style") {
                        fmt = fmt.set_border_top(parse_border_style(&style.extract::<String>()?));
                    }
                    if let Ok(Some(color)) = top.get_item("color") {
                        if let Some(clr) = parse_color(&color.extract::<String>()?) {
                            fmt = fmt.set_border_top_color(clr);
                        }
                    }
                }
                if let Ok(Some(bottom_obj)) = border.get_item("bottom") {
                    let bottom: &Bound<'_, PyDict> = bottom_obj.downcast()?;
                    if let Ok(Some(style)) = bottom.get_item("style") {
                        fmt = fmt.set_border_bottom(parse_border_style(&style.extract::<String>()?));
                    }
                    if let Ok(Some(color)) = bottom.get_item("color") {
                        if let Some(clr) = parse_color(&color.extract::<String>()?) {
                            fmt = fmt.set_border_bottom_color(clr);
                        }
                    }
                }
                has_format = true;
            }

            // Fill
            if let Ok(Some(fill_obj)) = cell.get_item("fill") {
                let fill: &Bound<'_, PyDict> = fill_obj.downcast()?;
                if let Ok(Some(ft)) = fill.get_item("fill_type") {
                    let fts: String = ft.extract()?;
                    let pattern = match fts.as_str() {
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
                if let Ok(Some(sc)) = fill.get_item("start_color") {
                    let c: String = sc.extract()?;
                    if let Ok(rgb) = u32::from_str_radix(&c, 16) {
                        fmt = fmt.set_background_color(rust_xlsxwriter::Color::from(rgb));
                    }
                }
                if let Ok(Some(ec)) = fill.get_item("end_color") {
                    let c: String = ec.extract()?;
                    if let Ok(rgb) = u32::from_str_radix(&c, 16) {
                        fmt = fmt.set_foreground_color(rust_xlsxwriter::Color::from(rgb));
                    }
                }
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

            // Check for datetime/date dict FIRST
            if let Ok(dt_dict) = value_obj.downcast::<PyDict>() {
                if let Ok(Some(type_obj)) = dt_dict.get_item("__type__") {
                    let type_str: String = type_obj.extract()?;
                    let year: i32 = dt_dict.get_item("year")?.unwrap().extract()?;
                    let month: i32 = dt_dict.get_item("month")?.unwrap().extract()?;
                    let day: i32 = dt_dict.get_item("day")?.unwrap().extract()?;
                    let serial = date_to_excel_serial(year, month, day);

                    let value = if type_str == "datetime" {
                        let hour: i32 = dt_dict.get_item("hour")?.unwrap().extract()?;
                        let minute: i32 = dt_dict.get_item("minute")?.unwrap().extract()?;
                        let second: i32 = dt_dict.get_item("second")?.unwrap().extract()?;
                        serial + (hour as f64 * 3600.0 + minute as f64 * 60.0 + second as f64) / 86400.0
                    } else {
                        serial
                    };

                    if has_format {
                        worksheet.write_number_with_format(row, col, value, &fmt).map_err(xlsx_err)?;
                    } else {
                        worksheet.write_number(row, col, value).map_err(xlsx_err)?;
                    }
                    continue;
                }
            }

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
