use pyo3::prelude::*;
use pyo3::types::PyBytes;
use rust_xlsxwriter::{Format, Formula, Workbook};
use crate::types::*;
use crate::format::*;
use crate::parse::*;

pub(crate) fn save_workbook(
    sheets: &[SheetData],
    defined_names: &[(String, String)],
    path: Option<&str>,
    py: Python<'_>,
) -> PyResult<PyObject> {
    let mut workbook = Workbook::new();

        for sd in sheets {
            let worksheet = workbook.add_worksheet();
            worksheet.set_name(&sd.title).map_err(xlsx_err)?;

            // Write cells
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

                worksheet.protect_with_options(&opts);
                if let Some(pw) = password {
                    worksheet.protect_with_password(pw);
                }
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
        for (name, formula) in defined_names {
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
