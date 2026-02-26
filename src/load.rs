use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use calamine::{open_workbook, Reader, Xlsx, Data};
use std::io::Cursor;

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
pub(crate) fn _load_workbook(py: Python<'_>, path: &str) -> PyResult<PyObject> {
    let workbook: Xlsx<_> = open_workbook(path)
        .map_err(|e: calamine::XlsxError| pyo3::exceptions::PyIOError::new_err(e.to_string()))?;
    _convert_workbook_to_py(py, workbook)
}

#[pyfunction]
pub(crate) fn _load_workbook_bytes(py: Python<'_>, data: &[u8]) -> PyResult<PyObject> {
    let cursor = Cursor::new(data);
    let workbook: Xlsx<_> = Xlsx::new(cursor)
        .map_err(|e: calamine::XlsxError| pyo3::exceptions::PyIOError::new_err(e.to_string()))?;
    _convert_workbook_to_py(py, workbook)
}
