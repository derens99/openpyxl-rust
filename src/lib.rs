mod format;
mod load;
mod parse;
mod save;
mod types;
mod workbook;

use pyo3::prelude::*;

#[pymodule]
fn _openpyxl_rust(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(load::_load_workbook, m)?)?;
    m.add_function(wrap_pyfunction!(load::_load_workbook_bytes, m)?)?;
    m.add_class::<workbook::RustWorkbook>()?;
    Ok(())
}
