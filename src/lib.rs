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
