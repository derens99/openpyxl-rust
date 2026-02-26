use pyo3::prelude::*;
use rust_xlsxwriter::XlsxError;

pub(crate) fn xlsx_err(e: XlsxError) -> PyErr {
    pyo3::exceptions::PyRuntimeError::new_err(e.to_string())
}

/// Map a conditional format type string to ConditionalFormatType.
pub(crate) fn parse_cf_type(s: &str) -> rust_xlsxwriter::ConditionalFormatType {
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
pub(crate) fn col_letters_to_index(letters: &str) -> Option<u16> {
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
pub(crate) fn parse_cell_ref(s: &str) -> Option<(u32, u16)> {
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
pub(crate) fn parse_cell_range(s: &str) -> Option<(u32, u16, u32, u16)> {
    let parts: Vec<&str> = s.split(':').collect();
    if parts.len() != 2 { return None; }
    let (r1, c1) = parse_cell_ref(parts[0])?;
    let (r2, c2) = parse_cell_ref(parts[1])?;
    Some((r1, c1, r2, c2))
}

/// Parse a row range like "1:3" to (first_row_0based, last_row_0based).
pub(crate) fn parse_row_range(s: &str) -> Option<(u32, u32)> {
    let parts: Vec<&str> = s.split(':').collect();
    if parts.len() != 2 { return None; }
    let first: u32 = parts[0].trim().parse().ok()?;
    let last: u32 = parts[1].trim().parse().ok()?;
    if first == 0 || last == 0 { return None; }
    Some((first - 1, last - 1))
}

/// Parse a column range like "A:B" to (first_col_0based, last_col_0based).
pub(crate) fn parse_col_range(s: &str) -> Option<(u16, u16)> {
    let parts: Vec<&str> = s.split(':').collect();
    if parts.len() != 2 { return None; }
    let first = col_letters_to_index(parts[0].trim())?;
    let last = col_letters_to_index(parts[1].trim())?;
    Some((first, last))
}
