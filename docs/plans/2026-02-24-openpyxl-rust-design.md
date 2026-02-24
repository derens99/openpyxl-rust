# openpyxl_rust — Rust-backed Excel Writer for Python

## Problem

openpyxl is pure Python. Writing large .xlsx files is slow. python-calamine provides fast Rust-based reading but no writing. A Rust-backed writer fills this gap.

## Architecture

```
Python (openpyxl-compatible API)
  └── PyO3 binding layer (src/lib.rs)
        └── rust_xlsxwriter v0.93+ (pure Rust crate by jmcnamara)
              └── .xlsx output (ZIP of XML)
```

- **Rust engine:** `rust_xlsxwriter` — pure Rust, feature-complete, ~3.8x faster than Python XlsxWriter, supports streaming/constant-memory mode
- **Bindings:** PyO3 + maturin
- **Python API:** mirrors openpyxl's write-path interface

## v1 Scope

### Included
- **Data types:** strings, numbers, booleans, dates/datetimes, formulas, None/blank
- **Formatting:** bold, italic, underline, font name/size/color, number formats
- **Structural:** multiple worksheets, column widths, row heights, freeze panes, merged cells
- **I/O:** save to file path, save to BytesIO/buffer

### Not in v1
- Charts, images, conditional formatting, data validation
- Full border/fill/alignment styling
- Reading existing files

## Python API

```python
from openpyxl_rust import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Sales"
ws2 = wb.create_sheet("Details")

ws['A1'] = "Revenue"
ws['B1'] = 42000.50
ws.cell(row=2, column=1, value="Q1")

from openpyxl_rust.styles import Font
ws['A1'].font = Font(bold=True, size=14, name="Arial")
ws['B1'].number_format = '$#,##0.00'

ws.column_dimensions['A'].width = 20
ws.row_dimensions[1].height = 30
ws.freeze_panes = 'A2'
ws.merge_cells('A10:D10')

wb.save("report.xlsx")
```

## Internal Design

Cell data is collected Python-side (dict per worksheet) and flushed to Rust in bulk at `wb.save()`. This avoids per-cell FFI overhead. Format objects translate to `rust_xlsxwriter::Format` at save time.

| Python class | Rust backing | Notes |
|---|---|---|
| Workbook | rust_xlsxwriter::Workbook | Owns all worksheets |
| Worksheet | rust_xlsxwriter::Worksheet | Accessed by index |
| Font | rust_xlsxwriter::Format fields | Translated at save time |
| Cell data | Python-side dict | Flushed to Rust at save |

## Rust Module (PyO3)

Exposes a single function `_save_workbook(data, path_or_none)` that:
1. Receives a Python dict describing all worksheets, cells, formats, and structural config
2. Constructs rust_xlsxwriter objects
3. Writes to file or returns bytes

## Project Structure

```
openpyxl-rust/
├── Cargo.toml
├── pyproject.toml
├── src/
│   └── lib.rs
├── python/
│   └── openpyxl_rust/
│       ├── __init__.py
│       ├── workbook.py
│       ├── worksheet.py
│       ├── cell.py
│       └── styles/
│           ├── __init__.py
│           └── fonts.py
├── tests/
│   ├── test_basic.py
│   └── test_compat.py
└── benchmarks/
    └── bench_vs_openpyxl.py
```

## Benchmark Test

`benchmarks/bench_vs_openpyxl.py` runs identical operations with both openpyxl and openpyxl_rust, comparing wall-clock time:

1. **Large data write:** 100k rows x 10 columns of mixed types (strings, numbers, dates)
2. **Formatted write:** 10k rows with bold headers, number formats, column widths
3. **Multi-sheet:** 5 sheets x 20k rows each

Outputs a comparison table with times and speedup ratios.

## Build & Distribution

- Python 3.9+
- Rust edition 2021
- Dev: `maturin develop`
- Release: `maturin build --release`
