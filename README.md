# openpyxl_rust

Fast, Rust-backed Excel (.xlsx) writer with an openpyxl-compatible Python API. Typically **2-3x faster** than openpyxl for write operations.

## Installation

```bash
# Development (requires Rust toolchain + maturin)
pip install -e .

# Or build a wheel
maturin build --release
pip install target/wheels/*.whl
```

## Quick Start

```python
from openpyxl_rust import Workbook
from openpyxl_rust.styles import Font, Alignment, Border, Side, PatternFill

wb = Workbook()
ws = wb.active
ws.title = "Sales"

# Write data
ws["A1"] = "Product"
ws["B1"] = "Revenue"
ws["A1"].font = Font(bold=True, size=14)

for i in range(2, 1002):
    ws.cell(row=i, column=1, value=f"Item {i}")
    ws.cell(row=i, column=2, value=i * 99.5)
    ws.cell(row=i, column=2).number_format = "$#,##0.00"

# Batch append for best performance
ws2 = wb.create_sheet("Batch")
rows = [[r * c for c in range(1, 11)] for r in range(1, 10001)]
ws2.append_rows(rows)

wb.save("report.xlsx")
```

## Features

| Feature | Status |
|---------|--------|
| Strings, numbers, booleans, formulas | Supported |
| Dates and datetimes | Supported |
| Font styling (bold, italic, underline, strikethrough, color, size) | Supported |
| Alignment (horizontal, vertical, wrap, rotation) | Supported |
| Borders (all sides + diagonal) | Supported |
| Pattern fills (solid, gray variants) | Supported |
| Number formats | Supported |
| Multiple worksheets | Supported |
| Column widths and row heights | Supported |
| Freeze panes | Supported |
| Merged cells | Supported |
| Hyperlinks | Supported |
| Comments/Notes | Supported |
| Images (PNG, JPEG) | Supported |
| Data validation | Supported |
| Conditional formatting | Supported |
| Auto-filter | Supported |
| Sheet protection | Supported |
| Page setup and print options | Supported |
| Named ranges | Supported |
| Row/column insert and delete | Supported |
| iter_rows / iter_cols / dimensions | Supported |
| Batch row append (append_rows) | Supported |
| Save to file or BytesIO | Supported |
| Load workbook (data only) | Supported |
| Gradient fills | Not supported |
| Named styles | Not supported |
| Charts | Not supported |
| VBA macros | Not supported |
| Load with formatting | Not supported |

## Performance

Benchmarks on 100k rows x 10 columns (mixed types):

| Library | Time | Speedup |
|---------|------|---------|
| openpyxl | ~7s | 1x |
| openpyxl_rust | ~3s | ~2.4x |
| openpyxl_rust (batch) | ~2.9s | ~2.4x |

Multi-sheet workbooks (5 sheets x 20k rows) see up to **3x** speedup.

## How It Works

Python classes (`Workbook`, `Worksheet`, `Cell`) mirror openpyxl's API. Cell data is stored in Rust memory via PyO3. At save time, all formatting is serialized to JSON and the Rust engine (`rust_xlsxwriter`) writes the .xlsx file. Reading uses `calamine` for fast parsing.

## License

MIT
