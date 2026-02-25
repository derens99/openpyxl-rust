# Rust-Backed Cells — Performance Optimization Design

## Problem

Profiling the 100k-row benchmark shows:
- **55% (1.72s)**: Python Cell object creation + per-cell FFI calls
- **6% (0.19s)**: JSON format serialization (`_flush_formats`)
- **38% (1.20s)**: Rust xlsx file writing

The Rust save layer is already 4.2x faster than openpyxl's save. The bottleneck is the Python layer: creating `Cell` objects, storing them in `_cells` dicts, making individual PyO3 calls per cell, then re-iterating everything at save time to serialize formats to JSON.

## Solution

Move cell storage entirely into Rust. Python `Cell` becomes a thin proxy that reads/writes directly through PyO3 — no intermediate Python dicts, no JSON serialization, no deferred format flush.

## Architecture

### Rust `CellFormat` struct (replaces `format_json: Option<String>`)

```rust
#[derive(Clone, Debug, Default)]
struct CellFormat {
    font_bold: bool,
    font_italic: bool,
    font_name: Option<String>,
    font_size: Option<f64>,
    font_color: Option<String>,
    font_underline: Option<u8>,    // 0=none, 1=single, 2=double
    font_strikethrough: bool,
    font_vert_align: Option<u8>,   // 1=superscript, 2=subscript
    number_format: Option<String>,
    align_horizontal: Option<u8>,
    align_vertical: Option<u8>,
    align_wrap_text: bool,
    align_shrink_to_fit: bool,
    align_indent: u8,
    align_text_rotation: i16,
    fill_type: Option<u8>,
    fill_start_color: Option<String>,
    fill_end_color: Option<String>,
    border_left_style: Option<u8>,
    border_left_color: Option<String>,
    border_right_style: Option<u8>,
    border_right_color: Option<String>,
    border_top_style: Option<u8>,
    border_top_color: Option<String>,
    border_bottom_style: Option<u8>,
    border_bottom_color: Option<String>,
    border_diagonal_style: Option<u8>,
    border_diagonal_color: Option<String>,
    border_diagonal_up: bool,
    border_diagonal_down: bool,
}
```

### Rust `CellValue` (updated)

```rust
struct CellValue {
    value: CellData,
    format: CellFormat,  // always present, default = no formatting
}
```

### New Rust PyO3 methods

```rust
// Direct format setters — no JSON
fn set_cell_font(&mut self, sheet: usize, row: u32, col: u16,
                 bold: bool, italic: bool, name: Option<String>,
                 size: Option<f64>, color: Option<String>,
                 underline: Option<u8>, strikethrough: bool,
                 vert_align: Option<u8>) -> PyResult<()>;

fn set_cell_alignment(&mut self, sheet: usize, row: u32, col: u16,
                      horizontal: Option<u8>, vertical: Option<u8>,
                      wrap_text: bool, shrink_to_fit: bool,
                      indent: u8, text_rotation: i16) -> PyResult<()>;

fn set_cell_fill(&mut self, sheet: usize, row: u32, col: u16,
                 fill_type: Option<u8>, start_color: Option<String>,
                 end_color: Option<String>) -> PyResult<()>;

fn set_cell_border(&mut self, sheet: usize, row: u32, col: u16,
                   left_style: Option<u8>, left_color: Option<String>,
                   right_style: Option<u8>, right_color: Option<String>,
                   top_style: Option<u8>, top_color: Option<String>,
                   bottom_style: Option<u8>, bottom_color: Option<String>,
                   diag_style: Option<u8>, diag_color: Option<String>,
                   diag_up: bool, diag_down: bool) -> PyResult<()>;

fn set_cell_number_format(&mut self, sheet: usize, row: u32, col: u16,
                          format: String) -> PyResult<()>;
```

### Python `Cell` proxy

```python
class Cell:
    __slots__ = ('_wb', '_sheet_idx', 'row', 'column',
                 '_font', '_alignment', '_border', '_fill',
                 '_number_format', '_hyperlink', '_comment')

    def __init__(self, wb, sheet_idx, row, column, value=None):
        self._wb = wb
        self._sheet_idx = sheet_idx
        self.row = row
        self.column = column
        self._font = None
        self._alignment = None
        self._border = None
        self._fill = None
        self._number_format = "General"
        self._hyperlink = None
        self._comment = None
        if value is not None:
            self.value = value

    @property
    def value(self):
        return self._wb.get_cell_value(self._sheet_idx, self.row - 1, self.column - 1)

    @value.setter
    def value(self, val):
        # Type dispatch + push to Rust in one call
        self._wb.set_cell_value(self._sheet_idx, self.row - 1, self.column - 1, val)

    @property
    def font(self):
        return self._font

    @font.setter
    def font(self, f):
        self._font = f
        if f is not None:
            self._wb.set_cell_font(self._sheet_idx, self.row - 1, self.column - 1,
                                   f.bold, f.italic, f.name, f.size, f.color,
                                   _underline_to_u8(f.underline), f.strikethrough,
                                   _vert_align_to_u8(f.vertAlign))
    # ... similar for alignment, border, fill, number_format
```

### Python `Worksheet` changes

```python
class Worksheet:
    def __init__(self, ...):
        self._cell_cache = {}  # lazy proxy cache

    def cell(self, row, column, value=None):
        key = (row, column)
        if key not in self._cell_cache:
            self._cell_cache[key] = Cell(self._workbook._rust_wb, self._sheet_idx, row, column)
        c = self._cell_cache[key]
        if value is not None:
            c.value = value
        return c

    # __setitem__, __getitem__, append, append_rows stay API-compatible
```

### Save path

```python
def save(self, filename):
    # No _flush_formats needed — formats already in Rust
    for ws in self._sheets:
        ws._flush_metadata()  # only hyperlinks, comments, images, etc.
    self._rust_wb.save(str(filename))
```

### `CellFormat` → `rust_xlsxwriter::Format` conversion

At save time, Rust converts `CellFormat` directly to `rust_xlsxwriter::Format` using field-by-field mapping. No JSON parsing. This replaces `build_format_from_json`.

```rust
fn cell_format_to_xlsx_format(cf: &CellFormat) -> Format {
    let mut fmt = Format::new();
    if cf.font_bold { fmt = fmt.set_bold(); }
    if cf.font_italic { fmt = fmt.set_italic(); }
    // ... direct field mapping, no string parsing
    fmt
}
```

## What Gets Eliminated

| Before | After |
|---|---|
| Python `_cells` dict as primary store | Rust HashMap is primary; Python has lazy cache |
| Per-cell `_set_rust_value()` type dispatch in Python | Single `set_cell_value()` with type dispatch in Rust |
| `_flush_formats()` full iteration at save time | Eliminated — formats already in Rust |
| `json.dumps()` per cell | Eliminated |
| `serde_json::from_str()` per cell at save | Eliminated |
| `build_format_from_json()` | Replaced by direct struct → Format conversion |

## Expected Performance

| Phase | Before | After |
|---|---|---|
| Python populate + FFI | 1.72s | ~0.3s (proxy creation + single Rust call per cell) |
| Format serialization | 0.19s | 0s (eliminated) |
| Rust save | 1.20s | ~1.1s (slightly faster, no JSON parsing) |
| **Total** | **3.1s** | **~1.4s** |
| **Speedup vs openpyxl** | **2.3x** | **~4.9x** |

## Compatibility

API stays identical — all existing user code works unchanged:
- `ws["A1"] = "hello"`
- `ws.cell(row=1, column=1).font = Font(bold=True)`
- `ws.append([1, 2, 3])`
- `ws.append_rows([[1,2], [3,4]])`
- `wb.save("out.xlsx")`

## Bug Fixes Included

1. Microsecond precision for datetime (add `value.microsecond / 1_000_000`)
2. Diagonal border `(false, false)` no longer defaults to `BorderUp`
3. Sheet protection password/options ordering fixed
4. Remove dead `is_date_only` field from `CellData::DateTime`
