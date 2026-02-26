# Rust Owns Everything - Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Rewrite openpyxl-rust so Rust is the sole data store — eliminate the Python `_cell_cache` entirely, make Python `Cell` a proxy, target 5-8x faster than openpyxl.

**Architecture:** Rust `SheetData` owns all cell values and tracks dimensions (min/max row/col). Python `Cell` becomes a `__slots__` proxy with `(_ws, _row, _col)` that reads/writes through FFI. Metadata (font, alignment, hyperlink, comment) lives in a small `_formatted_cells` dict on Worksheet, only populated for the tiny fraction of cells that have formatting.

**Tech Stack:** Rust (PyO3 0.24, rust_xlsxwriter 0.93), Python 3.12, maturin

---

### Task 1: Rust — Add Dimension Tracking + Helpers to SheetData

**Files:**
- Modify: `src/types.rs`

Add `min_row`, `max_row`, `min_col`, `max_col` (all `Option`), and `append_row: u32` to `SheetData`. Add `track_cell(row, col)` method that updates min/max. Add `recompute_dimensions()` that scans all keys (used after row/col delete).

Initialize all to `None`/`0` in `SheetData::new()`.

Compile: `maturin develop --release`

Commit.

---

### Task 2: Rust — Wire Tracking + Add New PyMethods to RustWorkbook

**Files:**
- Modify: `src/workbook.rs`

**2a.** Wire `sd.track_cell(row, col)` into every cell-writing method: `set_cell_string`, `set_cell_number`, `set_cell_boolean`, `set_cell_datetime`, `set_cell_empty`, `set_cell_value`, `set_rows_batch`, and all format setters (`set_cell_font`, `set_cell_alignment`, `set_cell_fill`, `set_cell_border`, `set_cell_number_format`).

**2b.** Add new pymethods:
- `get_dimensions(sheet) -> (Option<u32>, Option<u16>, Option<u32>, Option<u16>)` — returns 0-based min/max
- `touch_cell(sheet, row, col)` — inserts Empty cell if not present, updates dimensions
- `get_next_append_row(sheet) -> u32` — returns `max(append_row, max_row+1)` or 0
- `set_next_append_row(sheet, row)` — sets append_row
- `get_rows_batch(py, sheet, min_row, min_col, max_row, max_col) -> PyList` — bulk read, returns list of lists of Python values. Single FFI call for all cells in a rectangular region.

**2c.** Add row/col structural operations:
- `insert_rows(sheet, idx, amount)` — shift all cells/merges/hyperlinks/notes with row >= idx by +amount, recompute dimensions
- `delete_rows(sheet, idx, amount)` — remove cells in [idx, idx+amount), shift rest, recompute
- `insert_cols(sheet, idx, amount)` — same for columns
- `delete_cols(sheet, idx, amount)` — same for columns

Compile: `maturin develop --release`

Commit.

---

### Task 3: Rewrite Cell as Proxy

**Files:**
- Rewrite: `python/openpyxl_rust/cell.py`

Keep all existing top-level helpers unchanged (`_col_letter`, `_date_to_excel_serial`, all `_*_MAP` dicts, `_underline_to_u8`, `_vert_align_to_u8`).

Replace `Cell` class with `__slots__` proxy:

- Slots: `('_ws', '_row', '_col', '_value', '_font', '_number_format', '_alignment', '_border', '_fill', '_hyperlink', '_comment')`
- `__init__(row=1, column=1, value=None, worksheet=None)` — if worksheet is not None, push value to Rust. If None (standalone), store in `_value`.
- `value` property: if `_ws` is not None, read from `_ws._get_cell_value(row, col)`. Else return `_value`.
- `value` setter: if `_ws` is not None, call `_ws._set_cell_value(row, col, val)`. Else set `_value`.
- `row`/`column` as properties with getter/setter (for insert/delete row/col compat)
- `coordinate`, `data_type` same logic as before
- `font`, `alignment`, `fill`, `border`, `hyperlink`, `comment`, `number_format` — stored locally as slots. Setter calls `_mark_formatted()` which registers `self` in `ws._formatted_cells[(row, col)]`.

Commit.

---

### Task 4: Rewrite Worksheet — Eliminate _cell_cache

**Files:**
- Rewrite: `python/openpyxl_rust/worksheet.py`

**4a. Remove state:**
- Delete `_cell_cache`, `_current_row`, `_max_row`
- Add `_formatted_cells = {}` — maps `(row, col) -> Cell proxy` for cells with any non-default format/hyperlink/comment

**4b. New internal helpers:**
- `_set_cell_value(row, col, value)` — type-dispatch to Rust (same logic as old `cell()` value push: str→set_cell_string, float→set_cell_number, etc.)
- `_get_cell_value(row, col)` — calls `self._workbook._rust_wb.get_cell_value(self._sheet_idx, row-1, col-1)`

**4c. Rewrite `cell(row, column, value=None)`:**
- If `(row, column)` in `_formatted_cells`, return existing proxy (preserves format attrs)
- Else create `Cell(row, column, value, worksheet=self)`. If value is None, call `touch_cell` in Rust.
- Return proxy (not stored anywhere unless formatted)

**4d. Rewrite dimension properties:**
- `min_row`, `max_row`, `min_column`, `max_column` — call `get_dimensions()` from Rust, convert 0-based to 1-based
- `dimensions` — uses the above

**4e. Rewrite `_next_row()`:**
- Return `self._workbook._rust_wb.get_next_append_row(self._sheet_idx) + 1` (0→1-based)

**4f. Rewrite `append(iterable)`:**
- Convert values, handle datetime serials
- Single `set_rows_batch` FFI call with one row
- Set `set_next_append_row` to the 0-based row used
- No Python storage

**4g. Rewrite `append_rows(rows_data)`:**
- Same datetime conversion loop
- Single `set_rows_batch` FFI call
- Update `append_row`
- No Python storage

**4h. Rewrite `iter_rows` / `iter_cols`:**
- `values_only=True`: call `get_rows_batch` from Rust, convert to tuples. For `iter_cols`, transpose.
- `values_only=False`: same but wrap each value in `Cell(row, col, worksheet=self)`. Return existing proxy from `_formatted_cells` if one exists.

**4i. Rewrite `insert_rows/delete_rows/insert_cols/delete_cols`:**
- Call Rust `insert_rows`/`delete_rows`/`insert_cols`/`delete_cols` methods (they handle cell shifting, merge shifting, hyperlink/note shifting, dimension recompute)
- Also shift keys in `_formatted_cells` dict (small set, fast)
- Also shift `merged_cell_ranges` list on Python side (for API compat — some tests check `ws.merged_cell_ranges`)

**4j. Rewrite `_flush_metadata()`:**
- Iterate only `_formatted_cells` for format/hyperlink/comment pushing (instead of all cells)
- Everything else (column widths, row heights, freeze panes, autofilter, protection, page setup, data validations, conditional formatting, images) stays the same

**4k. Rewrite `_resync_rust()`:**
- No longer needed in the same form — Rust already has the data
- For merge operations that remove merges, just call `clear_merge_ranges` + re-push merges

**4l. `__setitem__` / `__getitem__`:**
- Delegate to `cell()` as before

Commit.

---

### Task 5: Update load_workbook (if needed)

**Files:**
- Check: `python/openpyxl_rust/__init__.py`

The `load_workbook` function calls `ws.cell(row=r, column=c, value=v)` per cell. This still works — `cell()` now pushes to Rust via the proxy. Verify it works. It also clears `wb._sheets = []` then re-creates sheets — make sure this still works with the new Worksheet init.

---

### Task 6: Build + Run All Tests + Fix Failures

**Step 1:** `maturin develop --release`

**Step 2:** `py -m pytest --tb=short -q`

**Step 3:** Fix failures. Expected common issues:
- Tests that import `Cell` directly and use it standalone (handled by `_ws is None` branch)
- Tests that check `isinstance(c, Cell)` (still works, class name unchanged)
- Tests that read back values after `append_rows` (now reads from Rust, should be identical)
- Tests for `insert_rows`/`delete_rows` that verify cell positions (now handled by Rust)
- Edge cases with None values in append_rows (handle in Rust `set_rows_batch`)

Target: 265/265 pass (264 + pre-existing image test failure if any).

Commit all fixes.

---

### Task 7: Benchmark + Update MEMORY.md

**Step 1:** `py benchmarks/bench_vs_openpyxl.py`

**Step 2:** Compare:
| Benchmark | Before | Target |
|---|---|---|
| Large data | 2.2x | 4-5x |
| Batch data | 3.1x | 6-8x |
| Formatted | 3.0x | 4-5x |
| Multi-sheet | 3.0x | 4-5x |

**Step 3:** Update `MEMORY.md` with new architecture and numbers.

**Step 4:** Final commit.

---

## Execution Notes

- Tasks 1-2 are pure Rust, compile independently
- Tasks 3-4 are coupled — do them together, then run tests
- The 265 existing tests ARE the regression suite — no new tests needed
- Datetime serial conversion stays in Python (fast, avoids chrono dependency)
- `_formatted_cells` is the ONLY Python-side cell storage, and only for the ~1% of cells with formatting
- The `_resync_rust()` method can be greatly simplified since Rust already owns the data
