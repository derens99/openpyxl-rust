# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed
- **save.rs crash on malformed protection JSON** — two `.unwrap()` calls on `as_object()` replaced with proper `PyRuntimeError` propagation (would panic/crash Python on malformed sheet-protection or conditional-format JSON)
- **set_rows_batch inserting Empty entries for None cells** — `CellData::Empty` was stored for every `None` in batch writes, bloating the HashMap; now skips `None` values entirely
- **worksheet `__getitem__` double FFI call** — single-row indexing (`ws[5]`) made two separate `get_dimensions()` calls; merged into one

### Changed
- **`has_format` dirty-bit on CellFormat** — skips expensive `Format::new()` + field conversion for the ~99% of cells with no formatting
- **Hoisted `Format::new()` out of merged-cells loop** — single allocation reused across all merge ranges instead of one per range
- **`std::mem::take` for row/column shift operations** — insert/delete rows/cols now do a single ownership transfer instead of `keys().collect()` + individual `.remove()` calls
- **O(1) dimension update on insert** — insert_rows/insert_cols update cached min/max directly instead of full `recompute_dimensions()` scan
- **`parse_cell_ref` zero-allocation rewrite** — byte-level split without heap `String` allocations
- **`_flush_metadata` single-pass iteration** — merged 3×column + 3×row loops into 1×column + 1×row (widths, hidden, outline levels in one pass)
- **`load_workbook` datetime pre-screening** — fast character check (`value[4]=='-'`) before attempting `strptime` on every string cell

## [0.5.0] - 2025-02-26

### Added
- Tables / ListObjects with style support
- Charts (Bar, Line, Pie, Area, Scatter, Doughnut, Radar) with Series and Reference API
- Conditional formatting (ColorScale, DataBar, IconSet, CellIs, Formula rules)
- Data validation
- Auto-filter
- Sheet protection with password support
- Page setup and print options (margins, orientation, paper size, headers/footers)
- Named ranges (defined_names)
- Row/column insert and delete with automatic cell/merge/hyperlink shifting
- Comments/Notes
- Images (PNG, JPEG) with optional scaling
- Hyperlinks with tooltip support
- `load_workbook()` — fast data-only loading via calamine, openpyxl fallback for formatting

### Changed
- **Rust Owns Everything rewrite** — Rust is now the sole data store; Python Cell is a thin proxy
- Replaced JSON format serialization with `CellFormat` struct and type-specific FFI calls
- Batch operations (`set_rows_batch`, `get_rows_batch`) for single-FFI-call bulk ops
- Dimensions tracked in Rust via `track_cell()` — O(1) reads
- Average **3.5x speedup** over openpyxl (up from ~2x in v0.3)

## [0.3.0] - 2024-12-15

### Added
- DateTime, date, and time support with Excel serial number conversion
- Alignment styling (horizontal, vertical, wrap text, text rotation)
- Border styling (all sides + diagonal, with style and color)
- Pattern fills (solid, gray variants)

### Changed
- Rust save engine handles all cell types, fonts, dimensions, freeze panes, merged cells

## [0.2.0] - 2024-12-01

### Added
- `Workbook.save()` wired to Rust engine via `rust_xlsxwriter`
- Rust save engine with cell types, fonts, column widths, row heights, freeze panes, merged cells
- openpyxl read-back compatibility tests
- Benchmarks comparing openpyxl_rust vs openpyxl

## [0.1.0] - 2024-11-15

### Added
- Initial project scaffolding with PyO3 + maturin
- Cell class with value, coordinate, number_format
- Font class with name, size, bold, italic, underline, color
- Worksheet class with cell access, dimensions, freeze panes, merged cells
- Workbook class with active sheet, create_sheet, getitem

[Unreleased]: https://github.com/derens99/openpyxl-rust/compare/v0.5.0...HEAD
[0.5.0]: https://github.com/derens99/openpyxl-rust/releases/tag/v0.5.0
[0.3.0]: https://github.com/derens99/openpyxl-rust/releases/tag/v0.3.0
[0.2.0]: https://github.com/derens99/openpyxl-rust/releases/tag/v0.2.0
[0.1.0]: https://github.com/derens99/openpyxl-rust/releases/tag/v0.1.0
