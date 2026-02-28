# Feature Parity Test Suite Design

**Date:** 2026-02-27
**Goal:** Create a comprehensive 1-to-1 test suite comparing every openpyxl feature against openpyxl-rust.

## Strategy

- **Single file:** `tests/test_parity.py`
- **Roundtrip verification:** Write with openpyxl_rust → save → read back with real openpyxl → assert values match
- **Unsupported features:** Marked `@pytest.mark.skip(reason="not yet implemented: <feature>")` — living checklist
- **Organization:** One test class per openpyxl module/category

## Helpers

```python
def _rust_write_and_read(setup_fn, tmp_path):
    """Write with openpyxl_rust, read back with real openpyxl."""

def _both_write_and_compare(rust_setup_fn, openpyxl_setup_fn, tmp_path, compare_fn):
    """Write with both libraries, read both back, compare outputs."""
```

## Test Classes (29 categories, ~180 tests)

### Core (always active)

| Class | Count | What it tests |
|---|---|---|
| TestWorkbookParity | 15 | create, save, load, sheetnames, active, `__getitem__`, `__iter__`, `__len__`, `__contains__` |
| TestSheetManagementParity | 10 | create_sheet, create_sheet with index, remove, rename |
| TestCellAccessParity | 10 | `ws["A1"]`, `ws.cell()`, `ws["A1:C3"]`, `ws["A"]`, `ws[1]`, slicing |
| TestCellDataTypesParity | 15 | string, int, float, bool, None, formula, date, time, datetime, error |
| TestIteratorsParity | 10 | iter_rows, iter_cols, values, append, append with dict |
| TestCellStylesParity | 20 | Font (all params), PatternFill, Border/Side, Alignment, number_format |
| TestRowColumnDimensionsParity | 10 | width, height, custom dimensions |
| TestMergedCellsParity | 6 | merge, unmerge, MergedCell properties |
| TestHyperlinksParity | 5 | external URL, internal ref, tooltip |
| TestCommentsParity | 5 | create, author, height/width |
| TestDataValidationParity | 12 | list, whole, decimal, textLength, custom, operators, messages |
| TestConditionalFormattingParity | 12 | CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule |
| TestAutoFilterParity | 5 | basic filter, filter with data |
| TestTablesParity | 5 | Table, TableColumn, TableStyleInfo |
| TestDefinedNamesParity | 5 | global named range, constants |
| TestPrintSetupParity | 10 | margins, orientation, paper_size, print_area, print_titles |
| TestSheetProtectionParity | 5 | enable, password, options |
| TestSheetViewsParity | 6 | freeze_panes, zoom (skip) |
| TestImagesParity | 4 | single image, anchor, dimensions |
| TestChartsParity | 18 | Bar, Line, Pie, Area, Scatter, Doughnut, Radar, Stock, 3D, axes, legend |
| TestFormulasParity | 4 | string formulas, preservation |
| TestDateTimeParity | 6 | serial conversion, format detection |

### Skipped (not yet implemented — living checklist)

| Class | Count | Missing feature |
|---|---|---|
| TestSheetManagementParity (partial) | 3 | copy_worksheet, move_sheet, sheet_state |
| TestCellDataTypesParity (partial) | 1 | CellRichText |
| TestCellStylesParity (partial) | 3 | GradientFill, NamedStyle, cell Protection |
| TestRowColumnDimensionsParity (partial) | 3 | hidden, outline_level, group |
| TestConditionalFormattingParity (partial) | 3 | top10, duplicates, text rules |
| TestHeaderFooterParity | 4 | odd/even/first headers and footers |
| TestPageBreaksParity | 3 | row breaks, col breaks |
| TestWorkbookProtectionParity | 3 | lock structure, lock windows |
| TestSheetViewsParity (partial) | 2 | zoom, split panes |
| TestMoveRangeParity | 3 | shift cells, translate formulas |
| TestDocumentPropertiesParity | 3 | title, author, created/modified |
| TestRichTextParity | 3 | CellRichText, TextBlock, InlineFont |
| TestPivotTablesParity | 2 | read/preserve |
| TestChartsParity (partial) | 3 | trendlines, data labels, error bars |

**Totals: ~180 tests, ~145 active, ~35 skipped**

## Test Pattern

```python
class TestCellDataTypesParity:
    def test_string_value(self, tmp_path):
        path = tmp_path / "test.xlsx"
        wb = rust_Workbook()
        wb.active["A1"] = "Hello World"
        wb.save(str(path))

        rb = real_openpyxl.load_workbook(str(path))
        assert rb.active["A1"].value == "Hello World"
        assert rb.active["A1"].data_type == "s"

    @pytest.mark.skip(reason="not yet implemented: CellRichText")
    def test_rich_text(self, tmp_path):
        """openpyxl supports CellRichText for mixed formatting in a cell."""
        path = tmp_path / "test.xlsx"
        wb = rust_Workbook()
        # Would need: from openpyxl_rust.cell.rich_text import CellRichText, TextBlock
        # wb.active["A1"] = CellRichText("Hello ", TextBlock(Font(bold=True), "World"))
        wb.save(str(path))
        rb = real_openpyxl.load_workbook(str(path))
        # assert isinstance(rb.active["A1"].value, str)
```

## Naming Convention

- Test names: `test_<feature>` for supported, same name with skip marker for unsupported
- All skip reasons start with `"not yet implemented: "` for easy grep
- `grep -c "not yet implemented" tests/test_parity.py` gives the gap count

## Dependencies

- `openpyxl` (already in dev dependencies for roundtrip verification)
- `pytest` (already used)
- `Pillow` (for image tests, likely already available)
