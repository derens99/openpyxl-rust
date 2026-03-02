import json
import re
from datetime import date, datetime, time

from openpyxl_rust.cell import (
    _BORDER_STYLE_MAP,
    _FILL_TYPE_MAP,
    _HALIGN_MAP,
    _VALIGN_MAP,
    Cell,
    _col_letter,
    _date_to_excel_serial,
    _underline_to_u8,
    _vert_align_to_u8,
)
from openpyxl_rust.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    DuplicateRule,
    FormulaRule,
    IconSetRule,
    TextRule,
    Top10Rule,
)
from openpyxl_rust.header_footer import HeaderFooter
from openpyxl_rust.page import PageMargins, PrintOptions, PrintPageSetup
from openpyxl_rust.page_break import BreakList
from openpyxl_rust.protection import SheetProtection


def _parse_cell_ref(ref_str):
    """Parse 'A1' -> (row, col) as 1-based integers."""
    m = re.match(r"^([A-Z]+)(\d+)$", ref_str.upper())
    if not m:
        raise ValueError(f"Invalid cell reference: {ref_str}")
    letters, row_str = m.group(1), m.group(2)
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - 64)
    return int(row_str), col


class ColumnDimension:
    def __init__(self):
        self.width = None
        self.hidden = False
        self.outline_level = 0
        self.bestFit = False


class RowDimension:
    def __init__(self):
        self.height = None
        self.hidden = False
        self.outline_level = 0


class _ColumnDimensionsDict:
    def __init__(self):
        self._dims = {}

    def __getitem__(self, key):
        if key not in self._dims:
            self._dims[key] = ColumnDimension()
        return self._dims[key]

    def items(self):
        return self._dims.items()


class _RowDimensionsDict:
    def __init__(self):
        self._dims = {}

    def __getitem__(self, key):
        if key not in self._dims:
            self._dims[key] = RowDimension()
        return self._dims[key]

    def items(self):
        return self._dims.items()


class _AutoFilter:
    def __init__(self):
        self._ref = None
        self._filter_columns = []

    @property
    def ref(self):
        return self._ref

    @ref.setter
    def ref(self, value):
        self._ref = value

    def add_filter_column(self, col_id, vals, blank=False):
        """Add a filter for a column. col_id is 0-based column index."""
        self._filter_columns.append({"col": col_id, "values": list(vals)})


class _ConditionalFormattingList:
    def __init__(self):
        self._rules = []

    def add(self, range_string, rule):
        self._rules.append((range_string, rule))


class Worksheet:
    ORIENTATION_PORTRAIT = "portrait"
    ORIENTATION_LANDSCAPE = "landscape"
    PAPERSIZE_LETTER = 1
    PAPERSIZE_LEGAL = 5
    PAPERSIZE_A4 = 9
    PAPERSIZE_A5 = 11

    def __init__(self, title="Sheet", workbook=None, sheet_idx=None):
        self._workbook = workbook
        self._sheet_idx = sheet_idx
        self._title = title
        self._formatted_cells = {}  # (row, col) -> Cell proxy, only for formatted cells
        self.column_dimensions = _ColumnDimensionsDict()
        self.row_dimensions = _RowDimensionsDict()
        self.freeze_panes = None
        self.merged_cell_ranges = []
        self.auto_filter = _AutoFilter()
        self.protection = SheetProtection()
        self.page_setup = PrintPageSetup()
        self.page_margins = PageMargins()
        self.print_options = PrintOptions()
        self.print_area = None
        self.print_title_rows = None
        self.print_title_cols = None
        self._images = []
        self._data_validations = []
        self.conditional_formatting = _ConditionalFormattingList()
        self._tables = []
        self._charts = []
        self._sheet_state = "visible"
        self._zoom_scale = None
        self._show_gridlines = True
        self._autofit = False
        self.row_breaks = BreakList()
        self.col_breaks = BreakList()
        self.oddHeader = HeaderFooter()
        self.oddFooter = HeaderFooter()

        if workbook is not None and sheet_idx is not None:
            workbook._rust_wb.set_sheet_title(sheet_idx, title)

    @property
    def title(self):
        return self._title

    @title.setter
    def title(self, value):
        if self._workbook is not None:
            for ws in self._workbook._sheets:
                if ws is not self and ws.title == value:
                    raise ValueError(f"Worksheet title '{value}' already exists. Use a unique title.")
        self._title = value
        if self._workbook is not None and self._sheet_idx is not None:
            self._workbook._rust_wb.set_sheet_title(self._sheet_idx, value)

    @property
    def sheet_state(self):
        return self._sheet_state

    @sheet_state.setter
    def sheet_state(self, value):
        if value not in ("visible", "hidden", "veryHidden"):
            raise ValueError(f"Invalid sheet state: {value}")
        self._sheet_state = value

    @property
    def zoom(self):
        return self._zoom_scale

    @zoom.setter
    def zoom(self, value):
        self._zoom_scale = value

    def auto_fit_columns(self):
        """Auto-size all columns to fit their content."""
        self._autofit = True

    # ---- Dimensions (from Rust, O(1)) ----

    def _get_dims(self):
        """Returns (min_row, min_col, max_row, max_col) as 1-based, or Nones."""
        if self._workbook is None:
            return None, None, None, None
        dims = self._workbook._rust_wb.get_dimensions(self._sheet_idx)
        if dims[0] is None:
            return None, None, None, None
        return dims[0] + 1, dims[1] + 1, dims[2] + 1, dims[3] + 1

    @property
    def min_row(self):
        return self._get_dims()[0]

    @property
    def max_row(self):
        return self._get_dims()[2]

    @property
    def min_column(self):
        return self._get_dims()[1]

    @property
    def max_column(self):
        return self._get_dims()[3]

    @property
    def dimensions(self):
        mn_r, mn_c, mx_r, mx_c = self._get_dims()
        if mn_r is None:
            return ""
        return f"{_col_letter(mn_c)}{mn_r}:{_col_letter(mx_c)}{mx_r}"

    # ---- Cell value helpers (Python <-> Rust) ----

    def _get_cell_value(self, row, col):
        """Read a cell value from Rust. row/col are 1-based."""
        if self._workbook is None:
            return None
        val = self._workbook._rust_wb.get_cell_value(self._sheet_idx, row - 1, col - 1)
        return val

    def _set_cell_value(self, row, col, value):
        """Push a cell value to Rust. row/col are 1-based."""
        if self._workbook is None:
            return
        t = type(value)
        r0 = row - 1
        c0 = col - 1
        wb = self._workbook._rust_wb
        idx = self._sheet_idx
        if t is str:
            wb.set_cell_string(idx, r0, c0, value)
        elif t is float:
            wb.set_cell_number(idx, r0, c0, value)
        elif t is int:
            wb.set_cell_number(idx, r0, c0, float(value))
        elif t is bool:
            wb.set_cell_boolean(idx, r0, c0, value)
        elif t is datetime:
            serial = _date_to_excel_serial(value.year, value.month, value.day)
            serial += (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86400.0
            wb.set_cell_datetime(idx, r0, c0, serial, 2)
            wb.set_cell_number_format(idx, r0, c0, "yyyy-mm-dd hh:mm:ss")
        elif t is date:
            serial = _date_to_excel_serial(value.year, value.month, value.day)
            wb.set_cell_datetime(idx, r0, c0, serial, 0)
            wb.set_cell_number_format(idx, r0, c0, "yyyy-mm-dd")
        elif t is time:
            serial = (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86400.0
            wb.set_cell_datetime(idx, r0, c0, serial, 1)
            wb.set_cell_number_format(idx, r0, c0, "hh:mm:ss")
        elif value is None:
            wb.set_cell_empty(idx, r0, c0)
        else:
            # Check for CellRichText (lazy import to avoid circular deps)
            from openpyxl_rust.rich_text import CellRichText

            if isinstance(value, CellRichText):
                import json as _json

                wb.set_cell_rich_text(idx, r0, c0, _json.dumps(value._to_json_segments()))
            else:
                raise TypeError(
                    f"Unsupported cell value type: {type(value).__name__}. "
                    f"Supported types: str, int, float, bool, datetime, date, time, None, CellRichText"
                )

    # ---- Core cell() method ----

    def cell(self, row, column, value=None):
        key = (row, column)
        # Return existing proxy if formatted (preserves format attrs)
        if key in self._formatted_cells:
            c = self._formatted_cells[key]
            if value is not None:
                c.value = value
            return c
        # Create new proxy
        c = Cell(row=row, column=column, value=value, worksheet=self)
        if self._workbook is None:
            # No Rust backend — cache cell locally so it persists
            self._formatted_cells[key] = c
        elif value is None:
            # Touch cell in Rust for dimension tracking
            self._workbook._rust_wb.touch_cell(self._sheet_idx, row - 1, column - 1)
        return c

    # ---- Iteration ----

    def iter_rows(self, min_row=None, max_row=None, min_col=None, max_col=None, values_only=False):
        mn_r, mn_c, mx_r, mx_c = self._get_dims()
        min_row = min_row if min_row is not None else mn_r
        max_row = max_row if max_row is not None else mx_r
        min_col = min_col if min_col is not None else mn_c
        max_col = max_col if max_col is not None else mx_c
        if any(v is None for v in (min_row, max_row, min_col, max_col)):
            return
        if values_only and self._workbook is not None:
            # Batch read from Rust — single FFI call
            rows = self._workbook._rust_wb.get_rows_batch(
                self._sheet_idx, min_row - 1, min_col - 1, max_row - 1, max_col - 1
            )
            for row_data in rows:
                yield tuple(None if v is None else v for v in row_data)
        else:
            for row in range(min_row, max_row + 1):
                if values_only:
                    yield tuple(self._get_cell_value(row, col) for col in range(min_col, max_col + 1))
                else:
                    yield tuple(self.cell(row=row, column=col) for col in range(min_col, max_col + 1))

    def iter_cols(self, min_col=None, max_col=None, min_row=None, max_row=None, values_only=False):
        mn_r, mn_c, mx_r, mx_c = self._get_dims()
        min_row = min_row if min_row is not None else mn_r
        max_row = max_row if max_row is not None else mx_r
        min_col = min_col if min_col is not None else mn_c
        max_col = max_col if max_col is not None else mx_c
        if any(v is None for v in (min_row, max_row, min_col, max_col)):
            return
        if values_only and self._workbook is not None:
            # Batch read, then transpose to column-major
            rows = self._workbook._rust_wb.get_rows_batch(
                self._sheet_idx, min_row - 1, min_col - 1, max_row - 1, max_col - 1
            )
            num_cols = max_col - min_col + 1
            for ci in range(num_cols):
                yield tuple(None if rows[ri][ci] is None else rows[ri][ci] for ri in range(len(rows)))
        else:
            for col in range(min_col, max_col + 1):
                if values_only:
                    yield tuple(self._get_cell_value(row, col) for row in range(min_row, max_row + 1))
                else:
                    yield tuple(self.cell(row=row, column=col) for row in range(min_row, max_row + 1))

    @property
    def values(self):
        return self.iter_rows(values_only=True)

    # ---- Item access ----

    def __setitem__(self, key, value):
        row, col = _parse_cell_ref(key)
        self.cell(row=row, column=col, value=value)

    def __getitem__(self, key):
        if isinstance(key, int):
            # Single FFI call for both dimensions instead of two property accesses
            _, mn_c, _, mx_c = self._get_dims()
            min_col = mn_c or 1
            max_col = mx_c or 1
            return tuple(self.cell(row=key, column=col) for col in range(min_col, max_col + 1))
        if ":" in key:
            start_ref, end_ref = key.split(":")
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            rows = []
            for row in range(r1, r2 + 1):
                rows.append(tuple(self.cell(row=row, column=col) for col in range(c1, c2 + 1)))
            return tuple(rows)
        row, col = _parse_cell_ref(key)
        return self.cell(row=row, column=col)

    # ---- Append ----

    def _next_row(self):
        if self._workbook is None:
            return 1
        return self._workbook._rust_wb.get_next_append_row(self._sheet_idx) + 1  # 0->1-based

    def append(self, iterable):
        row = self._next_row()
        values = list(iterable)

        if self._workbook is not None and self._sheet_idx is not None:
            converted = []
            datetime_cells = []  # (col_idx, fmt_str, kind, serial)
            for col_idx, value in enumerate(values):
                if value is None:
                    converted.append(None)
                else:
                    t = type(value)
                    if t is datetime:
                        serial = _date_to_excel_serial(value.year, value.month, value.day)
                        serial += (
                            value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000
                        ) / 86400.0
                        converted.append(serial)
                        datetime_cells.append((col_idx, "yyyy-mm-dd hh:mm:ss", 2, serial))
                    elif t is date:
                        serial = _date_to_excel_serial(value.year, value.month, value.day)
                        converted.append(serial)
                        datetime_cells.append((col_idx, "yyyy-mm-dd", 0, serial))
                    elif t is time:
                        serial = (
                            value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000
                        ) / 86400.0
                        converted.append(serial)
                        datetime_cells.append((col_idx, "hh:mm:ss", 1, serial))
                    elif t is int:
                        converted.append(float(value))
                    else:
                        converted.append(value)

            r0 = row - 1
            self._workbook._rust_wb.set_rows_batch(self._sheet_idx, r0, [converted])
            wb = self._workbook._rust_wb
            for col_idx, fmt_str, kind, serial in datetime_cells:
                wb.set_cell_datetime(self._sheet_idx, r0, col_idx, serial, kind)
                wb.set_cell_number_format(self._sheet_idx, r0, col_idx, fmt_str)
            wb.set_next_append_row(self._sheet_idx, r0 + 1)

    def append_rows(self, rows_data):
        """Append multiple rows at once via the Rust batch API."""
        rows_data = list(rows_data)
        if not rows_data:
            return
        start_row = self._next_row()
        start_row_0based = start_row - 1

        datetime_cells = []  # (r0, c0, fmt_str, kind, serial)

        if self._workbook is not None and self._sheet_idx is not None:
            rows_list = []
            for row_offset, row in enumerate(rows_data):
                converted_row = []
                for col_offset, value in enumerate(row):
                    if value is None:
                        converted_row.append(None)
                    else:
                        t = type(value)
                        if t is datetime:
                            serial = _date_to_excel_serial(value.year, value.month, value.day)
                            serial += (
                                value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000
                            ) / 86400.0
                            converted_row.append(serial)
                            datetime_cells.append(
                                (start_row_0based + row_offset, col_offset, "yyyy-mm-dd hh:mm:ss", 2, serial)
                            )
                        elif t is date:
                            serial = _date_to_excel_serial(value.year, value.month, value.day)
                            converted_row.append(serial)
                            datetime_cells.append((start_row_0based + row_offset, col_offset, "yyyy-mm-dd", 0, serial))
                        elif t is time:
                            serial = (
                                value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000
                            ) / 86400.0
                            converted_row.append(serial)
                            datetime_cells.append((start_row_0based + row_offset, col_offset, "hh:mm:ss", 1, serial))
                        elif t is int:
                            converted_row.append(float(value))
                        else:
                            converted_row.append(value)
                rows_list.append(converted_row)
            self._workbook._rust_wb.set_rows_batch(self._sheet_idx, start_row_0based, rows_list)
            wb = self._workbook._rust_wb
            for r0, c0, fmt_str, kind, serial in datetime_cells:
                wb.set_cell_datetime(self._sheet_idx, r0, c0, serial, kind)
                wb.set_cell_number_format(self._sheet_idx, r0, c0, fmt_str)
            # Update append cursor
            last_row_0based = start_row_0based + len(rows_data)
            wb.set_next_append_row(self._sheet_idx, last_row_0based)

    # ---- Merge cells ----

    def merge_cells(self, range_string):
        parts = range_string.split(":")
        if len(parts) != 2:
            raise ValueError(f"Invalid merge range: {range_string}")
        start_ref = parts[0].upper()
        end_ref = parts[1].upper()
        self.merged_cell_ranges.append((start_ref, end_ref))
        if self._workbook is not None and self._sheet_idx is not None:
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            self._workbook._rust_wb.add_merge_range(self._sheet_idx, r1 - 1, c1 - 1, r2 - 1, c2 - 1)

    def unmerge_cells(self, range_string):
        parts = range_string.split(":")
        if len(parts) != 2:
            raise ValueError(f"Invalid merge range: {range_string}")
        start_ref = parts[0].upper()
        end_ref = parts[1].upper()
        target = (start_ref, end_ref)
        if target not in self.merged_cell_ranges:
            raise ValueError(f"Cell range {range_string} is not merged")
        self.merged_cell_ranges.remove(target)
        if self._workbook is not None and self._sheet_idx is not None:
            wb = self._workbook._rust_wb
            wb.clear_merge_ranges(self._sheet_idx)
            for s, e in self.merged_cell_ranges:
                r1, c1 = _parse_cell_ref(s)
                r2, c2 = _parse_cell_ref(e)
                wb.add_merge_range(self._sheet_idx, r1 - 1, c1 - 1, r2 - 1, c2 - 1)

    # ---- Row/Col insert/delete ----

    def insert_rows(self, idx, amount=1):
        if self._workbook is not None:
            self._workbook._rust_wb.rust_insert_rows(self._sheet_idx, idx - 1, amount)
        # Shift formatted cells
        new_fc = {}
        for (r, c), cell in self._formatted_cells.items():
            if r >= idx:
                cell._row = r + amount
                new_fc[(r + amount, c)] = cell
            else:
                new_fc[(r, c)] = cell
        self._formatted_cells = new_fc
        # Shift merged_cell_ranges (Python-side list for API compat)
        new_merged = []
        for start_ref, end_ref in self.merged_cell_ranges:
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            if r1 >= idx:
                r1 += amount
            if r2 >= idx:
                r2 += amount
            new_merged.append((_col_letter(c1) + str(r1), _col_letter(c2) + str(r2)))
        self.merged_cell_ranges = new_merged

    def delete_rows(self, idx, amount=1):
        if self._workbook is not None:
            self._workbook._rust_wb.rust_delete_rows(self._sheet_idx, idx - 1, amount)
        new_fc = {}
        for (r, c), cell in self._formatted_cells.items():
            if idx <= r < idx + amount:
                continue
            elif r >= idx + amount:
                cell._row = r - amount
                new_fc[(r - amount, c)] = cell
            else:
                new_fc[(r, c)] = cell
        self._formatted_cells = new_fc
        new_merged = []
        for start_ref, end_ref in self.merged_cell_ranges:
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            if r1 >= idx and r2 < idx + amount:
                continue
            if r1 >= idx + amount:
                r1 -= amount
            if r2 >= idx + amount:
                r2 -= amount
            new_merged.append((_col_letter(c1) + str(r1), _col_letter(c2) + str(r2)))
        self.merged_cell_ranges = new_merged

    def insert_cols(self, idx, amount=1):
        if self._workbook is not None:
            self._workbook._rust_wb.rust_insert_cols(self._sheet_idx, idx - 1, amount)
        new_fc = {}
        for (r, c), cell in self._formatted_cells.items():
            if c >= idx:
                cell._col = c + amount
                new_fc[(r, c + amount)] = cell
            else:
                new_fc[(r, c)] = cell
        self._formatted_cells = new_fc
        new_merged = []
        for start_ref, end_ref in self.merged_cell_ranges:
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            if c1 >= idx:
                c1 += amount
            if c2 >= idx:
                c2 += amount
            new_merged.append((_col_letter(c1) + str(r1), _col_letter(c2) + str(r2)))
        self.merged_cell_ranges = new_merged

    def delete_cols(self, idx, amount=1):
        if self._workbook is not None:
            self._workbook._rust_wb.rust_delete_cols(self._sheet_idx, idx - 1, amount)
        new_fc = {}
        for (r, c), cell in self._formatted_cells.items():
            if idx <= c < idx + amount:
                continue
            elif c >= idx + amount:
                cell._col = c - amount
                new_fc[(r, c - amount)] = cell
            else:
                new_fc[(r, c)] = cell
        self._formatted_cells = new_fc
        new_merged = []
        for start_ref, end_ref in self.merged_cell_ranges:
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            if c1 >= idx and c2 < idx + amount:
                continue
            if c1 >= idx + amount:
                c1 -= amount
            if c2 >= idx + amount:
                c2 -= amount
            new_merged.append((_col_letter(c1) + str(r1), _col_letter(c2) + str(r2)))
        self.merged_cell_ranges = new_merged

    # ---- Images / Validation ----

    def add_image(self, img, anchor=None):
        if anchor is not None:
            img.anchor = anchor
        self._images.append(img)

    def add_data_validation(self, dv):
        self._data_validations.append(dv)

    def add_table(self, table):
        self._tables.append(table)

    def add_chart(self, chart, anchor=None):
        if anchor is not None:
            chart._anchor = anchor
        self._charts.append(chart)

    # ---- Flush formats to Rust at save time ----

    def _flush_formats_to_rust(self):
        """Push cell formats to Rust. Only iterates _formatted_cells (tiny set)."""
        if self._workbook is None or self._sheet_idx is None:
            return
        wb = self._workbook._rust_wb
        idx = self._sheet_idx

        for (row, col), cell in self._formatted_cells.items():
            r0, c0 = row - 1, col - 1

            if cell.number_format and cell.number_format != "General":
                wb.set_cell_number_format(idx, r0, c0, cell.number_format)

            if cell.font is not None:
                f = cell.font
                wb.set_cell_font(
                    idx,
                    r0,
                    c0,
                    f.bold,
                    f.italic,
                    f.name,
                    float(f.size) if f.size is not None else None,
                    f.color,
                    _underline_to_u8(f.underline),
                    f.strikethrough,
                    _vert_align_to_u8(f.vertAlign),
                )

            if cell.alignment is not None:
                a = cell.alignment
                wb.set_cell_alignment(
                    idx,
                    r0,
                    c0,
                    _HALIGN_MAP.get(a.horizontal) if a.horizontal else None,
                    _VALIGN_MAP.get(a.vertical) if a.vertical else None,
                    bool(a.wrap_text),
                    bool(a.shrink_to_fit),
                    int(a.indent) if a.indent else 0,
                    int(a.text_rotation) if a.text_rotation else 0,
                )

            if cell.fill is not None:
                fi = cell.fill
                wb.set_cell_fill(
                    idx,
                    r0,
                    c0,
                    _FILL_TYPE_MAP.get(fi.fill_type) if fi.fill_type else None,
                    fi.start_color,
                    fi.end_color,
                )

            if cell.border is not None:
                b = cell.border
                wb.set_cell_border(
                    idx,
                    r0,
                    c0,
                    _BORDER_STYLE_MAP.get(b.left.style) if b.left.style else None,
                    b.left.color,
                    _BORDER_STYLE_MAP.get(b.right.style) if b.right.style else None,
                    b.right.color,
                    _BORDER_STYLE_MAP.get(b.top.style) if b.top.style else None,
                    b.top.color,
                    _BORDER_STYLE_MAP.get(b.bottom.style) if b.bottom.style else None,
                    b.bottom.color,
                    _BORDER_STYLE_MAP.get(b.diagonal.style) if b.diagonal.style else None,
                    b.diagonal.color,
                    bool(b.diagonalUp),
                    bool(b.diagonalDown),
                )

            if cell.protection is not None:
                wb.set_cell_protection(
                    idx,
                    r0,
                    c0,
                    cell.protection.locked if cell.protection.locked is not None else None,
                    cell.protection.hidden if cell.protection.hidden is not None else None,
                )

            if cell.hyperlink is not None:
                url = cell.hyperlink
                text = None
                tooltip = None
                if isinstance(url, str) and url.startswith("#"):
                    url = "internal:" + url[1:]
                wb.add_hyperlink(idx, r0, c0, url, text, tooltip)

            if cell.comment is not None:
                author = cell.comment.author if cell.comment.author else None
                wb.add_note(idx, r0, c0, cell.comment.text, author)

    def _flush_metadata(self):
        """Push all metadata to Rust. Called right before save."""
        if self._workbook is None or self._sheet_idx is None:
            return
        wb = self._workbook._rust_wb
        idx = self._sheet_idx

        # Formats + hyperlinks + comments (only formatted cells)
        self._flush_formats_to_rust()

        # Column dimensions — single pass (widths + hidden + outline levels)
        for letter, dim in self.column_dimensions.items():
            if dim.width is not None or dim.hidden or (dim.outline_level and dim.outline_level > 0):
                _, col_idx = _parse_cell_ref(f"{letter}1")
                col_0 = col_idx - 1
                if dim.width is not None:
                    wb.set_column_width(idx, col_0, dim.width)
                if dim.hidden:
                    wb.set_col_hidden(idx, col_0)
                if dim.outline_level and dim.outline_level > 0:
                    wb.set_col_outline_level(idx, col_0, dim.outline_level)

        # Row dimensions — single pass (heights + hidden + outline levels)
        for row_num, dim in self.row_dimensions.items():
            row_0 = row_num - 1
            if dim.height is not None:
                wb.set_row_height(idx, row_0, dim.height)
            if dim.hidden:
                wb.set_row_hidden(idx, row_0)
            if dim.outline_level and dim.outline_level > 0:
                wb.set_row_outline_level(idx, row_0, dim.outline_level)

        # Freeze panes
        if self.freeze_panes:
            r, c = _parse_cell_ref(self.freeze_panes)
            wb.set_freeze_panes(idx, r - 1, c - 1)

        # Sheet visibility
        if self._sheet_state != "visible":
            state_map = {"hidden": 1, "veryHidden": 2}
            wb.set_sheet_visibility(idx, state_map[self._sheet_state])

        # Zoom
        if self._zoom_scale is not None:
            wb.set_zoom(idx, int(self._zoom_scale))

        # Gridlines
        if not self._show_gridlines:
            wb.set_show_gridlines(idx, self._show_gridlines)

        # Auto-fit columns
        if self._autofit:
            wb.set_autofit(idx, True)

        # Page breaks
        if self.row_breaks:
            breaks = [brk.id for brk in self.row_breaks]
            wb.set_row_breaks(idx, breaks)
        if self.col_breaks:
            breaks = [brk.id for brk in self.col_breaks]
            wb.set_col_breaks(idx, breaks)

        # Autofilter
        if self.auto_filter._ref:
            parts = self.auto_filter._ref.split(":")
            r1, c1 = _parse_cell_ref(parts[0])
            r2, c2 = _parse_cell_ref(parts[1])
            wb.set_autofilter(idx, r1 - 1, c1 - 1, r2 - 1, c2 - 1)

        # Autofilter column filters
        for fc in self.auto_filter._filter_columns:
            wb.add_autofilter_column(idx, json.dumps(fc))

        # Protection
        if self.protection.sheet:
            prot_data = {
                "password": self.protection._password,
                "format_cells": self.protection.format_cells,
                "format_rows": self.protection.format_rows,
                "format_columns": self.protection.format_columns,
                "insert_columns": self.protection.insert_columns,
                "insert_rows": self.protection.insert_rows,
                "insert_hyperlinks": self.protection.insert_hyperlinks,
                "delete_columns": self.protection.delete_columns,
                "delete_rows": self.protection.delete_rows,
                "select_locked_cells": self.protection.select_locked_cells,
                "select_unlocked_cells": self.protection.select_unlocked_cells,
                "sort": self.protection.sort,
                "autofilter": self.protection.autofilter,
                "pivot_tables": self.protection.pivot_tables,
                "objects": self.protection.objects,
                "scenarios": self.protection.scenarios,
            }
            wb.set_protection(idx, json.dumps(prot_data))

        # Page setup
        page_data = {}
        if self.page_setup.orientation:
            page_data["orientation"] = self.page_setup.orientation
        if self.page_setup.paperSize is not None:
            page_data["paper_size"] = self.page_setup.paperSize
        if self.page_setup.scale is not None:
            page_data["scale"] = self.page_setup.scale
        if self.page_setup.fitToWidth is not None or self.page_setup.fitToHeight is not None:
            page_data["fit_to_width"] = self.page_setup.fitToWidth or 0
            page_data["fit_to_height"] = self.page_setup.fitToHeight or 0
        page_data["margins"] = {
            "left": self.page_margins.left,
            "right": self.page_margins.right,
            "top": self.page_margins.top,
            "bottom": self.page_margins.bottom,
            "header": self.page_margins.header,
            "footer": self.page_margins.footer,
        }
        if self.print_area:
            page_data["print_area"] = self.print_area
        if self.print_title_rows:
            page_data["print_title_rows"] = self.print_title_rows
        if self.print_title_cols:
            page_data["print_title_cols"] = self.print_title_cols
        if self.print_options.horizontalCentered:
            page_data["center_horizontally"] = True
        if self.print_options.verticalCentered:
            page_data["center_vertically"] = True
        if self.print_options.gridLines:
            page_data["gridlines"] = True
        if self.print_options.headings:
            page_data["headings"] = True
        # Headers and footers
        header_str = self.oddHeader._build_format_string()
        if header_str:
            page_data["header_text"] = header_str
        footer_str = self.oddFooter._build_format_string()
        if footer_str:
            page_data["footer_text"] = footer_str
        if page_data:
            wb.set_page_setup(idx, json.dumps(page_data))

        # Data Validations
        for dv in self._data_validations:
            ranges = []
            if dv.sqref:
                for ref_str in dv.sqref.split():
                    if ":" in ref_str:
                        parts = ref_str.split(":")
                        r1, c1 = _parse_cell_ref(parts[0])
                        r2, c2 = _parse_cell_ref(parts[1])
                        ranges.append([r1 - 1, c1 - 1, r2 - 1, c2 - 1])
                    else:
                        r, c = _parse_cell_ref(ref_str)
                        ranges.append([r - 1, c - 1, r - 1, c - 1])
            for cell_ref in dv._cells:
                if ":" in cell_ref:
                    parts = cell_ref.split(":")
                    r1, c1 = _parse_cell_ref(parts[0])
                    r2, c2 = _parse_cell_ref(parts[1])
                    ranges.append([r1 - 1, c1 - 1, r2 - 1, c2 - 1])
                else:
                    r, c = _parse_cell_ref(cell_ref)
                    ranges.append([r - 1, c - 1, r - 1, c - 1])

            dv_data = {
                "type": dv.type,
                "formula1": str(dv.formula1) if dv.formula1 is not None else "",
                "formula2": str(dv.formula2) if dv.formula2 is not None else None,
                "operator": dv.operator,
                "allow_blank": dv.allow_blank,
                "show_dropdown": dv.showDropDown,
                "show_input_message": dv.showInputMessage,
                "show_error_message": dv.showErrorMessage,
                "input_title": dv.promptTitle,
                "input_message": dv.prompt,
                "error_title": dv.errorTitle,
                "error_message": dv.error,
                "error_style": dv.errorStyle,
                "ranges": ranges,
            }
            wb.add_data_validation(idx, json.dumps(dv_data))

        # Conditional Formatting
        for range_string, rule in self.conditional_formatting._rules:
            cf_json = self._serialize_conditional_format(range_string, rule)
            if cf_json:
                wb.add_conditional_format(idx, cf_json)

        # Images
        for img in self._images:
            if img.anchor:
                r, c = _parse_cell_ref(img.anchor)
                wb.add_image(idx, r - 1, c - 1, img._data, None, None)

        # Tables
        for table in self._tables:
            table_data = {
                "ref": table.ref,
                "name": table.displayName,
                "header_row": table.headerRowCount > 0,
                "total_row": table.totalsRowCount > 0,
            }
            if table.tableStyleInfo is not None:
                si = table.tableStyleInfo
                table_data["style"] = si.name
                table_data["first_column"] = si.showFirstColumn
                table_data["last_column"] = si.showLastColumn
                table_data["row_stripes"] = si.showRowStripes
                table_data["column_stripes"] = si.showColumnStripes
            if table.tableColumns:
                table_data["columns"] = [{"name": tc.name} for tc in table.tableColumns]
            wb.add_table(idx, json.dumps(table_data))

        # Charts
        for chart in self._charts:
            chart_data = self._serialize_chart(chart)
            if chart_data:
                wb.add_chart(idx, json.dumps(chart_data))

    def _serialize_chart(self, chart):
        """Serialize a chart object to a JSON-compatible dict."""
        from openpyxl_rust.chart.base import _CellTitle

        if not chart._anchor:
            return None

        r, c = _parse_cell_ref(chart._anchor)

        # Determine rust_xlsxwriter chart type
        chart_type = chart.chart_type
        if chart_type == "column" and hasattr(chart, "type") and chart.type == "bar":
            chart_type = "bar"

        # Handle grouping → stacked/percentStacked variants
        grouping = getattr(chart, "grouping", None)
        if grouping == "stacked" and chart_type in ("column", "bar", "line", "area"):
            chart_type = chart_type + "_stacked"
        elif grouping == "percentStacked" and chart_type in ("column", "bar", "line", "area"):
            chart_type = chart_type + "_percent_stacked"

        # Convert dimensions from cm to pixels (approx 37.8 px/cm)
        width_px = int(chart.width * 37.8) if chart.width else 480
        height_px = int(chart.height * 37.8) if chart.height else 288

        series_list = []
        for s in chart.series:
            s_data = {}
            if s.values is not None:
                ref = s.values
                sheet_title = ref.worksheet.title if ref.worksheet else "Sheet1"
                s_data["values"] = {
                    "sheet": sheet_title,
                    "r1": ref.min_row - 1,
                    "c1": ref.min_col - 1,
                    "r2": ref.max_row - 1,
                    "c2": ref.max_col - 1,
                }
            if s.categories is not None:
                ref = s.categories
                sheet_title = ref.worksheet.title if ref.worksheet else "Sheet1"
                s_data["categories"] = {
                    "sheet": sheet_title,
                    "r1": ref.min_row - 1,
                    "c1": ref.min_col - 1,
                    "r2": ref.max_row - 1,
                    "c2": ref.max_col - 1,
                }
            if s.title is not None:
                if isinstance(s.title, _CellTitle):
                    s_data["title"] = s.title.resolve()
                else:
                    s_data["title"] = str(s.title)

            # Trendline
            if s.trendline is not None:
                tl = s.trendline
                s_data["trendline"] = {
                    "type": tl.trendlineType,
                    "display_equation": getattr(tl, "displayEquation", False),
                    "display_r_squared": getattr(tl, "displayRSqr", False),
                }

            # Data labels
            if s.dLbls is not None:
                dl = s.dLbls
                s_data["data_labels"] = {
                    "show_value": getattr(dl, "showVal", False),
                    "show_category": getattr(dl, "showCatName", False),
                    "show_series": getattr(dl, "showSerName", False),
                }

            series_list.append(s_data)

        # Legend handling: support both bool and ChartLegend object
        if isinstance(chart.legend, bool):
            legend_val = chart.legend
        elif chart.legend._hidden:
            legend_val = False
        else:
            legend_val = True

        data = {
            "type": chart_type,
            "anchor_row": r - 1,
            "anchor_col": c - 1,
            "width": width_px,
            "height": height_px,
            "legend": legend_val,
            "series": series_list,
        }

        # Legend position
        if not isinstance(chart.legend, bool) and chart.legend.position:
            data["legend_position"] = chart.legend.position
        if chart.title is not None:
            data["title"] = str(chart.title)
        if chart.x_axis_title is not None:
            data["x_axis_title"] = str(chart.x_axis_title)
        if chart.y_axis_title is not None:
            data["y_axis_title"] = str(chart.y_axis_title)
        if chart.style is not None:
            data["style"] = chart.style
        return data

    # ---- Conditional formatting serialization (same as before) ----

    def _serialize_rule_format(self, rule):
        fmt = {}
        if rule.font is not None:
            font_data = {
                "bold": rule.font.bold,
                "italic": rule.font.italic,
                "name": rule.font.name,
                "size": rule.font.size,
            }
            if rule.font.underline is not None:
                font_data["underline"] = rule.font.underline
            if rule.font.color is not None:
                font_data["color"] = rule.font.color
            if rule.font.strikethrough:
                font_data["strikethrough"] = True
            if rule.font.vertAlign is not None:
                font_data["vertAlign"] = rule.font.vertAlign
            fmt["font"] = font_data
        if rule.fill is not None:
            fill_data = {}
            if rule.fill.fill_type is not None:
                fill_data["fill_type"] = rule.fill.fill_type
            if rule.fill.start_color is not None:
                fill_data["start_color"] = rule.fill.start_color
            if rule.fill.end_color is not None:
                fill_data["end_color"] = rule.fill.end_color
            if fill_data:
                fmt["fill"] = fill_data
        if rule.border is not None:
            border_data = {}
            for side_name in ("left", "right", "top", "bottom"):
                side = getattr(rule.border, side_name)
                if side.style is not None:
                    side_data = {"style": side.style}
                    if side.color is not None:
                        side_data["color"] = side.color
                    border_data[side_name] = side_data
            if rule.border.diagonal.style is not None:
                diag_data = {"style": rule.border.diagonal.style}
                if rule.border.diagonal.color is not None:
                    diag_data["color"] = rule.border.diagonal.color
                diag_data["diagonalUp"] = rule.border.diagonalUp
                diag_data["diagonalDown"] = rule.border.diagonalDown
                border_data["diagonal"] = diag_data
            if border_data:
                fmt["border"] = border_data
        return fmt

    def _serialize_conditional_format(self, range_string, rule):
        if isinstance(rule, ColorScaleRule):
            is_3_color = rule.mid_type is not None or rule.mid_color is not None
            data = {
                "rule_type": "3_color_scale" if is_3_color else "2_color_scale",
                "range": range_string,
                "start_type": rule.start_type,
                "start_value": rule.start_value,
                "start_color": rule.start_color,
                "end_type": rule.end_type,
                "end_value": rule.end_value,
                "end_color": rule.end_color,
            }
            if is_3_color:
                data["mid_type"] = rule.mid_type
                data["mid_value"] = rule.mid_value
                data["mid_color"] = rule.mid_color
            return json.dumps(data)

        elif isinstance(rule, DataBarRule):
            data = {
                "rule_type": "data_bar",
                "range": range_string,
                "color": rule.color,
                "bar_only": rule.showValue is False,
            }
            return json.dumps(data)

        elif isinstance(rule, IconSetRule):
            data = {
                "rule_type": "icon_set",
                "range": range_string,
                "icon_style": rule.icon_style or "3TrafficLights1",
                "reverse": bool(rule.reverse),
                "show_icons_only": rule.showValue is False,
            }
            return json.dumps(data)

        elif isinstance(rule, CellIsRule):
            data = {
                "rule_type": "cell_is",
                "range": range_string,
                "operator": rule.operator,
                "formula": list(rule.formula),
                "stop_if_true": bool(rule.stopIfTrue),
            }
            fmt = self._serialize_rule_format(rule)
            if fmt:
                data["format"] = fmt
            return json.dumps(data)

        elif isinstance(rule, FormulaRule):
            formula_list = list(rule.formula)
            data = {
                "rule_type": "formula",
                "range": range_string,
                "formula": formula_list[0] if formula_list else "",
                "stop_if_true": bool(rule.stopIfTrue),
            }
            fmt = self._serialize_rule_format(rule)
            if fmt:
                data["format"] = fmt
            return json.dumps(data)

        elif isinstance(rule, Top10Rule):
            data = {
                "rule_type": "top10",
                "range": range_string,
                "rank": rule.rank,
                "percent": rule.percent,
                "bottom": rule.bottom,
            }
            fmt = self._serialize_rule_format(rule)
            if fmt:
                data["format"] = fmt
            return json.dumps(data)

        elif isinstance(rule, DuplicateRule):
            data = {
                "rule_type": "duplicate",
                "range": range_string,
            }
            fmt = self._serialize_rule_format(rule)
            if fmt:
                data["format"] = fmt
            return json.dumps(data)

        elif isinstance(rule, TextRule):
            data = {
                "rule_type": "text",
                "range": range_string,
                "operator": rule.operator,
                "text": rule.text,
            }
            fmt = self._serialize_rule_format(rule)
            if fmt:
                data["format"] = fmt
            return json.dumps(data)

        return None
