import json
import re
from datetime import datetime, date, time
from openpyxl_rust.cell import (
    Cell, _col_letter, _date_to_excel_serial,
    _BORDER_STYLE_MAP, _FILL_TYPE_MAP, _HALIGN_MAP, _VALIGN_MAP,
    _underline_to_u8, _vert_align_to_u8,
)
from openpyxl_rust.protection import SheetProtection
from openpyxl_rust.page import PrintPageSetup, PageMargins, PrintOptions
from openpyxl_rust.formatting.rule import (
    ColorScaleRule, DataBarRule, IconSetRule, CellIsRule, FormulaRule
)


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


class RowDimension:
    def __init__(self):
        self.height = None


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

    @property
    def ref(self):
        return self._ref

    @ref.setter
    def ref(self, value):
        self._ref = value


class _ConditionalFormattingList:
    def __init__(self):
        self._rules = []

    def add(self, range_string, rule):
        self._rules.append((range_string, rule))


class Worksheet:
    # openpyxl compat constants
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
        self._cell_cache = {}  # (row, col) -> Cell or raw value
        self._current_row = 0  # 0 means no rows appended yet; next append goes to row 1
        self._max_row = 0  # tracks highest row written, for O(1) _next_row()
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

        # If connected to a RustWorkbook, sync the title
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
                    raise ValueError(
                        f"Worksheet title '{value}' already exists. "
                        "Use a unique title."
                    )
        self._title = value
        if self._workbook is not None and self._sheet_idx is not None:
            self._workbook._rust_wb.set_sheet_title(self._sheet_idx, value)

    @property
    def min_row(self):
        if not self._cell_cache:
            return None
        return min(r for r, _ in self._cell_cache)

    @property
    def max_row(self):
        if not self._cell_cache:
            return None
        return max(r for r, _ in self._cell_cache)

    @property
    def min_column(self):
        if not self._cell_cache:
            return None
        return min(c for _, c in self._cell_cache)

    @property
    def max_column(self):
        if not self._cell_cache:
            return None
        return max(c for _, c in self._cell_cache)

    @property
    def dimensions(self):
        if not self._cell_cache:
            return ""
        return f"{_col_letter(self.min_column)}{self.min_row}:{_col_letter(self.max_column)}{self.max_row}"

    def iter_rows(self, min_row=None, max_row=None, min_col=None, max_col=None, values_only=False):
        min_row = min_row if min_row is not None else self.min_row
        max_row = max_row if max_row is not None else self.max_row
        min_col = min_col if min_col is not None else self.min_column
        max_col = max_col if max_col is not None else self.max_column
        if any(v is None for v in (min_row, max_row, min_col, max_col)):
            return
        for row in range(min_row, max_row + 1):
            if values_only:
                yield tuple(
                    self._cell_value((row, col))
                    for col in range(min_col, max_col + 1)
                )
            else:
                yield tuple(
                    self.cell(row=row, column=col)
                    for col in range(min_col, max_col + 1)
                )

    def iter_cols(self, min_col=None, max_col=None, min_row=None, max_row=None, values_only=False):
        min_row = min_row if min_row is not None else self.min_row
        max_row = max_row if max_row is not None else self.max_row
        min_col = min_col if min_col is not None else self.min_column
        max_col = max_col if max_col is not None else self.max_column
        if any(v is None for v in (min_row, max_row, min_col, max_col)):
            return
        for col in range(min_col, max_col + 1):
            if values_only:
                yield tuple(
                    self._cell_value((row, col))
                    for row in range(min_row, max_row + 1)
                )
            else:
                yield tuple(
                    self.cell(row=row, column=col)
                    for row in range(min_row, max_row + 1)
                )

    @property
    def values(self):
        return self.iter_rows(values_only=True)

    def __setitem__(self, key, value):
        row, col = _parse_cell_ref(key)
        self.cell(row=row, column=col, value=value)

    def __getitem__(self, key):
        if isinstance(key, int):
            min_col = self.min_column or 1
            max_col = self.max_column or 1
            return tuple(
                self.cell(row=key, column=col)
                for col in range(min_col, max_col + 1)
            )
        if ':' in key:
            start_ref, end_ref = key.split(':')
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            rows = []
            for row in range(r1, r2 + 1):
                rows.append(tuple(
                    self.cell(row=row, column=col)
                    for col in range(c1, c2 + 1)
                ))
            return tuple(rows)
        row, col = _parse_cell_ref(key)
        return self.cell(row=row, column=col)

    def _cell_value(self, key):
        """Return the Python value for a cache entry (Cell or raw value)."""
        entry = self._cell_cache.get(key)
        if entry is None:
            return None
        if isinstance(entry, Cell):
            return entry.value
        return entry

    def cell(self, row, column, value=None):
        key = (row, column)
        existing = self._cell_cache.get(key)
        if existing is None:
            c = Cell(row, column)
            self._cell_cache[key] = c
        elif isinstance(existing, Cell):
            c = existing
        else:
            # Upgrade raw value to Cell object
            c = Cell(row, column, existing)
            self._cell_cache[key] = c
        if row > self._max_row:
            self._max_row = row
        if value is not None:
            c.value = value
            if self._workbook is not None:
                t = type(value)
                r0 = row - 1
                c0 = column - 1
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
                    wb.set_cell_datetime(idx, r0, c0, serial)
                    if c.number_format == "General":
                        c.number_format = "yyyy-mm-dd hh:mm:ss"
                elif t is date:
                    serial = _date_to_excel_serial(value.year, value.month, value.day)
                    wb.set_cell_datetime(idx, r0, c0, serial)
                    if c.number_format == "General":
                        c.number_format = "yyyy-mm-dd"
                elif t is time:
                    serial = (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86400.0
                    wb.set_cell_datetime(idx, r0, c0, serial)
                    if c.number_format == "General":
                        c.number_format = "hh:mm:ss"
                else:
                    raise TypeError(
                        f"Unsupported cell value type: {type(value).__name__}. "
                        f"Supported types: str, int, float, bool, datetime, date, time, None"
                    )
        return c

    def _set_rust_value(self, row, col, value):
        """Push a single cell value to Rust. Used by _resync_rust."""
        if self._workbook is None:
            return
        wb = self._workbook._rust_wb
        idx = self._sheet_idx
        t = type(value)
        if t is str:
            wb.set_cell_string(idx, row, col, value)
        elif t is float:
            wb.set_cell_number(idx, row, col, value)
        elif t is int:
            wb.set_cell_number(idx, row, col, float(value))
        elif t is bool:
            wb.set_cell_boolean(idx, row, col, value)
        elif t is datetime:
            serial = _date_to_excel_serial(value.year, value.month, value.day)
            serial += (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86400.0
            wb.set_cell_datetime(idx, row, col, serial)
        elif t is date:
            serial = _date_to_excel_serial(value.year, value.month, value.day)
            wb.set_cell_datetime(idx, row, col, serial)
        elif t is time:
            serial = (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86400.0
            wb.set_cell_datetime(idx, row, col, serial)

    def _next_row(self):
        return max(self._current_row, self._max_row) + 1

    def append(self, iterable):
        row = self._next_row()
        values = list(iterable)

        if self._workbook is not None and self._sheet_idx is not None:
            # Build converted row for single batch FFI call
            converted = []
            datetime_formats = []
            for col_idx, value in enumerate(values):
                t = type(value)
                if value is None:
                    converted.append(None)
                elif t is datetime:
                    serial = _date_to_excel_serial(value.year, value.month, value.day)
                    serial += (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86400.0
                    converted.append(serial)
                    datetime_formats.append((col_idx, "yyyy-mm-dd hh:mm:ss"))
                elif t is date:
                    serial = _date_to_excel_serial(value.year, value.month, value.day)
                    converted.append(serial)
                    datetime_formats.append((col_idx, "yyyy-mm-dd"))
                elif t is time:
                    serial = (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86400.0
                    converted.append(serial)
                    datetime_formats.append((col_idx, "hh:mm:ss"))
                elif t is int:
                    converted.append(float(value))
                else:
                    converted.append(value)

            r0 = row - 1
            self._workbook._rust_wb.set_rows_batch(
                self._sheet_idx, r0, [converted]
            )
            wb = self._workbook._rust_wb
            for col_idx, fmt_str in datetime_formats:
                wb.set_cell_number_format(self._sheet_idx, r0, col_idx, fmt_str)

        # Store Cell objects (append is interactive path, user may access cells)
        for col_idx, value in enumerate(values, start=1):
            c = Cell(row, col_idx, value)
            # Set datetime number_format
            if isinstance(value, datetime):
                c.number_format = "yyyy-mm-dd hh:mm:ss"
            elif isinstance(value, date):
                c.number_format = "yyyy-mm-dd"
            elif isinstance(value, time):
                c.number_format = "hh:mm:ss"
            self._cell_cache[(row, col_idx)] = c

        self._current_row = row
        if row > self._max_row:
            self._max_row = row

    def append_rows(self, rows_data):
        """Append multiple rows at once via the Rust batch API for maximum speed.

        Stores raw values in _cell_cache (not Cell objects) to avoid creating
        millions of Cell instances that are never accessed.
        """
        # Materialize generator to list for safe double-iteration
        rows_data = list(rows_data)
        start_row = self._next_row()
        start_row_0based = start_row - 1

        datetime_cells = []

        if self._workbook is not None and self._sheet_idx is not None:
            rows_list = []
            for row_offset, row in enumerate(rows_data):
                converted_row = []
                for col_offset, value in enumerate(row):
                    t = type(value)
                    if t is datetime:
                        serial = _date_to_excel_serial(value.year, value.month, value.day)
                        serial += (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86400.0
                        converted_row.append(serial)
                        datetime_cells.append((start_row_0based + row_offset, col_offset, "yyyy-mm-dd hh:mm:ss"))
                    elif t is date:
                        serial = _date_to_excel_serial(value.year, value.month, value.day)
                        converted_row.append(serial)
                        datetime_cells.append((start_row_0based + row_offset, col_offset, "yyyy-mm-dd"))
                    elif t is time:
                        serial = (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86400.0
                        converted_row.append(serial)
                        datetime_cells.append((start_row_0based + row_offset, col_offset, "hh:mm:ss"))
                    else:
                        converted_row.append(value)
                rows_list.append(converted_row)
            self._workbook._rust_wb.set_rows_batch(
                self._sheet_idx, start_row_0based, rows_list
            )

            wb = self._workbook._rust_wb
            for r0, c0, fmt_str in datetime_cells:
                wb.set_cell_number_format(self._sheet_idx, r0, c0, fmt_str)

        # Store raw values (not Cell objects) — avoids 1M Cell creations
        for row_offset, row_values in enumerate(rows_data):
            row = start_row + row_offset
            for col_idx, value in enumerate(row_values, start=1):
                if value is not None:
                    self._cell_cache[(row, col_idx)] = value
            self._current_row = row
        if self._current_row > self._max_row:
            self._max_row = self._current_row

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
            self._workbook._rust_wb.add_merge_range(
                self._sheet_idx, r1 - 1, c1 - 1, r2 - 1, c2 - 1
            )

    def unmerge_cells(self, range_string):
        parts = range_string.split(":")
        if len(parts) != 2:
            raise ValueError(f"Invalid merge range: {range_string}")
        start_ref = parts[0].upper()
        end_ref = parts[1].upper()
        target = (start_ref, end_ref)
        if target not in self.merged_cell_ranges:
            raise ValueError(
                f"Cell range {range_string} is not merged"
            )
        self.merged_cell_ranges.remove(target)
        self._resync_rust()

    def _resync_rust(self):
        """Clear and re-push all cell data to Rust after structural changes."""
        if self._workbook is None or self._sheet_idx is None:
            return
        wb = self._workbook._rust_wb
        idx = self._sheet_idx
        wb.clear_cells(idx)
        wb.clear_merge_ranges(idx)

        for (row, col), entry in self._cell_cache.items():
            value = entry.value if isinstance(entry, Cell) else entry
            if value is not None:
                self._set_rust_value(row - 1, col - 1, value)

        # Re-push formats
        self._flush_formats_to_rust()

        for (start_ref, end_ref) in self.merged_cell_ranges:
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            wb.add_merge_range(idx, r1 - 1, c1 - 1, r2 - 1, c2 - 1)

    def insert_rows(self, idx, amount=1):
        new_cells = {}
        for (row, col), entry in self._cell_cache.items():
            if row >= idx:
                new_row = row + amount
                if isinstance(entry, Cell):
                    entry.row = new_row
                new_cells[(new_row, col)] = entry
            else:
                new_cells[(row, col)] = entry
        self._cell_cache = new_cells
        new_merged = []
        for (start_ref, end_ref) in self.merged_cell_ranges:
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            if r1 >= idx:
                r1 += amount
            if r2 >= idx:
                r2 += amount
            new_merged.append((_col_letter(c1) + str(r1), _col_letter(c2) + str(r2)))
        self.merged_cell_ranges = new_merged
        if self._current_row >= idx:
            self._current_row += amount
        self._max_row = max((r for r, _ in self._cell_cache), default=0)
        self._resync_rust()

    def delete_rows(self, idx, amount=1):
        new_cells = {}
        for (row, col), entry in self._cell_cache.items():
            if idx <= row < idx + amount:
                continue
            elif row >= idx + amount:
                new_row = row - amount
                if isinstance(entry, Cell):
                    entry.row = new_row
                new_cells[(new_row, col)] = entry
            else:
                new_cells[(row, col)] = entry
        self._cell_cache = new_cells
        new_merged = []
        for (start_ref, end_ref) in self.merged_cell_ranges:
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
        if self._current_row >= idx + amount:
            self._current_row -= amount
        elif self._current_row >= idx:
            self._current_row = max(idx - 1, 0)
        self._max_row = max((r for r, _ in self._cell_cache), default=0)
        self._resync_rust()

    def insert_cols(self, idx, amount=1):
        new_cells = {}
        for (row, col), entry in self._cell_cache.items():
            if col >= idx:
                new_col = col + amount
                if isinstance(entry, Cell):
                    entry.column = new_col
                new_cells[(row, new_col)] = entry
            else:
                new_cells[(row, col)] = entry
        self._cell_cache = new_cells
        new_merged = []
        for (start_ref, end_ref) in self.merged_cell_ranges:
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            if c1 >= idx:
                c1 += amount
            if c2 >= idx:
                c2 += amount
            new_merged.append((_col_letter(c1) + str(r1), _col_letter(c2) + str(r2)))
        self.merged_cell_ranges = new_merged
        self._resync_rust()

    def delete_cols(self, idx, amount=1):
        new_cells = {}
        for (row, col), entry in self._cell_cache.items():
            if idx <= col < idx + amount:
                continue
            elif col >= idx + amount:
                new_col = col - amount
                if isinstance(entry, Cell):
                    entry.column = new_col
                new_cells[(row, new_col)] = entry
            else:
                new_cells[(row, col)] = entry
        self._cell_cache = new_cells
        new_merged = []
        for (start_ref, end_ref) in self.merged_cell_ranges:
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
        self._resync_rust()

    def add_image(self, img, anchor=None):
        if anchor is not None:
            img.anchor = anchor
        self._images.append(img)

    def add_data_validation(self, dv):
        self._data_validations.append(dv)

    def _flush_formats_to_rust(self):
        """Push cell formats to Rust using CellFormat struct (no JSON).

        Called at save time. Only processes cells that have non-default formats.
        """
        if self._workbook is None or self._sheet_idx is None:
            return
        wb = self._workbook._rust_wb
        idx = self._sheet_idx

        for (row, col), cell in self._cell_cache.items():
            if not isinstance(cell, Cell):
                continue  # raw value, no format to push
            r0, c0 = row - 1, col - 1

            # Number format
            if cell.number_format and cell.number_format != "General":
                wb.set_cell_number_format(idx, r0, c0, cell.number_format)

            # Font
            if cell.font is not None:
                f = cell.font
                wb.set_cell_font(
                    idx, r0, c0,
                    f.bold, f.italic, f.name,
                    float(f.size) if f.size is not None else None,
                    f.color,
                    _underline_to_u8(f.underline),
                    f.strikethrough,
                    _vert_align_to_u8(f.vertAlign),
                )

            # Alignment
            if cell.alignment is not None:
                a = cell.alignment
                wb.set_cell_alignment(
                    idx, r0, c0,
                    _HALIGN_MAP.get(a.horizontal) if a.horizontal else None,
                    _VALIGN_MAP.get(a.vertical) if a.vertical else None,
                    bool(a.wrap_text),
                    bool(a.shrink_to_fit),
                    int(a.indent) if a.indent else 0,
                    int(a.text_rotation) if a.text_rotation else 0,
                )

            # Fill
            if cell.fill is not None:
                fi = cell.fill
                wb.set_cell_fill(
                    idx, r0, c0,
                    _FILL_TYPE_MAP.get(fi.fill_type) if fi.fill_type else None,
                    fi.start_color,
                    fi.end_color,
                )

            # Border
            if cell.border is not None:
                b = cell.border
                wb.set_cell_border(
                    idx, r0, c0,
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

    def _flush_metadata(self):
        """Push column widths, row heights, freeze panes, hyperlinks, comments,
        and other sheet-level metadata to Rust. Called right before save.
        """
        if self._workbook is None or self._sheet_idx is None:
            return
        wb = self._workbook._rust_wb
        idx = self._sheet_idx

        # Single pass: formats + hyperlinks + comments (skip raw values)
        for (row, col), cell in self._cell_cache.items():
            if not isinstance(cell, Cell):
                continue
            r0, c0 = row - 1, col - 1

            # Number format
            if cell.number_format and cell.number_format != "General":
                wb.set_cell_number_format(idx, r0, c0, cell.number_format)

            # Font
            if cell.font is not None:
                f = cell.font
                wb.set_cell_font(
                    idx, r0, c0,
                    f.bold, f.italic, f.name,
                    float(f.size) if f.size is not None else None,
                    f.color,
                    _underline_to_u8(f.underline),
                    f.strikethrough,
                    _vert_align_to_u8(f.vertAlign),
                )

            # Alignment
            if cell.alignment is not None:
                a = cell.alignment
                wb.set_cell_alignment(
                    idx, r0, c0,
                    _HALIGN_MAP.get(a.horizontal) if a.horizontal else None,
                    _VALIGN_MAP.get(a.vertical) if a.vertical else None,
                    bool(a.wrap_text),
                    bool(a.shrink_to_fit),
                    int(a.indent) if a.indent else 0,
                    int(a.text_rotation) if a.text_rotation else 0,
                )

            # Fill
            if cell.fill is not None:
                fi = cell.fill
                wb.set_cell_fill(
                    idx, r0, c0,
                    _FILL_TYPE_MAP.get(fi.fill_type) if fi.fill_type else None,
                    fi.start_color,
                    fi.end_color,
                )

            # Border
            if cell.border is not None:
                b = cell.border
                wb.set_cell_border(
                    idx, r0, c0,
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

            # Hyperlink
            if cell.hyperlink is not None:
                url = cell.hyperlink
                text = None
                tooltip = None
                if isinstance(url, str) and url.startswith("#"):
                    url = "internal:" + url[1:]
                wb.add_hyperlink(idx, r0, c0, url, text, tooltip)

            # Comment/Note
            if cell.comment is not None:
                author = cell.comment.author if cell.comment.author else None
                wb.add_note(idx, r0, c0, cell.comment.text, author)

        # Column widths
        for letter, dim in self.column_dimensions.items():
            if dim.width is not None:
                _, col_idx = _parse_cell_ref(f"{letter}1")
                wb.set_column_width(idx, col_idx - 1, dim.width)

        # Row heights
        for row_num, dim in self.row_dimensions.items():
            if dim.height is not None:
                wb.set_row_height(idx, row_num - 1, dim.height)

        # Freeze panes
        if self.freeze_panes:
            r, c = _parse_cell_ref(self.freeze_panes)
            wb.set_freeze_panes(idx, r - 1, c - 1)

        # Autofilter
        if self.auto_filter._ref:
            parts = self.auto_filter._ref.split(":")
            r1, c1 = _parse_cell_ref(parts[0])
            r2, c2 = _parse_cell_ref(parts[1])
            wb.set_autofilter(idx, r1 - 1, c1 - 1, r2 - 1, c2 - 1)

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

        return None
