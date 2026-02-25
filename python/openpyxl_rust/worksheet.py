import json
import re
from datetime import datetime, date, time
from openpyxl_rust.cell import Cell, _col_letter
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


def _date_to_excel_serial(year, month, day):
    """Convert a date to Excel serial number (matches Rust implementation)."""
    y = year
    m = month
    if m <= 2:
        y -= 1
        m += 12
    a = y // 100
    b = 2 - a + a // 4
    jd = int(365.25 * (y + 4716)) + int(30.6001 * (m + 1)) + day + b - 1524
    excel_epoch_jd = 2415020
    serial = float(jd - excel_epoch_jd)
    if serial > 59.0:
        serial += 1.0
    return serial


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
        self._cells = {}  # (row, col) -> Cell
        self._current_row = 0  # 0 means no rows appended yet; next append goes to row 1
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
            # Check for duplicate title among other sheets
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
        """Smallest row index containing data (1-based), or None if empty."""
        if not self._cells:
            return None
        return min(r for r, _ in self._cells)

    @property
    def max_row(self):
        """Largest row index containing data (1-based), or None if empty."""
        if not self._cells:
            return None
        return max(r for r, _ in self._cells)

    @property
    def min_column(self):
        """Smallest column index containing data (1-based), or None if empty."""
        if not self._cells:
            return None
        return min(c for _, c in self._cells)

    @property
    def max_column(self):
        """Largest column index containing data (1-based), or None if empty."""
        if not self._cells:
            return None
        return max(c for _, c in self._cells)

    @property
    def dimensions(self):
        """Return string like 'A1:D10' representing the data extent, or '' if empty."""
        if not self._cells:
            return ""
        return f"{_col_letter(self.min_column)}{self.min_row}:{_col_letter(self.max_column)}{self.max_row}"

    def iter_rows(self, min_row=None, max_row=None, min_col=None, max_col=None, values_only=False):
        """Yield tuples of Cell objects (or values if values_only=True) per row."""
        min_row = min_row if min_row is not None else self.min_row
        max_row = max_row if max_row is not None else self.max_row
        min_col = min_col if min_col is not None else self.min_column
        max_col = max_col if max_col is not None else self.max_column
        if any(v is None for v in (min_row, max_row, min_col, max_col)):
            return
        for row in range(min_row, max_row + 1):
            if values_only:
                yield tuple(
                    self._cells[(row, col)].value if (row, col) in self._cells else None
                    for col in range(min_col, max_col + 1)
                )
            else:
                yield tuple(
                    self._cells.get((row, col)) or self.cell(row=row, column=col)
                    for col in range(min_col, max_col + 1)
                )

    def iter_cols(self, min_col=None, max_col=None, min_row=None, max_row=None, values_only=False):
        """Yield tuples of Cell objects (or values if values_only=True) per column."""
        min_row = min_row if min_row is not None else self.min_row
        max_row = max_row if max_row is not None else self.max_row
        min_col = min_col if min_col is not None else self.min_column
        max_col = max_col if max_col is not None else self.max_column
        if any(v is None for v in (min_row, max_row, min_col, max_col)):
            return
        for col in range(min_col, max_col + 1):
            if values_only:
                yield tuple(
                    self._cells[(row, col)].value if (row, col) in self._cells else None
                    for row in range(min_row, max_row + 1)
                )
            else:
                yield tuple(
                    self._cells.get((row, col)) or self.cell(row=row, column=col)
                    for row in range(min_row, max_row + 1)
                )

    @property
    def values(self):
        """Yield tuples of cell values per row."""
        return self.iter_rows(values_only=True)

    def _set_rust_value(self, row, col, value):
        """Push a cell value to the Rust side. row/col are 1-based."""
        if self._workbook is None or self._sheet_idx is None:
            return
        r = row - 1  # Rust uses 0-based
        c = col - 1
        wb = self._workbook._rust_wb
        # datetime must come before date because datetime is subclass of date
        if isinstance(value, datetime):
            serial = _date_to_excel_serial(value.year, value.month, value.day)
            serial += (value.hour * 3600 + value.minute * 60 + value.second) / 86400.0
            wb.set_cell_datetime(self._sheet_idx, r, c, serial, False)
        elif isinstance(value, date):
            serial = _date_to_excel_serial(value.year, value.month, value.day)
            wb.set_cell_datetime(self._sheet_idx, r, c, serial, True)
        elif isinstance(value, time):
            serial = (value.hour * 3600 + value.minute * 60 + value.second) / 86400.0
            wb.set_cell_datetime(self._sheet_idx, r, c, serial, False)
        elif isinstance(value, bool):
            wb.set_cell_boolean(self._sheet_idx, r, c, value)
        elif isinstance(value, (int, float)):
            wb.set_cell_number(self._sheet_idx, r, c, float(value))
        elif isinstance(value, str):
            wb.set_cell_string(self._sheet_idx, r, c, value)
        elif value is None:
            wb.set_cell_empty(self._sheet_idx, r, c)
        else:
            raise TypeError(
                f"Unsupported cell value type: {type(value).__name__}. "
                f"Supported types: str, int, float, bool, datetime, date, time, None"
            )

    def __setitem__(self, key, value):
        row, col = _parse_cell_ref(key)
        if (row, col) in self._cells:
            self._cells[(row, col)].value = value
        else:
            self._cells[(row, col)] = Cell(row=row, column=col, value=value)
        self._set_rust_value(row, col, value)

    def __getitem__(self, key):
        # Integer key: return tuple of cells for that row (openpyxl compat)
        if isinstance(key, int):
            min_col = self.min_column or 1
            max_col = self.max_column or 1
            return tuple(
                self.cell(row=key, column=col)
                for col in range(min_col, max_col + 1)
            )
        # Range access like 'A1:C3'
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
        # Single cell access like 'A1'
        row, col = _parse_cell_ref(key)
        if (row, col) not in self._cells:
            self._cells[(row, col)] = Cell(row=row, column=col)
        return self._cells[(row, col)]

    def cell(self, row, column, value=None):
        if (row, column) in self._cells:
            cell = self._cells[(row, column)]
            if value is not None:
                cell.value = value
                self._set_rust_value(row, column, value)
            return cell
        c = Cell(row=row, column=column, value=value)
        self._cells[(row, column)] = c
        if value is not None:
            self._set_rust_value(row, column, value)
        return c

    def _next_row(self):
        """Return the next available 1-based row number for append operations.

        Takes into account both cells written via cell()/setitem and prior appends.
        """
        if self._cells:
            max_cell_row = max(r for r, _ in self._cells)
        else:
            max_cell_row = 0
        return max(self._current_row, max_cell_row) + 1

    def append(self, iterable):
        """Append a single row of values, matching openpyxl's API.

        Values are stored in self._cells AND pushed to Rust via _set_rust_value.
        """
        row = self._next_row()
        for col_idx, value in enumerate(iterable, start=1):
            c = Cell(row=row, column=col_idx, value=value)
            self._cells[(row, col_idx)] = c
            self._set_rust_value(row, col_idx, value)
        self._current_row = row

    def append_rows(self, rows_data):
        """Append multiple rows at once via the Rust batch API for maximum speed.

        Each element of rows_data should be a list/tuple of cell values.
        Values are also stored in self._cells for format access.
        Datetime/date/time values are pre-converted to Excel serial floats for Rust,
        but original objects are kept in _cells so _flush_formats can detect them.
        """
        start_row = self._next_row()  # 1-based
        start_row_0based = start_row - 1

        # Push all data to Rust in a single batch call
        if self._workbook is not None and self._sheet_idx is not None:
            # Convert to list-of-lists for Rust, pre-processing datetime values
            rows_list = []
            for row in rows_data:
                converted_row = []
                for value in row:
                    # datetime must come before date because datetime is subclass of date
                    if isinstance(value, datetime):
                        serial = _date_to_excel_serial(value.year, value.month, value.day)
                        serial += (value.hour * 3600 + value.minute * 60 + value.second) / 86400.0
                        converted_row.append(serial)
                    elif isinstance(value, date):
                        serial = _date_to_excel_serial(value.year, value.month, value.day)
                        converted_row.append(serial)
                    elif isinstance(value, time):
                        serial = (value.hour * 3600 + value.minute * 60 + value.second) / 86400.0
                        converted_row.append(serial)
                    else:
                        converted_row.append(value)
                rows_list.append(converted_row)
            self._workbook._rust_wb.set_rows_batch(
                self._sheet_idx, start_row_0based, rows_list
            )

        # Also store original values in self._cells for Python-side access
        # (keeping original datetime objects so _flush_formats can detect them)
        for row_offset, row_values in enumerate(rows_data):
            row = start_row + row_offset
            for col_idx, value in enumerate(row_values, start=1):
                c = Cell(row=row, column=col_idx, value=value)
                self._cells[(row, col_idx)] = c
            self._current_row = row

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
        """Unmerge a previously merged cell range.

        Raises ValueError if the range is not currently merged.
        """
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
        """Clear and re-push all cell values and merge ranges to Rust after structural changes."""
        if self._workbook is None or self._sheet_idx is None:
            return
        wb = self._workbook._rust_wb
        idx = self._sheet_idx
        wb.clear_cells(idx)
        wb.clear_merge_ranges(idx)
        for (row, col), cell in self._cells.items():
            self._set_rust_value(row, col, cell.value)
        for (start_ref, end_ref) in self.merged_cell_ranges:
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            wb.add_merge_range(idx, r1 - 1, c1 - 1, r2 - 1, c2 - 1)

    def insert_rows(self, idx, amount=1):
        """Insert `amount` rows before row `idx`, shifting existing rows down."""
        new_cells = {}
        for (row, col), cell in self._cells.items():
            if row >= idx:
                new_row = row + amount
                cell.row = new_row
                new_cells[(new_row, col)] = cell
            else:
                new_cells[(row, col)] = cell
        self._cells = new_cells
        # Shift merged ranges
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
        self._resync_rust()

    def delete_rows(self, idx, amount=1):
        """Delete `amount` rows starting at row `idx`, shifting remaining rows up."""
        new_cells = {}
        for (row, col), cell in self._cells.items():
            if idx <= row < idx + amount:
                continue  # deleted
            elif row >= idx + amount:
                new_row = row - amount
                cell.row = new_row
                new_cells[(new_row, col)] = cell
            else:
                new_cells[(row, col)] = cell
        self._cells = new_cells
        new_merged = []
        for (start_ref, end_ref) in self.merged_cell_ranges:
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            if r1 >= idx and r2 < idx + amount:
                continue  # fully within deleted range
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
        self._resync_rust()

    def insert_cols(self, idx, amount=1):
        """Insert `amount` columns before column `idx`, shifting existing columns right."""
        new_cells = {}
        for (row, col), cell in self._cells.items():
            if col >= idx:
                new_col = col + amount
                cell.column = new_col
                new_cells[(row, new_col)] = cell
            else:
                new_cells[(row, col)] = cell
        self._cells = new_cells
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
        """Delete `amount` columns starting at column `idx`, shifting remaining columns left."""
        new_cells = {}
        for (row, col), cell in self._cells.items():
            if idx <= col < idx + amount:
                continue  # deleted
            elif col >= idx + amount:
                new_col = col - amount
                cell.column = new_col
                new_cells[(row, new_col)] = cell
            else:
                new_cells[(row, col)] = cell
        self._cells = new_cells
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
        """Add an image to the worksheet. anchor is a cell reference like 'A1'."""
        if anchor is not None:
            img.anchor = anchor
        self._images.append(img)

    def add_data_validation(self, dv):
        """Add a DataValidation to this worksheet."""
        self._data_validations.append(dv)

    def _flush_formats(self):
        """Push all cell formats, column widths, row heights, freeze panes to Rust.
        Called right before save."""
        if self._workbook is None or self._sheet_idx is None:
            return
        wb = self._workbook._rust_wb
        idx = self._sheet_idx

        # Push cell formats
        format_batch = []
        for (row, col), cell in self._cells.items():
            fmt = {}
            if cell.font is not None:
                font_data = {
                    "bold": cell.font.bold,
                    "italic": cell.font.italic,
                    "name": cell.font.name,
                    "size": cell.font.size,
                }
                if cell.font.underline is not None:
                    font_data["underline"] = cell.font.underline
                if cell.font.color is not None:
                    font_data["color"] = cell.font.color
                if cell.font.strikethrough:
                    font_data["strikethrough"] = True
                if cell.font.vertAlign is not None:
                    font_data["vertAlign"] = cell.font.vertAlign
                fmt["font"] = font_data
            if cell.alignment is not None:
                align_data = {}
                if cell.alignment.horizontal is not None:
                    align_data["horizontal"] = cell.alignment.horizontal
                if cell.alignment.vertical is not None:
                    align_data["vertical"] = cell.alignment.vertical
                if cell.alignment.wrap_text:
                    align_data["wrap_text"] = True
                if cell.alignment.shrink_to_fit:
                    align_data["shrink_to_fit"] = True
                if cell.alignment.indent:
                    align_data["indent"] = cell.alignment.indent
                if cell.alignment.text_rotation:
                    align_data["text_rotation"] = cell.alignment.text_rotation
                if align_data:
                    fmt["alignment"] = align_data
            if cell.border is not None:
                border_data = {}
                for side_name in ("left", "right", "top", "bottom"):
                    side = getattr(cell.border, side_name)
                    if side.style is not None:
                        side_data = {"style": side.style}
                        if side.color is not None:
                            side_data["color"] = side.color
                        border_data[side_name] = side_data
                if cell.border.diagonal.style is not None:
                    diag_data = {"style": cell.border.diagonal.style}
                    if cell.border.diagonal.color is not None:
                        diag_data["color"] = cell.border.diagonal.color
                    diag_data["diagonalUp"] = cell.border.diagonalUp
                    diag_data["diagonalDown"] = cell.border.diagonalDown
                    border_data["diagonal"] = diag_data
                if border_data:
                    fmt["border"] = border_data
            if cell.fill is not None:
                fill_data = {}
                if cell.fill.fill_type is not None:
                    fill_data["fill_type"] = cell.fill.fill_type
                if cell.fill.start_color is not None:
                    fill_data["start_color"] = cell.fill.start_color
                if cell.fill.end_color is not None:
                    fill_data["end_color"] = cell.fill.end_color
                if fill_data:
                    fmt["fill"] = fill_data

            # Handle number format for datetime cells
            # datetime check must come before date because datetime is subclass of date
            if isinstance(cell.value, datetime):
                if cell.number_format == "General":
                    fmt["number_format"] = "yyyy-mm-dd hh:mm:ss"
                else:
                    fmt["number_format"] = cell.number_format
            elif isinstance(cell.value, date):
                if cell.number_format == "General":
                    fmt["number_format"] = "yyyy-mm-dd"
                else:
                    fmt["number_format"] = cell.number_format
            elif isinstance(cell.value, time):
                if cell.number_format == "General":
                    fmt["number_format"] = "hh:mm:ss"
                else:
                    fmt["number_format"] = cell.number_format
            elif cell.number_format != "General":
                fmt["number_format"] = cell.number_format

            if fmt:
                format_batch.append((row - 1, col - 1, json.dumps(fmt)))

        if format_batch:
            wb.set_cell_formats_batch(idx, format_batch)

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

        # Hyperlinks
        for (row, col), cell in self._cells.items():
            if cell.hyperlink is not None:
                url = cell.hyperlink
                text = None
                tooltip = None
                # Handle internal links (openpyxl uses # prefix)
                if isinstance(url, str) and url.startswith("#"):
                    url = "internal:" + url[1:]
                wb.add_hyperlink(idx, row - 1, col - 1, url, text, tooltip)

        # Comments/Notes
        for (row, col), cell in self._cells.items():
            if cell.comment is not None:
                author = cell.comment.author if cell.comment.author else None
                wb.add_note(idx, row - 1, col - 1, cell.comment.text, author)

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
        # Margins
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
            # Collect all cell ranges
            ranges = []
            # From sqref
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
            # From _cells list (from dv.add())
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
        """Serialize font/fill/border from a CellIsRule or FormulaRule into a format dict."""
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
        """Serialize a conditional formatting rule into a JSON string."""
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
                "bar_only": rule.showValue is False,  # showValue=False means bar only
            }
            return json.dumps(data)

        elif isinstance(rule, IconSetRule):
            data = {
                "rule_type": "icon_set",
                "range": range_string,
                "icon_style": rule.icon_style or "3TrafficLights1",
                "reverse": bool(rule.reverse),
                "show_icons_only": rule.showValue is False,  # showValue=False means icons only
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

