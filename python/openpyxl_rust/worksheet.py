import re
from datetime import datetime, date
from openpyxl_rust.cell import Cell, _col_letter


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


class Worksheet:
    def __init__(self, title="Sheet"):
        self.title = title
        self._cells = {}  # (row, col) -> Cell
        self.column_dimensions = _ColumnDimensionsDict()
        self.row_dimensions = _RowDimensionsDict()
        self.freeze_panes = None
        self.merged_cell_ranges = []

    def __setitem__(self, key, value):
        row, col = _parse_cell_ref(key)
        if (row, col) in self._cells:
            self._cells[(row, col)].value = value
        else:
            self._cells[(row, col)] = Cell(row=row, column=col, value=value)

    def __getitem__(self, key):
        row, col = _parse_cell_ref(key)
        if (row, col) not in self._cells:
            self._cells[(row, col)] = Cell(row=row, column=col)
        return self._cells[(row, col)]

    def cell(self, row, column, value=None):
        if (row, column) in self._cells:
            cell = self._cells[(row, column)]
            if value is not None:
                cell.value = value
            return cell
        c = Cell(row=row, column=column, value=value)
        self._cells[(row, column)] = c
        return c

    def merge_cells(self, range_string):
        parts = range_string.split(":")
        if len(parts) != 2:
            raise ValueError(f"Invalid merge range: {range_string}")
        self.merged_cell_ranges.append((parts[0].upper(), parts[1].upper()))

    def _to_save_dict(self):
        """Serialize worksheet data for the Rust save engine."""
        cells = []
        for (row, col), cell in self._cells.items():
            cell_data = {
                "row": row - 1,   # Rust uses 0-based
                "col": col - 1,   # Rust uses 0-based
            }

            # Handle datetime/date values (datetime check MUST come first
            # because datetime is a subclass of date)
            if isinstance(cell.value, datetime):
                cell_data["value"] = {"__type__": "datetime",
                                      "year": cell.value.year,
                                      "month": cell.value.month,
                                      "day": cell.value.day,
                                      "hour": cell.value.hour,
                                      "minute": cell.value.minute,
                                      "second": cell.value.second}
                if cell.number_format == "General":
                    cell_data["number_format"] = "yyyy-mm-dd hh:mm:ss"
            elif isinstance(cell.value, date):
                cell_data["value"] = {"__type__": "date",
                                      "year": cell.value.year,
                                      "month": cell.value.month,
                                      "day": cell.value.day}
                if cell.number_format == "General":
                    cell_data["number_format"] = "yyyy-mm-dd"
            else:
                cell_data["value"] = cell.value
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
                cell_data["font"] = font_data
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
                    cell_data["alignment"] = align_data
            if cell.border is not None:
                border_data = {}
                for side_name in ("left", "right", "top", "bottom"):
                    side = getattr(cell.border, side_name)
                    if side.style is not None:
                        side_data = {"style": side.style}
                        if side.color is not None:
                            side_data["color"] = side.color
                        border_data[side_name] = side_data
                if border_data:
                    cell_data["border"] = border_data
            if cell.fill is not None:
                fill_data = {}
                if cell.fill.fill_type is not None:
                    fill_data["fill_type"] = cell.fill.fill_type
                if cell.fill.start_color is not None:
                    fill_data["start_color"] = cell.fill.start_color
                if cell.fill.end_color is not None:
                    fill_data["end_color"] = cell.fill.end_color
                if fill_data:
                    cell_data["fill"] = fill_data
            if cell.number_format != "General" and "number_format" not in cell_data:
                cell_data["number_format"] = cell.number_format
            cells.append(cell_data)

        # Column widths: convert letter key to 0-based index
        col_widths = {}
        for letter, dim in self.column_dimensions.items():
            if dim.width is not None:
                _, col_idx = _parse_cell_ref(f"{letter}1")
                col_widths[col_idx - 1] = dim.width

        # Row heights: convert 1-based to 0-based
        row_heights = {}
        for row_num, dim in self.row_dimensions.items():
            if dim.height is not None:
                row_heights[row_num - 1] = dim.height

        # Freeze panes
        freeze = None
        if self.freeze_panes:
            r, c = _parse_cell_ref(self.freeze_panes)
            freeze = [r - 1, c - 1]

        # Merged cells: convert refs to 0-based [r1, c1, r2, c2]
        merges = []
        for start_ref, end_ref in self.merged_cell_ranges:
            r1, c1 = _parse_cell_ref(start_ref)
            r2, c2 = _parse_cell_ref(end_ref)
            merges.append([r1 - 1, c1 - 1, r2 - 1, c2 - 1])

        result = {
            "title": self.title,
            "cells": cells,
        }
        if col_widths:
            result["column_widths"] = col_widths
        if row_heights:
            result["row_heights"] = row_heights
        if freeze is not None:
            result["freeze_panes"] = freeze
        if merges:
            result["merged_cells"] = merges

        return result
