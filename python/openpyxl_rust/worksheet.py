import re
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
