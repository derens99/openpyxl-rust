def _col_letter(col_idx):
    """Convert 1-based column index to Excel column letter(s). 1->A, 27->AA."""
    result = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


class Cell:
    # Type constants matching openpyxl
    TYPE_STRING = 's'
    TYPE_FORMULA = 'f'
    TYPE_NUMERIC = 'n'
    TYPE_BOOL = 'b'
    TYPE_NULL = 'n'
    TYPE_INLINE = 's'
    TYPE_ERROR = 'e'
    TYPE_FORMULA_CACHE_STRING = 's'

    def __init__(self, row=1, column=1, value=None):
        self.row = row
        self.column = column
        self.value = value
        self.font = None
        self.number_format = "General"
        self.alignment = None
        self.border = None
        self.fill = None
        self.hyperlink = None
        self.comment = None

    @property
    def coordinate(self):
        return f"{_col_letter(self.column)}{self.row}"

    @property
    def data_type(self):
        from datetime import datetime, date, time
        v = self.value
        if v is None:
            return 'n'  # openpyxl returns 'n' for empty cells
        if isinstance(v, bool):
            return 'b'
        if isinstance(v, (int, float)):
            return 'n'
        if isinstance(v, (datetime, date, time)):
            return 'd'
        if isinstance(v, str):
            if v.startswith('='):
                return 'f'
            return 's'
        return 'n'
