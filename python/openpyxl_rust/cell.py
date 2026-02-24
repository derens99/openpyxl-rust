def _col_letter(col_idx):
    """Convert 1-based column index to Excel column letter(s). 1->A, 27->AA."""
    result = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


class Cell:
    def __init__(self, row=1, column=1, value=None):
        self.row = row
        self.column = column
        self.value = value
        self.font = None
        self.number_format = "General"

    @property
    def coordinate(self):
        return f"{_col_letter(self.column)}{self.row}"
