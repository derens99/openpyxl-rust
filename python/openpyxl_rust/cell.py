from datetime import date, datetime, time


def _col_letter(col_idx):
    """Convert 1-based column index to Excel column letter(s). 1->A, 27->AA."""
    result = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


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


# ---------- encoding maps for Rust FFI ----------

_BORDER_STYLE_MAP = {
    "thin": 1,
    "medium": 2,
    "thick": 3,
    "dashed": 4,
    "dotted": 5,
    "double": 6,
    "hair": 7,
    "mediumDashed": 8,
    "dashDot": 9,
    "mediumDashDot": 10,
    "dashDotDot": 11,
    "mediumDashDotDot": 12,
    "slantDashDot": 13,
}

_FILL_TYPE_MAP = {
    "solid": 1,
    "darkGray": 2,
    "mediumGray": 3,
    "lightGray": 4,
    "gray125": 5,
    "gray0625": 6,
}

_HALIGN_MAP = {
    "left": 1,
    "center": 2,
    "right": 3,
    "fill": 4,
    "justify": 5,
    "centerContinuous": 6,
    "center_continuous": 6,
    "distributed": 7,
}

_VALIGN_MAP = {
    "top": 1,
    "center": 2,
    "bottom": 3,
    "justify": 4,
    "distributed": 5,
}


def _underline_to_u8(val):
    """Convert an openpyxl underline value to a u8 for Rust."""
    if val is None:
        return None
    mapping = {
        "single": 1,
        "double": 2,
        "singleAccounting": 3,
        "doubleAccounting": 4,
    }
    return mapping.get(val, 1)


def _vert_align_to_u8(val):
    """Convert openpyxl vertAlign ('superscript', 'subscript', 'baseline') to u8."""
    if val is None:
        return None
    mapping = {
        "superscript": 1,
        "subscript": 2,
        "baseline": 3,
    }
    return mapping.get(val)


class Cell:
    __slots__ = (
        "_alignment",
        "_border",
        "_col",
        "_comment",
        "_fill",
        "_font",
        "_hyperlink",
        "_number_format",
        "_row",
        "_value",
        "_ws",
    )

    TYPE_STRING = "s"
    TYPE_FORMULA = "f"
    TYPE_NUMERIC = "n"
    TYPE_BOOL = "b"
    TYPE_NULL = "n"
    TYPE_INLINE = "s"
    TYPE_ERROR = "e"
    TYPE_FORMULA_CACHE_STRING = "s"

    def __init__(self, row=1, column=1, value=None, worksheet=None):
        self._ws = worksheet
        self._row = row
        self._col = column
        self._value = None
        self._font = None
        self._number_format = "General"
        self._alignment = None
        self._border = None
        self._fill = None
        self._hyperlink = None
        self._comment = None
        if value is not None:
            if worksheet is not None and worksheet._workbook is not None:
                worksheet._set_cell_value(row, column, value)
            else:
                self._value = value

    @property
    def row(self):
        return self._row

    @row.setter
    def row(self, val):
        self._row = val

    @property
    def column(self):
        return self._col

    @column.setter
    def column(self, val):
        self._col = val

    @property
    def value(self):
        if self._ws is not None and self._ws._workbook is not None:
            return self._ws._get_cell_value(self._row, self._col)
        return self._value

    @value.setter
    def value(self, val):
        if self._ws is not None and self._ws._workbook is not None:
            self._ws._set_cell_value(self._row, self._col, val)
        else:
            self._value = val

    @property
    def coordinate(self):
        return f"{_col_letter(self._col)}{self._row}"

    @property
    def data_type(self):
        v = self.value
        if v is None:
            return "n"
        if isinstance(v, bool):
            return "b"
        if isinstance(v, (int, float)):
            return "n"
        if isinstance(v, (datetime, date, time)):
            return "d"
        if isinstance(v, str):
            if v.startswith("="):
                return "f"
            return "s"
        return "n"

    @property
    def font(self):
        return self._font

    @font.setter
    def font(self, val):
        self._font = val
        self._mark_formatted()

    @property
    def number_format(self):
        return self._number_format

    @number_format.setter
    def number_format(self, val):
        self._number_format = val
        self._mark_formatted()

    @property
    def alignment(self):
        return self._alignment

    @alignment.setter
    def alignment(self, val):
        self._alignment = val
        self._mark_formatted()

    @property
    def border(self):
        return self._border

    @border.setter
    def border(self, val):
        self._border = val
        self._mark_formatted()

    @property
    def fill(self):
        return self._fill

    @fill.setter
    def fill(self, val):
        self._fill = val
        self._mark_formatted()

    @property
    def hyperlink(self):
        return self._hyperlink

    @hyperlink.setter
    def hyperlink(self, val):
        self._hyperlink = val
        self._mark_formatted()

    @property
    def comment(self):
        return self._comment

    @comment.setter
    def comment(self, val):
        self._comment = val
        self._mark_formatted()

    def _mark_formatted(self):
        """Register this cell in the worksheet's formatted cells tracking."""
        if self._ws is not None:
            self._ws._formatted_cells[(self._row, self._col)] = self
