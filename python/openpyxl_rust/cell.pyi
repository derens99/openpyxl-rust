from typing import Any, Optional

def _col_letter(col_idx: int) -> str: ...

class Cell:
    # Type constants (openpyxl compat)
    TYPE_STRING: str
    TYPE_FORMULA: str
    TYPE_NUMERIC: str
    TYPE_BOOL: str
    TYPE_NULL: str
    TYPE_INLINE: str
    TYPE_ERROR: str
    TYPE_FORMULA_CACHE_STRING: str

    row: int
    column: int
    value: Any
    font: Optional[Any]
    number_format: str
    alignment: Optional[Any]
    border: Optional[Any]
    fill: Optional[Any]
    hyperlink: Optional[str]
    comment: Optional[Any]
    def __init__(self, row: int = 1, column: int = 1, value: Any = None) -> None: ...
    @property
    def coordinate(self) -> str: ...
    @property
    def data_type(self) -> str: ...
