from openpyxl_rust.styles.alignment import Alignment
from openpyxl_rust.styles.borders import Border
from openpyxl_rust.styles.fills import PatternFill
from openpyxl_rust.styles.fonts import Font
from openpyxl_rust.styles.protection import Protection

class NamedStyle:
    name: str
    font: Font | None
    fill: PatternFill | None
    border: Border | None
    alignment: Alignment | None
    number_format: str | None
    protection: Protection | None
    builtinId: int | None
    hidden: bool
    def __init__(
        self,
        name: str = "Normal",
        font: Font | None = None,
        fill: PatternFill | None = None,
        border: Border | None = None,
        alignment: Alignment | None = None,
        number_format: str | None = None,
        protection: Protection | None = None,
        builtinId: int | None = None,
        hidden: bool = False,
    ) -> None: ...
    def __eq__(self, other: object) -> bool: ...
    def __repr__(self) -> str: ...
