class Font:
    name: str
    size: int | float
    bold: bool
    italic: bool
    underline: str | None
    color: str | None
    strikethrough: bool
    vertAlign: str | None
    def __init__(
        self,
        name: str = "Calibri",
        size: int | float = 11,
        bold: bool = False,
        italic: bool = False,
        underline: str | None = None,
        color: str | None = None,
        strikethrough: bool = False,
        vertAlign: str | None = None,
    ) -> None: ...
    def __eq__(self, other: object) -> bool: ...
    def __repr__(self) -> str: ...
