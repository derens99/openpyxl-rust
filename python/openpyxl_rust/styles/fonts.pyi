from typing import Optional, Union

class Font:
    name: str
    size: Union[int, float]
    bold: bool
    italic: bool
    underline: Optional[str]
    color: Optional[str]
    strikethrough: bool
    vertAlign: Optional[str]
    def __init__(
        self,
        name: str = "Calibri",
        size: Union[int, float] = 11,
        bold: bool = False,
        italic: bool = False,
        underline: Optional[str] = None,
        color: Optional[str] = None,
        strikethrough: bool = False,
        vertAlign: Optional[str] = None,
    ) -> None: ...
    def __eq__(self, other: object) -> bool: ...
    def __repr__(self) -> str: ...
