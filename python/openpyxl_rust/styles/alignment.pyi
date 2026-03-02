class Alignment:
    horizontal: str | None
    vertical: str | None
    wrap_text: bool
    shrink_to_fit: bool
    indent: int | float
    text_rotation: int | float
    def __init__(
        self,
        horizontal: str | None = None,
        vertical: str | None = None,
        wrap_text: bool = False,
        shrink_to_fit: bool = False,
        indent: int | float = 0,
        text_rotation: int | float = 0,
    ) -> None: ...
    def __eq__(self, other: object) -> bool: ...
    def __repr__(self) -> str: ...
