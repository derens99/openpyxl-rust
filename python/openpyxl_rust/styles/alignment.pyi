from typing import Optional, Union

class Alignment:
    horizontal: Optional[str]
    vertical: Optional[str]
    wrap_text: bool
    shrink_to_fit: bool
    indent: Union[int, float]
    text_rotation: Union[int, float]
    def __init__(
        self,
        horizontal: Optional[str] = None,
        vertical: Optional[str] = None,
        wrap_text: bool = False,
        shrink_to_fit: bool = False,
        indent: Union[int, float] = 0,
        text_rotation: Union[int, float] = 0,
    ) -> None: ...
    def __eq__(self, other: object) -> bool: ...
    def __repr__(self) -> str: ...
