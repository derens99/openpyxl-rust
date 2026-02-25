from typing import Optional

class Side:
    style: Optional[str]
    color: Optional[str]
    def __init__(
        self,
        style: Optional[str] = None,
        color: Optional[str] = None,
    ) -> None: ...
    def __eq__(self, other: object) -> bool: ...
    def __repr__(self) -> str: ...

class Border:
    left: Side
    right: Side
    top: Side
    bottom: Side
    diagonal: Side
    diagonalUp: bool
    diagonalDown: bool
    def __init__(
        self,
        left: Optional[Side] = None,
        right: Optional[Side] = None,
        top: Optional[Side] = None,
        bottom: Optional[Side] = None,
        diagonal: Optional[Side] = None,
        diagonalUp: bool = False,
        diagonalDown: bool = False,
    ) -> None: ...
    def __eq__(self, other: object) -> bool: ...
    def __repr__(self) -> str: ...
