class Side:
    style: str | None
    color: str | None
    def __init__(
        self,
        style: str | None = None,
        color: str | None = None,
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
        left: Side | None = None,
        right: Side | None = None,
        top: Side | None = None,
        bottom: Side | None = None,
        diagonal: Side | None = None,
        diagonalUp: bool = False,
        diagonalDown: bool = False,
    ) -> None: ...
    def __eq__(self, other: object) -> bool: ...
    def __repr__(self) -> str: ...
