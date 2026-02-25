from typing import Optional

class PatternFill:
    fill_type: Optional[str]
    start_color: Optional[str]
    end_color: Optional[str]
    def __init__(
        self,
        fill_type: Optional[str] = None,
        start_color: Optional[str] = None,
        end_color: Optional[str] = None,
        fgColor: Optional[str] = None,
        bgColor: Optional[str] = None,
    ) -> None: ...
    @property
    def fgColor(self) -> Optional[str]: ...
    @fgColor.setter
    def fgColor(self, value: Optional[str]) -> None: ...
    @property
    def bgColor(self) -> Optional[str]: ...
    @bgColor.setter
    def bgColor(self, value: Optional[str]) -> None: ...
    def __eq__(self, other: object) -> bool: ...
    def __repr__(self) -> str: ...
