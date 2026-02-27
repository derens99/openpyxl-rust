class PatternFill:
    def __init__(self, fill_type=None, start_color=None, end_color=None, fgColor=None, bgColor=None):
        self.fill_type = fill_type
        self.start_color = fgColor or start_color  # fgColor takes precedence
        self.end_color = bgColor or end_color  # bgColor takes precedence

    @property
    def fgColor(self):
        return self.start_color

    @fgColor.setter
    def fgColor(self, value):
        self.start_color = value

    @property
    def bgColor(self):
        return self.end_color

    @bgColor.setter
    def bgColor(self, value):
        self.end_color = value

    def __eq__(self, other):
        if not isinstance(other, PatternFill):
            return NotImplemented
        return (
            self.fill_type == other.fill_type
            and self.start_color == other.start_color
            and self.end_color == other.end_color
        )

    def __repr__(self):
        return (
            f"PatternFill(fill_type={self.fill_type!r}, start_color={self.start_color!r}, end_color={self.end_color!r})"
        )
