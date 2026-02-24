class PatternFill:
    def __init__(self, fill_type=None, start_color=None, end_color=None):
        self.fill_type = fill_type
        self.start_color = start_color
        self.end_color = end_color

    def __eq__(self, other):
        if not isinstance(other, PatternFill):
            return NotImplemented
        return (self.fill_type == other.fill_type
                and self.start_color == other.start_color
                and self.end_color == other.end_color)

    def __repr__(self):
        return (f"PatternFill(fill_type={self.fill_type!r}, "
                f"start_color={self.start_color!r}, end_color={self.end_color!r})")
