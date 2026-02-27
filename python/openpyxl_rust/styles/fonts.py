class Font:
    def __init__(
        self,
        name="Calibri",
        size=11,
        bold=False,
        italic=False,
        underline=None,
        color=None,
        strikethrough=False,
        vertAlign=None,
    ):
        self.name = name
        self.size = size
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.color = color
        self.strikethrough = strikethrough
        self.vertAlign = vertAlign

    def __eq__(self, other):
        if not isinstance(other, Font):
            return NotImplemented
        return (
            self.name == other.name
            and self.size == other.size
            and self.bold == other.bold
            and self.italic == other.italic
            and self.underline == other.underline
            and self.color == other.color
            and self.strikethrough == other.strikethrough
            and self.vertAlign == other.vertAlign
        )

    def __repr__(self):
        return (
            f"Font(name={self.name!r}, size={self.size}, bold={self.bold}, "
            f"italic={self.italic}, underline={self.underline!r}, color={self.color!r}, "
            f"strikethrough={self.strikethrough}, vertAlign={self.vertAlign!r})"
        )
