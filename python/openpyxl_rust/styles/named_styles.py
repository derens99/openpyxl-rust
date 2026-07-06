class NamedStyle:
    """A reusable, named collection of style components.

    Register with a workbook via ``wb.add_named_style(style)`` (or by assigning
    the style object directly to ``cell.style``, which auto-registers it), then
    apply to cells by name: ``cell.style = "highlight"``.
    """

    def __init__(
        self,
        name="Normal",
        font=None,
        fill=None,
        border=None,
        alignment=None,
        number_format=None,
        protection=None,
        builtinId=None,
        hidden=False,
    ):
        self.name = name
        self.font = font
        self.fill = fill
        self.border = border
        self.alignment = alignment
        self.number_format = number_format
        self.protection = protection
        self.builtinId = builtinId
        self.hidden = hidden

    def __eq__(self, other):
        if not isinstance(other, NamedStyle):
            return NotImplemented
        return (
            self.name == other.name
            and self.font == other.font
            and self.fill == other.fill
            and self.border == other.border
            and self.alignment == other.alignment
            and self.number_format == other.number_format
            and self.protection == other.protection
        )

    def __repr__(self):
        return (
            f"NamedStyle(name={self.name!r}, font={self.font!r}, fill={self.fill!r}, "
            f"border={self.border!r}, alignment={self.alignment!r}, "
            f"number_format={self.number_format!r}, protection={self.protection!r})"
        )
