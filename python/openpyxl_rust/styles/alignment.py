class Alignment:
    def __init__(self, horizontal=None, vertical=None, wrap_text=False, shrink_to_fit=False, indent=0, text_rotation=0):
        self.horizontal = horizontal
        self.vertical = vertical
        self.wrap_text = wrap_text
        self.shrink_to_fit = shrink_to_fit
        self.indent = indent
        self.text_rotation = text_rotation

    def __eq__(self, other):
        if not isinstance(other, Alignment):
            return NotImplemented
        return (
            self.horizontal == other.horizontal
            and self.vertical == other.vertical
            and self.wrap_text == other.wrap_text
            and self.shrink_to_fit == other.shrink_to_fit
            and self.indent == other.indent
            and self.text_rotation == other.text_rotation
        )

    def __repr__(self):
        return f"Alignment(horizontal={self.horizontal!r}, vertical={self.vertical!r}, wrap_text={self.wrap_text})"
