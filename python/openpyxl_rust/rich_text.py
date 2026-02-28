"""Rich text support for openpyxl-rust."""


class TextBlock:
    """A segment of rich text with optional formatting.

    Args:
        font: A Font object for formatting, or a plain string (shorthand).
        text: The text content of this segment.

    Usage:
        TextBlock(Font(bold=True), "Bold text")
        TextBlock("plain text")  # shorthand for no formatting
        TextBlock(text="plain text")  # keyword-only
    """

    def __init__(self, font=None, text=""):
        if isinstance(font, str):
            # Allow TextBlock("plain text") shorthand
            self.text = font
            self.font = None
        else:
            self.font = font
            self.text = text

    def __str__(self):
        return self.text


class CellRichText:
    """Rich text content for a cell, allowing mixed formatting.

    Usage:
        from openpyxl_rust.styles import Font
        rt = CellRichText(
            "Normal text ",
            TextBlock(Font(bold=True), "Bold text"),
            " more normal"
        )
        ws["A1"].value = rt
    """

    def __init__(self, *args):
        self.segments = []
        for arg in args:
            if isinstance(arg, str):
                self.segments.append(TextBlock(text=arg))
            elif isinstance(arg, TextBlock):
                self.segments.append(arg)
            else:
                raise TypeError(f"Expected str or TextBlock, got {type(arg).__name__}")

    def __str__(self):
        return "".join(str(s) for s in self.segments)

    def _to_json_segments(self):
        """Convert to JSON-serializable list of segment dicts."""
        result = []
        for seg in self.segments:
            d = {"text": seg.text}
            if seg.font is not None:
                if seg.font.bold:
                    d["bold"] = True
                if seg.font.italic:
                    d["italic"] = True
                if seg.font.name:
                    d["name"] = seg.font.name
                if seg.font.size:
                    d["size"] = seg.font.size
                if seg.font.color:
                    d["color"] = seg.font.color
                if seg.font.strikethrough:
                    d["strikethrough"] = True
                if seg.font.underline:
                    d["underline"] = seg.font.underline
            result.append(d)
        return result
