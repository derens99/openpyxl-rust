class HeaderFooterItem:
    """Represents left, center, or right section of a header/footer."""
    def __init__(self):
        self.text = None


class HeaderFooter:
    """openpyxl-compatible header/footer with left, center, right sections."""
    def __init__(self):
        self.left = HeaderFooterItem()
        self.center = HeaderFooterItem()
        self.right = HeaderFooterItem()

    def _build_format_string(self):
        """Build OOXML header/footer format string: &LLeft&CCenter&RRight"""
        parts = []
        if self.left.text:
            parts.append(f"&L{self.left.text}")
        if self.center.text:
            parts.append(f"&C{self.center.text}")
        if self.right.text:
            parts.append(f"&R{self.right.text}")
        return "".join(parts) if parts else None

    def __bool__(self):
        return bool(self.left.text or self.center.text or self.right.text)
