class Side:
    def __init__(self, style=None, color=None):
        self.style = style
        self.color = color

    def __eq__(self, other):
        if not isinstance(other, Side):
            return NotImplemented
        return self.style == other.style and self.color == other.color

    def __repr__(self):
        return f"Side(style={self.style!r}, color={self.color!r})"


class Border:
    def __init__(self, left=None, right=None, top=None, bottom=None,
                 diagonal=None, diagonalUp=False, diagonalDown=False):
        self.left = left or Side()
        self.right = right or Side()
        self.top = top or Side()
        self.bottom = bottom or Side()
        self.diagonal = diagonal or Side()
        self.diagonalUp = diagonalUp
        self.diagonalDown = diagonalDown

    def __eq__(self, other):
        if not isinstance(other, Border):
            return NotImplemented
        return (self.left == other.left and self.right == other.right
                and self.top == other.top and self.bottom == other.bottom
                and self.diagonal == other.diagonal
                and self.diagonalUp == other.diagonalUp
                and self.diagonalDown == other.diagonalDown)

    def __repr__(self):
        return (f"Border(left={self.left!r}, right={self.right!r}, top={self.top!r}, "
                f"bottom={self.bottom!r}, diagonal={self.diagonal!r}, "
                f"diagonalUp={self.diagonalUp}, diagonalDown={self.diagonalDown})")
