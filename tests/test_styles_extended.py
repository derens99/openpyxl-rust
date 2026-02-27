from openpyxl_rust.styles import Alignment, Border, PatternFill, Side


def test_alignment_defaults():
    a = Alignment()
    assert a.horizontal is None
    assert a.vertical is None
    assert a.wrap_text is False
    assert a.shrink_to_fit is False
    assert a.indent == 0
    assert a.text_rotation == 0


def test_alignment_values():
    a = Alignment(horizontal="center", vertical="center", wrap_text=True)
    assert a.horizontal == "center"
    assert a.vertical == "center"
    assert a.wrap_text is True


def test_alignment_equality():
    a1 = Alignment(horizontal="center")
    a2 = Alignment(horizontal="center")
    assert a1 == a2


def test_side_defaults():
    s = Side()
    assert s.style is None
    assert s.color is None


def test_side_values():
    s = Side(style="thin", color="000000")
    assert s.style == "thin"
    assert s.color == "000000"


def test_border_defaults():
    b = Border()
    assert b.left.style is None
    assert b.right.style is None
    assert b.top.style is None
    assert b.bottom.style is None


def test_border_values():
    s = Side(style="thin", color="FF0000")
    b = Border(left=s, right=s, top=s, bottom=s)
    assert b.left.style == "thin"
    assert b.left.color == "FF0000"


def test_border_equality():
    s = Side(style="thin")
    b1 = Border(left=s)
    b2 = Border(left=Side(style="thin"))
    assert b1 == b2


def test_patternfill_defaults():
    f = PatternFill()
    assert f.fill_type is None
    assert f.start_color is None
    assert f.end_color is None


def test_patternfill_solid():
    f = PatternFill(fill_type="solid", start_color="FFFF00")
    assert f.fill_type == "solid"
    assert f.start_color == "FFFF00"


def test_patternfill_equality():
    f1 = PatternFill(fill_type="solid", start_color="FFFF00")
    f2 = PatternFill(fill_type="solid", start_color="FFFF00")
    assert f1 == f2


def test_border_diagonal_defaults():
    b = Border()
    assert b.diagonal.style is None
    assert b.diagonalUp is False
    assert b.diagonalDown is False


def test_border_diagonal_up():
    b = Border(diagonal=Side(style="thin", color="FF0000"), diagonalUp=True)
    assert b.diagonal.style == "thin"
    assert b.diagonalUp is True
    assert b.diagonalDown is False
