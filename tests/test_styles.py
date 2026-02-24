from openpyxl_rust.styles import Font


def test_font_defaults():
    f = Font()
    assert f.bold is False
    assert f.italic is False
    assert f.underline is None
    assert f.name == "Calibri"
    assert f.size == 11


def test_font_bold():
    f = Font(bold=True, size=14, name="Arial")
    assert f.bold is True
    assert f.size == 14
    assert f.name == "Arial"


def test_font_color():
    f = Font(color="FF0000")
    assert f.color == "FF0000"


def test_font_equality():
    f1 = Font(bold=True)
    f2 = Font(bold=True)
    assert f1 == f2


def test_font_inequality():
    f1 = Font(bold=True)
    f2 = Font(bold=False)
    assert f1 != f2
