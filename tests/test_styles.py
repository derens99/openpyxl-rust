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


def test_font_strikethrough():
    f = Font(strikethrough=True)
    assert f.strikethrough is True


def test_font_strikethrough_default():
    f = Font()
    assert f.strikethrough is False


def test_font_vertAlign_superscript():
    f = Font(vertAlign="superscript")
    assert f.vertAlign == "superscript"


def test_font_vertAlign_default():
    f = Font()
    assert f.vertAlign is None


def test_font_equality_with_new_fields():
    f1 = Font(strikethrough=True, vertAlign="subscript")
    f2 = Font(strikethrough=True, vertAlign="subscript")
    assert f1 == f2
    f3 = Font(strikethrough=False)
    assert f1 != f3
