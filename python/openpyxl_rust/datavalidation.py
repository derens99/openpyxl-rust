from openpyxl_rust.worksheet import _parse_cell_ref


class DataValidation:
    def __init__(self, type=None, formula1=None, formula2=None,
                 allow_blank=False, showDropDown=False,
                 showErrorMessage=False, showInputMessage=False,
                 promptTitle=None, prompt=None,
                 errorTitle=None, error=None, errorStyle=None,
                 operator=None, sqref=None):
        self.type = type
        self.formula1 = formula1
        self.formula2 = formula2
        self.allow_blank = allow_blank
        self.showDropDown = showDropDown  # True = HIDE dropdown (Excel quirk)
        self.showErrorMessage = showErrorMessage
        self.showInputMessage = showInputMessage
        self.promptTitle = promptTitle
        self.prompt = prompt
        self.errorTitle = errorTitle
        self.error = error
        self.errorStyle = errorStyle
        self.operator = operator
        self.sqref = sqref or ""
        self._cells = []

    def add(self, cell_or_range):
        if isinstance(cell_or_range, str):
            self._cells.append(cell_or_range)
        else:
            self._cells.append(cell_or_range.coordinate)

    @property
    def cells(self):
        return self.sqref

    @cells.setter
    def cells(self, value):
        self.sqref = value
