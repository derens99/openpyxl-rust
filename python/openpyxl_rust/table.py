"""Table support for openpyxl-rust, compatible with openpyxl's Table API."""


class TableStyleInfo:
    """Style information for an Excel table."""

    def __init__(self, name="TableStyleMedium9", showFirstColumn=False,
                 showLastColumn=False, showRowStripes=True, showColumnStripes=False):
        self.name = name
        self.showFirstColumn = showFirstColumn
        self.showLastColumn = showLastColumn
        self.showRowStripes = showRowStripes
        self.showColumnStripes = showColumnStripes


class TableColumn:
    """A column definition for an Excel table."""

    def __init__(self, id=None, name=""):
        self.id = id
        self.name = name


class Table:
    """An Excel table (ListObject).

    Usage (openpyxl-compatible):
        tab = Table(displayName="Sales", ref="A1:D10")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1")
        ws.add_table(tab)
    """

    def __init__(self, displayName=None, ref=None):
        self.displayName = displayName
        self.ref = ref  # e.g. "A1:D10"
        self.tableStyleInfo = None
        self.tableColumns = []
        self.headerRowCount = 1  # 0 = no headers
        self.totalsRowCount = 0
        self.totalsRowShown = False
