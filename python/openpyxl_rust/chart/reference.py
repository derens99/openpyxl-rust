"""Chart data reference, compatible with openpyxl's Reference API."""


class Reference:
    """A reference to a range of data for chart series.

    Usage (openpyxl-compatible):
        ref = Reference(ws, min_col=2, min_row=1, max_row=10)
        chart.add_data(ref, titles_from_data=True)
    """

    def __init__(self, worksheet=None, min_col=None, min_row=None, max_col=None, max_row=None):
        self.worksheet = worksheet
        self.min_col = min_col
        self.min_row = min_row
        self.max_col = max_col if max_col is not None else min_col
        self.max_row = max_row if max_row is not None else min_row
