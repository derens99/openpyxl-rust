"""Chart series, compatible with openpyxl's Series API."""


class Series:
    """A data series for a chart.

    Usage:
        s = Series(values=ref_values, categories=ref_cats, title="Q1")
        chart.append(s)
    """

    def __init__(self, values=None, categories=None, title=None):
        self.values = values  # Reference
        self.categories = categories  # Reference or None
        self.title = title  # str or None
