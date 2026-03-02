"""Chart series, compatible with openpyxl's Series API."""


class Trendline:
    """A trendline for a chart series.

    Args:
        trendlineType: Type of trendline. One of "linear", "exponential",
            "polynomial", "power", "log", "movingAvg".
        displayEquation: If True, display the trendline equation on the chart.
        displayRSqr: If True, display the R-squared value on the chart.
    """

    def __init__(self, trendlineType="linear", displayEquation=False, displayRSqr=False):
        self.trendlineType = trendlineType
        self.displayEquation = displayEquation
        self.displayRSqr = displayRSqr


class DataLabelList:
    """Data labels configuration for a chart series.

    Args:
        showVal: If True, show the value on data labels.
        showCatName: If True, show the category name on data labels.
        showSerName: If True, show the series name on data labels.
    """

    def __init__(self, showVal=False, showCatName=False, showSerName=False):
        self.showVal = showVal
        self.showCatName = showCatName
        self.showSerName = showSerName


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
        self.trendline = None  # Trendline object or None
        self.dLbls = None  # DataLabelList object or None
        self.graphicalProperties = None  # For future use
