"""Chart base class and concrete chart types, compatible with openpyxl API."""

from openpyxl_rust.chart.series import Series


class _ChartBase:
    """Base class for all chart types."""

    chart_type = None  # Override in subclasses

    def __init__(self):
        self.series = []
        self.title = None
        self.style = None
        self.x_axis_title = None
        self.y_axis_title = None
        self.width = 15      # cm (openpyxl default)
        self.height = 7.5    # cm (openpyxl default)
        self.legend = True
        self._anchor = None  # set by ws.add_chart(chart, "E5")

    def add_data(self, ref, titles_from_data=False, from_rows=False):
        """Add data from a Reference, splitting multi-column ranges into series.

        Args:
            ref: Reference object pointing to data range.
            titles_from_data: If True, first row of each column is the series name.
            from_rows: If True, read data row-wise instead of column-wise.
        """
        if from_rows:
            start_row = ref.min_row
            if titles_from_data:
                start_row += 1
            for r in range(ref.min_row, ref.max_row + 1):
                title = None
                if titles_from_data:
                    # Title is in the first column of that row
                    title = None  # Will be resolved from cell at flush time
                s = Series(values=_RowReference(ref.worksheet, r, ref.min_col, ref.max_col),
                           title=title)
                self.series.append(s)
        else:
            # Column-wise: one series per column
            start_col = ref.min_col
            if titles_from_data:
                # First row is header, data starts at min_row + 1
                data_start_row = ref.min_row + 1
            else:
                data_start_row = ref.min_row

            for col in range(start_col, ref.max_col + 1):
                title = None
                if titles_from_data:
                    # Title comes from first row of this column
                    title = _CellTitle(ref.worksheet, ref.min_row, col)
                from openpyxl_rust.chart.reference import Reference
                val_ref = Reference(
                    worksheet=ref.worksheet,
                    min_col=col, max_col=col,
                    min_row=data_start_row, max_row=ref.max_row,
                )
                s = Series(values=val_ref, title=title)
                self.series.append(s)

    def set_categories(self, ref):
        """Set category labels for all series that don't have them."""
        for s in self.series:
            if s.categories is None:
                s.categories = ref

    def append(self, series):
        """Add a pre-built Series to the chart."""
        self.series.append(series)


class _CellTitle:
    """Lazy title resolved from a cell reference at serialization time."""

    __slots__ = ('worksheet', 'row', 'col')

    def __init__(self, worksheet, row, col):
        self.worksheet = worksheet
        self.row = row
        self.col = col

    def resolve(self):
        """Get the string value of the title cell."""
        val = self.worksheet._get_cell_value(self.row, self.col)
        return str(val) if val is not None else None


class _RowReference:
    """A reference to a single row range (for from_rows=True)."""

    __slots__ = ('worksheet', 'row', 'min_col', 'max_col')

    def __init__(self, worksheet, row, min_col, max_col):
        self.worksheet = worksheet
        self.row = row
        self.min_col = min_col
        self.max_col = max_col
        # Make it look like a Reference for serialization
        self.min_row = row
        self.max_row = row


# ---- Concrete chart types ----

class BarChart(_ChartBase):
    """Vertical bar chart (Column chart in Excel terms).

    openpyxl's BarChart defaults to type="col" (vertical columns).
    Set self.type = "bar" for horizontal bars.
    """
    chart_type = "column"

    def __init__(self):
        super().__init__()
        self.type = "col"  # openpyxl compat: "col" = vertical, "bar" = horizontal
        self.grouping = "clustered"  # "clustered", "stacked", "percentStacked"


class BarChart3D(BarChart):
    """3D bar chart — renders same as BarChart in rust_xlsxwriter."""
    pass


class LineChart(_ChartBase):
    """Line chart."""
    chart_type = "line"

    def __init__(self):
        super().__init__()
        self.grouping = "standard"


class LineChart3D(LineChart):
    """3D line chart — renders same as LineChart in rust_xlsxwriter."""
    pass


class PieChart(_ChartBase):
    """Pie chart."""
    chart_type = "pie"


class PieChart3D(PieChart):
    """3D pie chart — renders same as PieChart in rust_xlsxwriter."""
    pass


class AreaChart(_ChartBase):
    """Area chart."""
    chart_type = "area"

    def __init__(self):
        super().__init__()
        self.grouping = "standard"


class AreaChart3D(AreaChart):
    """3D area chart — renders same as AreaChart in rust_xlsxwriter."""
    pass


class ScatterChart(_ChartBase):
    """Scatter (XY) chart."""
    chart_type = "scatter"


class DoughnutChart(_ChartBase):
    """Doughnut chart."""
    chart_type = "doughnut"


class RadarChart(_ChartBase):
    """Radar chart."""
    chart_type = "radar"


class StockChart(_ChartBase):
    """Stock chart."""
    chart_type = "stock"
