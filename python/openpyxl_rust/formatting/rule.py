class ColorScaleRule:
    def __init__(
        self,
        start_type=None,
        start_value=None,
        start_color=None,
        mid_type=None,
        mid_value=None,
        mid_color=None,
        end_type=None,
        end_value=None,
        end_color=None,
    ):
        self.start_type = start_type
        self.start_value = start_value
        self.start_color = start_color
        self.mid_type = mid_type
        self.mid_value = mid_value
        self.mid_color = mid_color
        self.end_type = end_type
        self.end_value = end_value
        self.end_color = end_color


class DataBarRule:
    def __init__(
        self,
        start_type=None,
        start_value=None,
        end_type=None,
        end_value=None,
        color=None,
        showValue=None,
        minLength=None,
        maxLength=None,
    ):
        self.start_type = start_type
        self.start_value = start_value
        self.end_type = end_type
        self.end_value = end_value
        self.color = color
        self.showValue = showValue
        self.minLength = minLength
        self.maxLength = maxLength


class IconSetRule:
    def __init__(self, icon_style=None, type=None, values=None, showValue=None, percent=None, reverse=None):
        self.icon_style = icon_style
        self.type = type
        self.values = values or []
        self.showValue = showValue
        self.percent = percent
        self.reverse = reverse


class CellIsRule:
    def __init__(self, operator=None, formula=None, stopIfTrue=None, font=None, border=None, fill=None):
        self.operator = operator
        self.formula = formula or []
        self.stopIfTrue = stopIfTrue
        self.font = font
        self.border = border
        self.fill = fill


class FormulaRule:
    def __init__(self, formula=None, stopIfTrue=None, font=None, border=None, fill=None):
        self.formula = formula or []
        self.stopIfTrue = stopIfTrue
        self.font = font
        self.border = border
        self.fill = fill
