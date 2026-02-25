class PrintPageSetup:
    ORIENTATION_PORTRAIT = "portrait"
    ORIENTATION_LANDSCAPE = "landscape"

    def __init__(self):
        self.orientation = None
        self.paperSize = None
        self.scale = None
        self.fitToWidth = None
        self.fitToHeight = None


class PageMargins:
    def __init__(self, left=0.75, right=0.75, top=1.0, bottom=1.0, header=0.5, footer=0.5):
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.header = header
        self.footer = footer


class PrintOptions:
    def __init__(self):
        self.horizontalCentered = None
        self.verticalCentered = None
        self.gridLines = None
        self.headings = None
