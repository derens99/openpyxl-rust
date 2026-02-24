from openpyxl_rust.worksheet import Worksheet


class Workbook:
    def __init__(self):
        self._sheets = [Worksheet(title="Sheet")]

    @property
    def active(self):
        return self._sheets[0] if self._sheets else None

    def create_sheet(self, title=None):
        title = title or f"Sheet{len(self._sheets) + 1}"
        ws = Worksheet(title=title)
        self._sheets.append(ws)
        return ws

    @property
    def sheetnames(self):
        return [s.title for s in self._sheets]

    def __getitem__(self, name):
        for s in self._sheets:
            if s.title == name:
                return s
        raise KeyError(f"Worksheet '{name}' not found")

    def save(self, filename):
        raise NotImplementedError("Rust save not yet wired up")
