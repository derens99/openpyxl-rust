import json
import os

from openpyxl_rust.properties import DocumentProperties
from openpyxl_rust.worksheet import Worksheet


class DefinedName:
    def __init__(self, name, attr_text=None):
        self.name = name
        self.attr_text = attr_text


class _DefinedNames:
    def __init__(self, workbook):
        self._wb = workbook
        self._names = {}

    def add(self, defined_name):
        self._names[defined_name.name] = defined_name
        self._wb._rust_wb.add_defined_name(defined_name.name, defined_name.attr_text)

    def __getitem__(self, name):
        return self._names[name]

    def __contains__(self, name):
        return name in self._names

    def __iter__(self):
        return iter(self._names.values())


class Workbook:
    def __init__(self):
        from openpyxl_rust._openpyxl_rust import RustWorkbook

        self._rust_wb = RustWorkbook()
        self._sheets = [Worksheet(title="Sheet", workbook=self, sheet_idx=0)]
        self._active_sheet_index = 0
        self.defined_names = _DefinedNames(self)
        self.properties = DocumentProperties()

    @property
    def active(self):
        if not self._sheets:
            return None
        idx = self._active_sheet_index
        if idx < 0 or idx >= len(self._sheets):
            idx = 0
        return self._sheets[idx]

    @active.setter
    def active(self, value):
        if isinstance(value, int):
            if value < 0 or value >= len(self._sheets):
                raise IndexError(f"Sheet index {value} is out of range (0-{len(self._sheets) - 1})")
            self._active_sheet_index = value
        elif isinstance(value, Worksheet):
            try:
                self._active_sheet_index = self._sheets.index(value)
            except ValueError as err:
                raise ValueError("Worksheet is not part of this workbook") from err
        else:
            raise TypeError("Value must be a Worksheet or an integer index")

    def __iter__(self):
        return iter(self._sheets)

    def __len__(self):
        return len(self._sheets)

    def _unique_sheet_title(self, title):
        """Return a unique sheet title, appending a number suffix if needed."""
        existing = set(self.sheetnames)
        if title not in existing:
            return title
        # Try appending incrementing numbers
        i = 1
        while f"{title}{i}" in existing:
            i += 1
        return f"{title}{i}"

    def create_sheet(self, title=None):
        title = title or f"Sheet{len(self._sheets) + 1}"
        title = self._unique_sheet_title(title)
        idx = self._rust_wb.add_sheet(title)
        ws = Worksheet(title=title, workbook=self, sheet_idx=idx)
        self._sheets.append(ws)
        return ws

    def remove(self, worksheet):
        """Remove a worksheet from this workbook."""
        if worksheet not in self._sheets:
            raise ValueError("Worksheet not found in this workbook")
        removed_idx = self._sheets.index(worksheet)
        self._sheets.remove(worksheet)
        self._rust_wb.remove_sheet(worksheet._sheet_idx)
        # Re-index remaining sheets
        for i, ws in enumerate(self._sheets):
            ws._sheet_idx = i
        # Adjust active sheet index
        if self._sheets:
            if self._active_sheet_index >= len(self._sheets):
                self._active_sheet_index = len(self._sheets) - 1
            elif self._active_sheet_index > removed_idx:
                self._active_sheet_index -= 1
        else:
            self._active_sheet_index = 0

    @property
    def sheetnames(self):
        return [s.title for s in self._sheets]

    def __getitem__(self, name):
        for s in self._sheets:
            if s.title == name:
                return s
        raise KeyError(f"Worksheet '{name}' not found")

    def save(self, filename):
        for ws in self._sheets:
            ws._flush_metadata()

        # Document properties
        props = self.properties
        props_data = {}
        if props.title:
            props_data["title"] = props.title
        if props.creator:
            props_data["creator"] = props.creator
        if props.description:
            props_data["description"] = props.description
        if props.subject:
            props_data["subject"] = props.subject
        if props.keywords:
            props_data["keywords"] = props.keywords
        if props.category:
            props_data["category"] = props.category
        if props_data:
            self._rust_wb.set_doc_properties(json.dumps(props_data))

        if isinstance(filename, (str, bytes, os.PathLike)):
            self._rust_wb.save(str(filename))
        else:
            # Assume file-like object (BytesIO etc.)
            result_bytes = self._rust_wb.save(None)
            filename.write(result_bytes)
