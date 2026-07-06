import json
import os

from openpyxl_rust.properties import DocumentProperties
from openpyxl_rust.styles.named_styles import NamedStyle
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
        self._named_styles = {"Normal": NamedStyle(name="Normal", builtinId=0)}

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

    def _reindex_sheets(self):
        """Sync each worksheet's _sheet_idx with its position (mirrors the Rust sheet order)."""
        for i, ws in enumerate(self._sheets):
            ws._sheet_idx = i

    def create_sheet(self, title=None, index=None):
        title = title or f"Sheet{len(self._sheets) + 1}"
        title = self._unique_sheet_title(title)
        idx = self._rust_wb.add_sheet(title)
        ws = Worksheet(title=title, workbook=self, sheet_idx=idx)
        if index is None:
            self._sheets.append(ws)
        else:
            self._sheets.insert(index, ws)
            self._rust_wb.move_sheet(idx, self._sheets.index(ws))
            self._reindex_sheets()
        return ws

    def move_sheet(self, sheet, offset=0):
        """Move a worksheet (or sheet name) by offset within the workbook order."""
        if isinstance(sheet, str):
            sheet = self[sheet]
        if sheet not in self._sheets:
            raise ValueError("Worksheet is not part of this workbook")
        old = self._sheets.index(sheet)
        self._sheets.pop(old)
        self._sheets.insert(old + offset, sheet)
        self._rust_wb.move_sheet(old, self._sheets.index(sheet))
        self._reindex_sheets()

    def copy_worksheet(self, from_worksheet):
        """Copy a worksheet within this workbook.

        Like openpyxl, cell values, styles, dimensions, and merged cells are
        copied; images, charts, and tables are not.
        """
        from openpyxl_rust.cell import Cell

        if from_worksheet not in self._sheets:
            raise ValueError("Worksheet is not part of this workbook")
        title = self._unique_sheet_title(f"{from_worksheet.title} Copy")
        idx = self._rust_wb.clone_sheet(from_worksheet._sheet_idx, title)
        ws = Worksheet(title=title, workbook=self, sheet_idx=idx)
        # Rust clone already carries values and merges; copy the Python-side
        # mirrors directly (no re-adding to Rust).
        ws.merged_cell_ranges = list(from_worksheet.merged_cell_ranges)
        ws.freeze_panes = from_worksheet.freeze_panes
        for letter, dim in from_worksheet.column_dimensions.items():
            new_dim = ws.column_dimensions[letter]
            new_dim.width = dim.width
            new_dim.hidden = dim.hidden
            new_dim.outline_level = dim.outline_level
        for num, dim in from_worksheet.row_dimensions.items():
            new_dim = ws.row_dimensions[num]
            new_dim.height = dim.height
            new_dim.hidden = dim.hidden
            new_dim.outline_level = dim.outline_level
        # Formats live on Python proxies until save; clone them onto the copy.
        for (r, c), src in from_worksheet._formatted_cells.items():
            dst = Cell(row=r, column=c, worksheet=ws)
            dst._font = src._font
            dst._fill = src._fill
            dst._border = src._border
            dst._alignment = src._alignment
            dst._protection = src._protection
            dst._number_format = src._number_format
            dst._hyperlink = src._hyperlink
            dst._comment = src._comment
            dst._style_name = src._style_name
            ws._formatted_cells[(r, c)] = dst
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
    def named_styles(self):
        return list(self._named_styles)

    def add_named_style(self, style):
        """Register a NamedStyle so cells can reference it by name."""
        if style.name in self._named_styles:
            raise ValueError(f"Style {style.name} exists already")
        self._named_styles[style.name] = style

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
