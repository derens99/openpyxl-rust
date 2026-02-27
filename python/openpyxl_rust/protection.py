class SheetProtection:
    def __init__(
        self,
        sheet=False,
        password=None,
        format_cells=True,
        format_rows=True,
        format_columns=True,
        insert_columns=True,
        insert_rows=True,
        insert_hyperlinks=True,
        delete_columns=True,
        delete_rows=True,
        select_locked_cells=False,
        select_unlocked_cells=False,
        sort=True,
        autofilter=True,
        pivot_tables=True,
        objects=False,
        scenarios=False,
    ):
        self.sheet = sheet
        self._password = password
        self.format_cells = format_cells
        self.format_rows = format_rows
        self.format_columns = format_columns
        self.insert_columns = insert_columns
        self.insert_rows = insert_rows
        self.insert_hyperlinks = insert_hyperlinks
        self.delete_columns = delete_columns
        self.delete_rows = delete_rows
        self.select_locked_cells = select_locked_cells
        self.select_unlocked_cells = select_unlocked_cells
        self.sort = sort
        self.autofilter = autofilter
        self.pivot_tables = pivot_tables
        self.objects = objects
        self.scenarios = scenarios

    def enable(self):
        self.sheet = True

    def disable(self):
        self.sheet = False

    @property
    def password(self):
        return self._password

    @password.setter
    def password(self, value):
        self._password = value

    def set_password(self, value):
        self._password = value
        self.sheet = True
