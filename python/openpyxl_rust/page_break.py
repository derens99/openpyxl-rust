class Break:
    def __init__(self, id=0):
        self.id = id


class BreakList:
    def __init__(self):
        self._breaks = []

    def append(self, brk):
        if isinstance(brk, Break):
            self._breaks.append(brk)
        else:
            self._breaks.append(Break(id=brk))

    def __iter__(self):
        return iter(self._breaks)

    def __len__(self):
        return len(self._breaks)

    def __bool__(self):
        return bool(self._breaks)
