class Protection:
    def __init__(self, locked=True, hidden=False):
        self.locked = locked
        self.hidden = hidden

    def __eq__(self, other):
        if not isinstance(other, Protection):
            return False
        return self.locked == other.locked and self.hidden == other.hidden

    def __repr__(self):
        return f"Protection(locked={self.locked}, hidden={self.hidden})"
