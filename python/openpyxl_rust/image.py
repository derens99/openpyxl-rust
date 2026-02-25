class Image:
    def __init__(self, img):
        """img can be a file path (str) or bytes."""
        if isinstance(img, str):
            with open(img, 'rb') as f:
                self._data = f.read()
        elif isinstance(img, bytes):
            self._data = img
        elif isinstance(img, bytearray):
            self._data = bytes(img)
        else:
            raise TypeError(f"Expected str path or bytes, got {type(img)}")
        self.anchor = None
        self.width = None
        self.height = None
