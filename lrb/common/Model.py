class Excel:
    path = ''
    xlsx = None
    active = None
    max_row = 0
    max_column = 0
    lines: list = None

    def __init__(self, path, xlsx, active, max_row, max_column, lines):
        self.path = path
        self.xlsx = xlsx
        self.active = active
        self.max_row = max_row
        self.max_column = max_column
        self.lines = lines
