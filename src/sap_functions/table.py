class Table:
    def __init__(self, table):
        self.table_obj = table

    def get_cell_value(self, row: int, column: str) -> str:
        try:
            return self.table_obj.getCell(row, column).Text
        except:
            raise Exception("Get cell value failed.")
