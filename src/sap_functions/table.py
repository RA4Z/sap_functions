class Table:
   def __init__(self, table):
      self.table = table

   def get_cell_value(self, row, column):
      try:
         return self.table.getCell(row, column).Text
      except:
          raise Exception("Get cell value failed.")