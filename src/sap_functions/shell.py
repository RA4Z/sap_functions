
class Shell:
   def __init__(self, shell):
      self.shell = shell

   def select_layout(self, layout, skip_error=False):
      try:
         self.shell.selectColumn("VARIANT")
         self.shell.contextMenu()
         self.shell.selectContextMenuItem("&FILTER")
         self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = layout
         self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
         self.session.findById(
               "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
         self.session.findById(
               "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
      except:
         if not skip_error: raise Exception("Select layout failed.")

   def count_rows(self):
      try:
         rows = self.shell.RowCount
         if rows > 0:
               visiblerow = self.shell.VisibleRowCount
               visiblerow0 = self.shell.VisibleRowCount
               npagedown = rows // visiblerow0
               if npagedown > 1:
                  for j in range(1, npagedown + 1):
                     try:
                           self.shell.firstVisibleRow = visiblerow - 1
                           visiblerow += visiblerow0
                     except:
                           break
               self.shell.firstVisibleRow = 0
         return rows
      except:
         raise Exception("Count rows failed.")
      
   def get_cell_value(self, index, id):
      try:
         return self.shell.getCellValue(index, id)
      except:
          raise Exception("Get cell value failed.")
      
   def view_in_list_form(self): 
      try:
         self.shell.pressToolbarContextButton("&MB_VIEW")
         self.shell.SelectContextMenuItem("&PRINT_BACK_PREVIEW")
      except:
            raise Exception("View in list form failed.")