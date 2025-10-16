class Shell:
    def __init__(self, shell_obj, session):
        self.shell_obj = shell_obj
        self.session = session

    def select_layout(self, layout):
        try:
            self.shell_obj.selectColumn("VARIANT")
            self.shell_obj.contextMenu()
            self.shell_obj.selectContextMenuItem("&FILTER")
            self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = layout
            self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
            self.session.findById(
                "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
            self.session.findById(
                "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        except:
            raise Exception("Select layout failed.")

    def count_rows(self):
        try:
            rows = self.shell_obj.RowCount
            if rows > 0:
                visiblerow = self.shell_obj.VisibleRowCount
                visiblerow0 = self.shell_obj.VisibleRowCount
                npagedown = rows // visiblerow0
                if npagedown > 1:
                    for j in range(1, npagedown + 1):
                        try:
                            self.shell_obj.firstVisibleRow = visiblerow - 1
                            visiblerow += visiblerow0
                        except:
                            break
                self.shell_obj.firstVisibleRow = 0
            return rows
        except:
            raise Exception("Count rows failed.")

    def get_cell_value(self, index, column_id):
        try:
            return self.shell_obj.getCellValue(index, column_id)
        except:
            raise Exception("Get cell value failed.")

    def get_shell_content(self):
        try:
            grid_column = self.shell_obj.ColumnOrder
            rows = self.count_rows()
            cols = self.shell_obj.ColumnCount

            data = [[self.shell_obj.getCellValue(i, grid_column(c)) for c in range(cols)] for i in range(-1, rows)]
            return data
        except:
            raise Exception("Get all Shell Content Failed.")

    def select_all_content(self):
        try:
            self.shell_obj.selectAll()
        except:
            raise Exception("Select All Content Failed.")

    def select_column(self, column_id):
        try:
            self.shell_obj.selectColumn(column_id)
        except:
            raise Exception("Select Column Failed.")

    def click_cell(self, index, column_id):
        try:
            self.shell_obj.SetCurrentCell(index, column_id)
            self.shell_obj.doubleClickCurrentCell()
        except:
            raise Exception("Click Cell Failed.")

    def press_button(self, field_name: str):
        for i in range(100):
            try:
                button_id = self.shell_obj.GetToolbarButtonId(i)
                button_tooltip = self.shell_obj.GetToolbarButtonTooltip(i)
                if field_name == button_tooltip:
                    self.shell_obj.pressToolbarContextButton(button_id)
                    return True

            except Exception as err:
                raise Exception(f"The error {err} happenned trying to press the button {field_name}.")

        return False

    def press_nested_button(self, *nested_fields):
        for i in range(100):
            try:
                button_id = self.shell_obj.GetToolbarButtonId(i)
                button_tooltip = self.shell_obj.GetToolbarButtonTooltip(i)
                if nested_fields[0] == button_tooltip:
                    self.shell_obj.pressToolbarContextButton(button_id)
                    self.shell_obj.SelectContextMenuItemByText(nested_fields[1])
                    return True

            except Exception as err:
                raise Exception(f"The error {err} happenned trying to press the nest {nested_fields}.")

        return False
