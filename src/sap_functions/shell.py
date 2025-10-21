import win32com


class Shell:
    def __init__(self, shell_obj: win32com.client.CDispatch, session: win32com.client.CDispatch):
        self.shell_obj = shell_obj
        self.session = session

    def select_layout(self, layout: str) -> None:
        """
        This function will select the desired Shell Layout when a SAP select Layout Pop up is open
        :param layout: The desired Layout name
        """
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

    def count_rows(self) -> int:
        """
        This function will count all the rows in the current Shell
        :return: A integer with the total number of rows in the current Shell
        """
        try:
            rows = self.shell_obj.RowCount
            if rows > 0:
                visible_row = self.shell_obj.VisibleRowCount
                visible_row0 = self.shell_obj.VisibleRowCount
                n_page_down = rows // visible_row0
                if n_page_down > 1:
                    for j in range(1, n_page_down + 1):
                        try:
                            self.shell_obj.firstVisibleRow = visible_row - 1
                            visible_row += visible_row0
                        except:
                            break
                self.shell_obj.firstVisibleRow = 0
            return rows
        except:
            raise Exception("Count rows failed.")

    def get_cell_value(self, index: int, column_id: str) -> str:
        """
        Get the value of a specific Shell cell
        :param index: Row number of the desired cell
        :param column_id: Shell column "Field Name" found in the respective column Technical Information tab
        :return: The value of the cell
        """
        try:
            return self.shell_obj.getCellValue(index, column_id)
        except:
            raise Exception("Get cell value failed.")

    def get_shell_content(self) -> dict:
        """
        Store all the content from a SAP Shell, the data will be stored and returned in a dictionary with 'header' and
        'content' items
        :return: A dictionary with 'header' and 'content' items
        """
        try:
            grid_column = self.shell_obj.ColumnOrder
            rows = self.count_rows()
            cols = self.shell_obj.ColumnCount

            header = [self.shell_obj.getCellValue(i, grid_column(c)) for c in range(cols) for i in range(-1, 0)]
            data = [[self.shell_obj.getCellValue(i, grid_column(c)) for c in range(cols)] for i in range(0, rows)]
            return {'header': header, 'content': data}

        except:
            raise Exception("Get all Shell Content Failed.")

    def select_all_content(self) -> None:
        """
        Select all the table, using the SAP native function to select all items
        """
        try:
            self.shell_obj.selectAll()
        except:
            raise Exception("Select All Content Failed.")

    def select_column(self, column_id: int) -> None:
        """
        Select a specific column
        :param column_id: Shell column "Field Name" found in the respective column Technical Information tab
        """
        try:
            self.shell_obj.selectColumn(column_id)
        except:
            raise Exception("Select Column Failed.")

    def click_cell(self, index: int, column_id: str) -> None:
        """
        This function will select and double-click in a SAP Shell cell
        :param index: Row number of the desired cell
        :param column_id: Shell column "Field Name" found in the respective column Technical Information tab
        """
        try:
            self.shell_obj.SetCurrentCell(index, column_id)
            self.shell_obj.doubleClickCurrentCell()
        except:
            raise Exception("Click Cell Failed.")

    def press_button(self, field_name: str, skip_error: bool = False) -> None:
        """
        This function will press any button in the SAP Shell component
        :param field_name: The button that you want to press, this text need to be inside the button or in the tooltip of the button
        :param skip_error: Skip this function if occur any error
        """
        try:
            found = False
            for i in range(100):
                button_id = self.shell_obj.GetToolbarButtonId(i)
                button_tooltip = self.shell_obj.GetToolbarButtonTooltip(i)
                if field_name == button_tooltip:
                    self.shell_obj.pressToolbarContextButton(button_id)
                    found = True
            if not found and not skip_error:
                raise Exception()
        except:
            raise Exception("Press button failed")

    def press_nested_button(self, *nested_fields: str, skip_error: bool = False) -> None:
        """
        This function needs to receive several strings that have the texts that appear written in the button destination
        that you want to press, it must be written in the order that it appears to reach the final button
        :param nested_fields: The nested path that you want to navigate to press the button
        :param skip_error: Skip this function if occur any error
        """
        try:
            found = False
            for i in range(100):
                button_id = self.shell_obj.GetToolbarButtonId(i)
                button_tooltip = self.shell_obj.GetToolbarButtonTooltip(i)
                if nested_fields[0] == button_tooltip:
                    self.shell_obj.pressToolbarContextButton(button_id)
                    self.shell_obj.SelectContextMenuItemByText(nested_fields[1])
                    found = True
            if not found and not skip_error:
                raise Exception()
        except:
            raise Exception("Press nested button failed")
