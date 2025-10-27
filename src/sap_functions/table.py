from .utils import *
import copy


# https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/ce1d9e64355d49568e5def5271aea2db.html?locale=en-US
class Table:
    def __init__(self, table, session, target_index: int):
        self._component_target_index = target_index
        self.__target_index = target_index
        self.table_obj = table
        self.session = session
        self.window = active_window(self)

    def __return_table(self):
        self._component_target_index = copy.copy(self.__target_index)
        return scroll_through_table(self, f'wnd[{self.window}]/usr')

    def get_cell_value(self, row: int, column: int, skip_error: bool = False) -> str:
        """
        Return the content of a SAP Table cell, using the relative visible table row. The desired cell needs to be
        visible for this function be able to work
        :param row: Table relative row index
        :param column: Table column index
        :param skip_error: Skip this function if occur any error
        :return: A String with the desired cell text
        """
        try:
            return self.table_obj.getCell(row, column).text
        except:
            if not skip_error:
                raise Exception("Get cell value failed.")

    def count_visible_rows(self, skip_error: bool = False) -> int:
        """
        Count all the visible rows from a SAP Table
        :param skip_error: Skip this function if occur any error
        :return: An Integer with the number of visible rows in the active SAP Table
        """
        try:
            return self.table_obj.visibleRowCount
        except:
            if not skip_error:
                raise Exception("Get cell value failed.")

    def write_cell_value(self, row: int, column: int, desired_text: str, skip_error: bool = False) -> None:
        """
        Write any value in a SAP Table cell, using the relative visible table row. The desired cell needs to be
        visible for this function be able to work
        :param row: Table relative row index
        :param column: Table column index
        :param desired_text: The text that will overwrite the cell in the SAP Table
        :param skip_error: Skip this function if occur any error
        """
        try:
            self.table_obj.getCell(row, column).text = desired_text
        except:
            if not skip_error:
                raise Exception("Write cell value failed.")

    def select_entire_row(self, absolute_row: int, skip_error: bool = False) -> None:
        """
        Select the entire row from a SAP Table, it uses the absolute table row. The desired cell needs to be
        visible for this function be able to work
        :param absolute_row: Table absolute row index
        :param skip_error: Skip this function if occur any error
        """
        try:
            self.table_obj.GetAbsoluteRow(absolute_row).selected = True
        except:
            if not skip_error:
                raise Exception("Select Entire Row Failed.")

    def unselect_entire_row(self, absolute_row: int, skip_error: bool = False) -> None:
        """
        Unselect the entire row from a SAP Table, it uses the absolute table row. The desired cell needs to be
        visible for this function be able to work
        :param absolute_row: Table absolute row index
        :param skip_error: Skip this function if occur any error
        """
        try:
            self.table_obj.GetAbsoluteRow(absolute_row).selected = False
        except:
            if not skip_error:
                raise Exception("Unselect Entire Row Failed.")

    def flag_cell(self, row: int, column: int, desired_operator: bool, skip_error: bool = False) -> None:
        """
        Flags a checkbox in a SAP Table cell, using the relative visible table row. The desired cell needs to be
        visible for this function be able to work
        :param row: Table relative row index
        :param column: Table column index
        :param skip_error: Skip this function if occur any error
        :param desired_operator: Boolean with the desired operator in the SAP Table cell's checkbox
        """
        try:
            self.table_obj.getCell(row, column).Selected = desired_operator
        except:
            if not skip_error:
                raise Exception("Flag Cell Failed.")

    def click_cell(self, row: int, column: int, skip_error: bool = False) -> None:
        """
        Focus in a cell and double-click in it, using the relative visible table row. The desired cell needs to be
        visible for this function be able to work
        :param row: Table relative row index
        :param column: Table column index
        :param skip_error: Skip this function if occur any error
        """
        try:
            self.table_obj.getCell(row, column).SetFocus()
            self.session.findById(f"wnd[{self.window}]").sendVKey(2)
        except:
            if not skip_error:
                raise Exception("Click Cell Failed.")

    def get_table_content(self, skip_error: bool = False) -> dict:
        """
        Store all the content from a SAP Table, the data will be stored and returned in a dictionary with 'header' and
        'content' items
        :param skip_error: Skip this function if occur any error
        :return: A dictionary with 'header' and 'content' items
        """
        try:
            self.__return_table().VerticalScrollbar.Position = 0
            obj_now = self.__return_table()
            added_rows = []

            header = []
            content = []

            columns = obj_now.columns.count
            visible_rows = obj_now.visibleRowCount
            rows = obj_now.rowCount / visible_rows

            iteration_plus = 0
            if obj_now.rowCount > visible_rows:
                iteration_plus = 1

            absolute_row = 0

            for c in range(columns):
                col_name = obj_now.columns.elementAt(c).title
                header.append(col_name)

            for i in range(int(rows) + iteration_plus):
                for visible_row in range(visible_rows):
                    active_row = []
                    for c in range(columns):
                        try:
                            active_row.append(obj_now.getCell(visible_row, c).text)
                        except:
                            active_row.append(None)

                    absolute_row += 1

                    if not all(value is None for value in active_row) and absolute_row not in added_rows:
                        added_rows.append(absolute_row)
                        content.append(active_row)

                obj_now.VerticalScrollbar.Position = (visible_row + 1) * i
                obj_now = self.__return_table()
            return {'header': header, 'content': content}

        except:
            if not skip_error:
                raise Exception("Get table content failed.")
