import os
import uuid
from src.sap_functions.utils import *
from src.sap_functions.base_sap_connection import BaseSapConnection
from src.sap_functions.SAP.getter import SAPGetter
from src.sap_functions.SAP.action import SAPAction
from src.sap_functions.SAP.setter import SAPSetter
from typing import Union


# SAP Scripting Documentation:
# https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/a2e9357389334dc89eecc1fb13999ee3.html

# module SAP Functions, development started in 2024/03/01
class SAP(BaseSapConnection):

    def __init__(self, window: int = 0) -> None:
        BaseSapConnection.__init__(self, window)
        self._component_target_index = None
        self._desired_operator = None
        self._selected_tab_id = None
        self._desired_text = None
        self._target_index = None
        self._target_tab = None
        self._side_index = None
        self._field_name = None
        self._found_text = None
        self._selected_tab_name = ''
        self._get = SAPGetter(self)
        self._set = SAPSetter(self)
        self._action = SAPAction(self)

        if self.session.info.transaction == 'S000':
            self.select_main_screen()

    @property
    def get(self) -> SAPGetter:
        return self._get

    @property
    def set(self) -> SAPSetter:
        return self._set

    @property
    def action(self) -> SAPAction:
        return self._action

    def select_transaction(self, transaction: str) -> None:
        """
        Navigate to a transaction in SAP GUI
        :param transaction: The name of the desired transaction
        """
        try:
            transaction_upper = transaction.upper()
            self.session.startTransaction(transaction_upper)
            if self.session.activeWindow.name == 'wnd[1]' and 'CN' in transaction_upper:
                self.session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            if not self.session.info.transaction == transaction_upper:
                raise Exception()
        except:
            raise Exception("Select transaction failed.\n" + self.get.footer_message())

    def select_main_screen(self, skip_error: bool = False) -> None:
        """
        Navigate to the SAP main Screen
        :param skip_error: Skip this function if occur any error
        """
        try:
            if not self.session.info.transaction == "SESSION_MANAGER":
                self.session.startTransaction('SESSION_MANAGER')
                if self.session.activeWindow.name == "wnd[1]":
                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            if not skip_error: raise Exception("Select main screen failed.")

    def run_actual_transaction(self, skip_error: bool = False) -> None:
        """
        Run the active transaction, this function will try to press Enter, and after that will try to press F8
        :param skip_error: Skip this function if occur any error
        """
        try:
            self.window = active_window(self)
            screen_title = self.session.activeWindow.text
            self.session.findById(f'wnd[{self.window}]').sendVKey(0)
            if screen_title == self.session.activeWindow.text:
                self.session.findById(f'wnd[{self.window}]').sendVKey(8)
        except:
            if not skip_error:
                raise Exception("Run actual transaction failed.")

    def change_active_tab(self, selected_tab: Union[int, str], skip_error: bool = False) -> None:
        """
        This function will try to select the transaction tab using the number "selected_tab"
        :param selected_tab: Tab desired number, the SAP default tab is 0
        :param skip_error: Skip this function if occur any error
        """
        try:
            self.window = active_window(self)
            if type(selected_tab).__name__ == 'int':
                area = scroll_through_tabs_by_id(self, self.session.findById(f"wnd[{self.window}]/usr"),
                                                 f"wnd[{self.window}]/usr", selected_tab)
                try:
                    area.Select()
                except:
                    pass
            else:
                self._target_tab = selected_tab
                scroll_through_fields(self, f"wnd[{self.window}]/usr", 'select_tab_by_name')

        except:
            if not skip_error: raise Exception("Change active tab failed.")

    def find_text_field(self, field_name: str, selected_tab: Union[int, str] = 0) -> bool:
        """
        Verify if a text exists in the SAP screen
        :param field_name: The text that you want to search
        :param selected_tab: Desired Tab, where this field can be found, the SAP default tab is 0
        :return: A boolean, True if the text was found and False if it was not found
        """
        self.window = active_window(self)
        self._field_name = field_name
        if selected_tab != self._selected_tab_id and selected_tab != self._selected_tab_name:
            self.change_active_tab(selected_tab)
        return scroll_through_fields(self, f"wnd[{self.window}]/usr", 'find_text_field')

    def navigate_into_menu_header(self, *nested_path: str) -> None:
        """
        This function needs to receive several strings that have the texts that appear written in the header destination
        that you want to press, it must be written in the order that it appears in the SAP header
        :param nested_path: The nested path that you want to navigate into the header
        """
        id_path = 'wnd[0]/mbar'
        for active_path in nested_path:
            children = self.session.findById(id_path).children
            for i in range(children.count):
                Obj = children[i]
                if active_path in Obj.text:
                    menu_address = Obj.id.split("/")[-1]
                    id_path += f'/{menu_address}'
                    break
        self.session.findById(id_path).Select()
