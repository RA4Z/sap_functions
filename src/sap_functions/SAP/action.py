from src.sap_functions.utils import *
from typing import Union
import uuid
import os


class SAPAction:
    def __init__(self, sap):
        self._sap = sap

    def clean_all_fields(self, selected_tab: Union[int, str] = 0, skip_error=False) -> None:
        """
        Clean all the input fields in the actual screen
        :param selected_tab: Transaction desired tab
        :param skip_error: Skip this function if occur any error
        """
        try:
            self._sap.window = active_window(self._sap)
            if type(selected_tab).__name__ == 'int':
                area = scroll_through_tabs_by_id(self._sap, self._sap.session.findById(f"wnd[{self._sap.window}]/usr"),
                                                 f"wnd[{self._sap.window}]/usr", selected_tab)
            else:
                area = scroll_through_tabs_by_name(self._sap, self._sap.session.findById(f"wnd[{self._sap.window}]/usr"),
                                                   f"wnd[{self._sap.window}]/usr", selected_tab)

            children = area.children
            for child in children:
                if child.type == "GuiCTextField":
                    try:
                        child.Text = ""
                    except:
                        pass
        except:
            if not skip_error: raise Exception("Clean all fields failed.")

    def save_file(self, file_name: str, path: str, option: int = 0, type_of_file: str = 'txt',
                  skip_error: bool = False) -> None:
        """
        This function will easily navigate into SAP menu header to save the current transaction data, commonly used to
        extract data Labels
        :param file_name: The name of the file that you want to save
        :param path: The path that you want to save the file
        :param option: The txt option of save format 0=>Unconverted,1=>Text with Tabs,2=>Rich text format
        :param type_of_file: The extension that you want for the file
        :param skip_error: Skip this function if occur any error
        """
        try:
            if 'xls' in type_of_file:
                self._sap.session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").Select()
                self._sap.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            else:
                self._sap.session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select()
                self._sap.session.findById(
                    f"wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[{option},0]").Select()
                self._sap.session.findById("wnd[1]/tbar[0]/btn[0]").press()

            self._sap.session.findById("wnd[1]/usr/ctxtDY_PATH").Text = path
            self._sap.session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = f'{file_name}.{type_of_file}'
            self._sap.session.findById("wnd[1]/tbar[0]/btn[11]").press()
        except:
            if not skip_error: raise Exception("Save file failed.")

    def insert_variant(self, variant_name: str, skip_error: bool = False) -> None:
        """
        This function will try to press the "Get Variant" button in the transaction, after that it will overwrite the
        "Created By" field with an empty string, and fill the "Variant" field with the variant_name param, THIS FUNCTION
        DOESN'T WORK IN EVERY TRANSACTION
        :param variant_name: The transaction variant name
        :param skip_error: Skip this function if occur any error
        """
        try:
            self._sap.session.findById("wnd[0]/tbar[1]/btn[17]").press()
            if self._sap.session.activeWindow.name == 'wnd[1]':
                self._sap.session.findById("wnd[1]/usr/txtV-LOW").Text = variant_name
                self._sap.session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
                self._sap.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                if self._sap.session.activeWindow.name == 'wnd[1]':
                    raise Exception()
        except:
            if not skip_error: raise Exception("Insert variant failed.")

    def multiple_selection_field(self, field_name: str, target_index: int = 0, selected_tab: Union[int, str] = 0,
                                 skip_error: bool = False) -> None:
        """
        This function will press the "Multiple Selection" button in the respective field
        :param field_name: The text that precedes the desired text field box
        :param target_index: Target index, determines how many occurrences precede the desired field
        :param selected_tab: Desired Tab, where this field can be found, the SAP default tab is 0
        :param skip_error: Skip this function if occur any error
        """
        try:
            self._sap.window = active_window(self._sap)
            self._sap._field_name = field_name
            self._sap._target_index = target_index
            if selected_tab != self._sap._selected_tab_id and selected_tab != self._sap._selected_tab_name:
                self._sap.change_active_tab(selected_tab)
            if not scroll_through_fields(self._sap, f"wnd[{self._sap.window}]/usr", 'multiple_selection_field'):
                raise Exception()
        except:
            if not skip_error: raise Exception("Multiple selection field failed.")

    def multiple_selection_paste_data(self, data: list[str], delete_values: bool = False,
                                      skip_error: bool = False) -> None:
        """
        With the Multiple Selection window open, it's possible to execute this function to easily paste all the data
        from a list
        :param data: An array with the data that you want to insert in the multiple selection
        :param delete_values: Boolean to determine if you want to insert the list to delete it from the final result
        :param skip_error: Skip this function if occur any error
        """
        uid = str(uuid.uuid4())
        self._sap.window = active_window(self._sap)

        if delete_values:
            self._sap.change_active_tab(2)

        try:
            with open(f'C:/Temp/{uid}.txt', 'w') as file:
                file.write('\n'.join(data))
            self._sap.session.findById(f"wnd[{self._sap.window}]/tbar[0]/btn[23]").press()
            self._sap.session.findById(f"wnd[{self._sap.window + 1}]/usr/ctxtDY_PATH").text = 'C:/Temp'
            self._sap.session.findById(f"wnd[{self._sap.window + 1}]/usr/ctxtDY_FILENAME").text = f"{uid}.txt"
            self._sap.session.findById(f"wnd[{self._sap.window + 1}]/tbar[0]/btn[0]").press()
            self._sap.session.findById(f"wnd[{self._sap.window}]/tbar[0]/btn[8]").press()
            if os.path.exists(f'C:/Temp/{uid}.txt'):
                os.remove(f'C:/Temp/{uid}.txt')
        except:
            if not skip_error: raise Exception("Multiple selection paste data failed.")

    def press_button(self, field_name: str, target_index: int = 0, selected_tab: Union[int, str] = 0,
                     skip_error: bool = False) -> None:
        """
        Press any button in the SAP screens, except in shells and tables components
        :param field_name: The button that you want to press, this text need to be inside the button or in the tooltip of the button
        :param target_index: Target index, determines how many occurrences precede the desired field
        :param selected_tab: Desired Tab, where this field can be found, the SAP default tab is 0
        :param skip_error: Skip this function if occur any error
        """
        try:
            self._sap.window = active_window(self._sap)
            self._sap._field_name = field_name
            self._sap._target_index = target_index
            if selected_tab != self._sap._selected_tab_id and selected_tab != self._sap._selected_tab_name:
                self._sap.change_active_tab(selected_tab)
            if not scroll_through_fields(self._sap, f"wnd[{self._sap.window}]", 'press_button'):
                raise Exception()
        except:
            if not skip_error: raise Exception("Press button failed.")

    def set_focus(self, field_name, side_index: int = 0, target_index: int = 0, selected_tab: Union[int, str] = 0):
        """
        This function will select/focus in the field with the text received as a parameter
        :param field_name: The text that you want to focus
        :param side_index: Number of components at the side of the respective field_name, with positive numbers the code will go through components at right, if negative it will go through components at left
        :param target_index: Target index, determines how many occurrences precede the desired field
        :param selected_tab: Desired Tab, where this field can be found, the SAP default tab is 0
        """
        self._sap.window = active_window(self._sap)
        self._sap._field_name = field_name
        self._sap._target_index = target_index
        self._sap._side_index = side_index
        if selected_tab != self._sap._selected_tab_id and selected_tab != self._sap._selected_tab_name:
            self._sap.change_active_tab(selected_tab)
        scroll_through_fields(self._sap, f"wnd[{self._sap.window}]", 'set_focus')

    def open_focused_field_modal(self):
        """
        This function will open the modal of a focused field
        """
        self._sap.session.findById(f'wnd[{self._sap.window}]').sendVKey(4)
