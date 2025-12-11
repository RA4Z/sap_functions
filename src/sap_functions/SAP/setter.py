from src.sap_functions.utils import *
from typing import Union


class SAPSetter:
    def __init__(self, sap):
        self._sap = sap

    def write_text_field(self, field_name: str, desired_text: str, target_index: int = 0,
                         selected_tab: Union[int, str] = 0,
                         skip_error: bool = False) -> None:
        """
        This function will write the desired text in the respective input at the side of the field name
        :param field_name: The text that precedes the desired text field box
        :param desired_text: The text that will overwrite the actual text in the field box
        :param target_index: Target index, determines how many occurrences precede the desired field
        :param selected_tab: Desired Tab, where this field can be found, the SAP default tab is 0
        :param skip_error: Skip this function if occur any error
        """
        try:
            self._sap.window = active_window(self._sap)
            self._sap._field_name = field_name
            self._sap._desired_text = desired_text
            self._sap._target_index = target_index
            if selected_tab != self._sap._selected_tab_id and selected_tab != self._sap._selected_tab_name:
                self._sap.change_active_tab(selected_tab)
            if not scroll_through_fields(self._sap, f"wnd[{self._sap.window}]/usr", 'write_text_field'):
                raise Exception()
        except:
            if not skip_error: raise Exception("Write text field failed.")

    def write_text_field_until(self, field_name: str, desired_text: str, target_index: int = 0,
                               selected_tab: Union[int, str] = 0,
                               skip_error: bool = False) -> None:
        """
        This function will write the desired text in the "until" field in the respective input at the side of the
        field name
        :param field_name: The text that precedes the desired text field box
        :param desired_text: The text that will overwrite the actual text in the field box
        :param target_index: Target index, determines how many occurrences precede the desired field
        :param selected_tab: Desired Tab, where this field can be found, the SAP default tab is 0
        :param skip_error: Skip this function if occur any error
        """
        try:
            self._sap.window = active_window(self._sap)
            self._sap._field_name = field_name
            self._sap._desired_text = desired_text
            self._sap._target_index = target_index
            if selected_tab != self._sap._selected_tab_id and selected_tab != self._sap._selected_tab_name:
                self._sap.change_active_tab(selected_tab)
            if not scroll_through_fields(self._sap, f"wnd[{self._sap.window}]/usr", 'write_text_field_until'):
                raise Exception()
        except:
            if not skip_error: raise Exception("Write text field until failed.")

    def choose_text_combo(self, field_name: str, desired_text: str, target_index: int = 0,
                          selected_tab: Union[int, str] = 0,
                          skip_error: bool = False) -> None:
        """
        This function has the ability to choose a specific text that is found within a combo box component
        :param field_name: The text that precedes the desired combo box
        :param desired_text: The text that will be selected in the combo box
        :param target_index: Target index, determines how many occurrences precede the desired field
        :param selected_tab: Desired Tab, where this field can be found, the SAP default tab is 0
        :param skip_error: Skip this function if occur any error
        """
        try:
            self._sap.window = active_window(self._sap)
            self._sap._field_name = field_name
            self._sap._desired_text = desired_text
            self._sap._target_index = target_index
            if selected_tab != self._sap._selected_tab_id and selected_tab != self._sap._selected_tab_name:
                self._sap.change_active_tab(selected_tab)
            if not scroll_through_fields(self._sap, f"wnd[{self._sap.window}]/usr", 'choose_text_combo'):
                raise Exception()
        except:
            if not skip_error: raise Exception("Choose text combo failed.")

    def flag_field(self, field_name: str, desired_operator: bool, target_index: int = 0,
                   selected_tab: Union[int, str] = 0,
                   skip_error: bool = False) -> None:
        """
        This function can flag and unflag checkboxes based on the field_name, it will flag/unflag the checkbox in the
        respective field_name
        :param field_name: The text with the checkbox you want to flag/unflag
        :param desired_operator: Boolean to say if you want to flag or unflag the checkbox
        :param target_index: Target index, determines how many occurrences precede the desired field
        :param selected_tab: Desired Tab, where this field can be found, the SAP default tab is 0
        :param skip_error: Skip this function if occur any error
        """
        try:
            self._sap.window = active_window(self._sap)
            self._sap._field_name = field_name
            self._sap._desired_operator = desired_operator
            self._sap._target_index = target_index
            if selected_tab != self._sap._selected_tab_id and selected_tab != self._sap._selected_tab_name:
                self._sap.change_active_tab(selected_tab)
            if not scroll_through_fields(self._sap, f"wnd[{self._sap.window}]/usr", 'flag_field'):
                raise Exception()
        except:
            if not skip_error: raise Exception("Flag field failed.")

    def flag_field_at_side(self, field_name: str, desired_operator: bool, side_index: int = 0, target_index: int = 0,
                           selected_tab: Union[int, str] = 0, skip_error: bool = False) -> None:
        """
        This function can flag and unflag checkboxes based on the field_name, it will flag/unflag the checkbox at the
        side of the respective field_name
        :param field_name: The text at the side of the checkbox you want to flag/unflag
        :param desired_operator: Boolean to say if you want to flag or unflag the checkbox
        :param side_index: Number of components at the side of the respective field_name, with positive numbers the code will go through components at right, if negative it will go through components at left
        :param target_index: Target index, determines how many occurrences precede the desired field
        :param selected_tab: Desired Tab, where this field can be found, the SAP default tab is 0
        :param skip_error: Skip this function if occur any error
        """
        try:
            self._sap.window = active_window(self._sap)
            self._sap._field_name = field_name
            self._sap._desired_operator = desired_operator
            self._sap._target_index = target_index
            self._sap._side_index = side_index
            if selected_tab != self._sap._selected_tab_id and selected_tab != self._sap._selected_tab_name:
                self._sap.change_active_tab(selected_tab)
            if not scroll_through_fields(self._sap, f"wnd[{self._sap.window}]/usr", 'flag_field_at_side'):
                raise Exception()
        except:
            if not skip_error: raise Exception("Flag field at side failed.")

    def option_field(self, field_name: str, target_index: int = 0, selected_tab: Union[int, str] = 0,
                     skip_error: bool = False) -> None:
        """
        This function will select an option field
        :param field_name: The text with the option field you want to select
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
            if not scroll_through_fields(self._sap, f"wnd[{self._sap.window}]/usr", 'option_field'):
                raise Exception()
        except:
            if not skip_error: raise Exception("Option field failed.")

    def option_field_at_side(self, field_name: str, side_index: int = 0, target_index: int = 0,
                             selected_tab: Union[int, str] = 0, skip_error: bool = False) -> None:
        """
        This function will select an option field
        :param field_name: The text with the option field you want to select
        :param side_index: Number of components at the side of the respective field_name, with positive numbers the code will go through components at right, if negative it will go through components at left
        :param target_index: Target index, determines how many occurrences precede the desired field
        :param selected_tab: Desired Tab, where this field can be found, the SAP default tab is 0
        :param skip_error: Skip this function if occur any error
        """
        try:
            self._sap.window = active_window(self._sap)
            self._sap._field_name = field_name
            self._sap._target_index = target_index
            self._sap._side_index = side_index
            if selected_tab != self._sap._selected_tab_id and selected_tab != self._sap._selected_tab_name:
                self._sap.change_active_tab(selected_tab)
            if not scroll_through_fields(self._sap, f"wnd[{self._sap.window}]/usr", 'option_field_at_side'):
                raise Exception()
        except:
            if not skip_error: raise Exception("Option field failed.")
