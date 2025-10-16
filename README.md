# sap_functions
Library with utility classes and functions to facilitate the development of SAP automations in python.

This module is built on top of [SAP](https://sap.com) Scripting and aims to making the development of automated workflows easier and quicker.

## Implementation example
```python
from sap_functions import SAP

sap = SAP()
sap.select_transaction("COOIS")
```
This scripts:
1. Checks for existant SAP GUI instances.
2. Connects to that instance.
3. Write "COOIS" in the transaction field.

## Classes overview

### SAP
Handle general functions for SAP interactions.

**Attributes**:

`shell_obj`: Object with SAP native functions.  
`session`: the SAP user session.  
`window`: the SAP user session specific window.  
`side_index`: Saves temporary data of private functions.  
`desired_operator`: Saves temporary data of private functions.  
`desired_text`: Saves temporary data of private functions.  
`field_name`: Saves temporary data of private functions.  
`target_index`: Saves temporary data of private functions.  

**Methods**:

`__get_sap_connection`: Get sap session or throw error.  
`__count_and_create_sap_screens`: Create SAP windows until its the number passed.  
`__active_window`: Get active window number.  
`__scroll_through_tabs`: Recursively searches for items in an area.  
`__scroll_through_shell`: Recursively searches for items in an shell based on extension.  
`__scroll_through_table`: Recursively searches for items in an table based on extension.  
`__scroll_through_fields`: Recursively searched for items in a field area.  
`__scroll_through_fields`: Conditionals for different function calling.  
`select_transaction`: Select a transaction.  
`select_main_screen`: Get to main screen.  
`clean_all_fields`: Clean all fields in current screen.  
`run_actual_transaction`: Run actual transaction.  
`insert_variant`: Insert variant in transaction.  
`change_active_tab`: Change tab if there are more than one (popups).  
`write_text_field`: Write text to a specific field.  
`write_text_field_until`: Write text to a specific field "until" section.  
`choose_text_combo`: Interacts with choose text combo of a given field.  
`flag_field`: Flag checkbox fields.  
`flag_field_at_side`: Flag checkbox fields on the side.  
`option_field`: Chooses a radio button within a field.  
`press_button`: Interacts with a button by pressing it.  
`multiple_selection_field`: Interacts with multiple selection of a given field.  
`find_text_field`: Get the value of a field or None if not found.  
`get_text_at_side`: Get the value at the side of a field.  
`multiple_selection_paste_data`: Paste a string into the multiple selection field.  
`navigate_into_menu_header`: Choose a path to go in the menu_header.  
`save_file`: Save a file in specific format and path.  
`get_table`: Get Table instance.  
`get_shell`: Get Shell instance.  
`get_footer_message`: Get footer message.  

### Shell
Represents a SAP shell

**Attributes**:

`shell_obj`: Object with SAP native functions.  
`session`: the SAP user session.  

**Methods**:  
`select_layout`: Select layout within a shell.  
`count_rows`: Get number of rows.   
`get_cell_value`: Get value of given cell.  
`get_shell_content`: Get the shell content in an object.  
`select_all_content`: Select all the shell content.  
`click_cell`: Click on specific cell.  
`press_button`: Press button in shell toolbar.  
`press_nested_button`: Press nested button in shell toolbar.  

### Table
Represents a SAP table

**Attributes**:

`table_obj`: Object with SAP native functions. 

**Methods**:

`get_cell_value`: Get value of given cell.

## Contributing

The sap_functions module is open for contributing, feel free to create an issue or submit your pr. 

## Maintainers

Gabriel Volles Marinho - gabrielvmarinho1711@gmail.com  
Robert Aron Zimmermann - robert.raz@gmail.com