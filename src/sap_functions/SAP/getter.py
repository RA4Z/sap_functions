from src.sap_functions.tree import Tree
from src.sap_functions.node import Node
from src.sap_functions.grid import Grid
from src.sap_functions.table import Table
from src.sap_functions.label import Label
from src.sap_functions.utils import *


class SAPGetter:
    def __init__(self, sap):
        self._sap = sap

    def label(self) -> Label:
        """
        Get the SAP Label object from the current SAP Label Window
        :return: A SAP Label object, that can be used to extract data from Label components in SAP
        """
        try:
            self._sap.window = active_window(self._sap)
            label = Label(self._sap.session, self._sap.window)
            return label
        except:
            raise Exception("Get label failed.")

    def table(self, target_index: int = 0) -> Table:
        """
        Get the SAP Table object from the current SAP Table Window
        :param target_index: Target index, determines how many occurrences precede the desired component
        :return: A SAP Table object, that can be used to extract data from Table components in SAP
        """
        try:
            self._sap.window = active_window(self._sap)
            self._sap._component_target_index = target_index
            table_obj = scroll_through_table(self._sap, f'wnd[{self._sap.window}]/usr')
            if not table_obj:
                raise Exception()
            table = Table(table_obj, self._sap.session, target_index)
            return table
        except:
            raise Exception("Get table failed.")

    def tree(self) -> Tree:
        """
        Get the SAP Tree object from the current SAP Tree Window
        :return: A SAP Tree object, that can be used to extract data from Tree tables in SAP
        """
        try:
            self._sap.window = active_window(self._sap)
            tree_obj = scroll_through_tree(self._sap, f'wnd[{self._sap.window}]')

            if not tree_obj:
                raise Exception("Tree Object not found")

            tree = Tree(tree_obj)
            return tree
        except:
            raise Exception("Get Tree failed.")

    def grid(self, target_index: int = 0) -> Grid:
        """
        Get the SAP Grid object from the current SAP Grid Window
        :param target_index: Target index, determines how many occurrences precede the desired component
        :return: A SAP Grid object, that can be used to extract data from Grid tables in SAP
        """
        try:
            self._sap.window = active_window(self._sap)
            self._sap._component_target_index = target_index
            grid_obj = scroll_through_grid(self._sap, f'wnd[{self._sap.window}]')

            if not grid_obj:
                raise Exception()

            grid = Grid(grid_obj, self._sap.session)
            return grid
        except:
            raise Exception("Get grid failed.")

    def node(self, target_index: int = 0) -> Node:
        """
        Get the SAP Node object from the current SAP Node Window
        :param target_index: Target index, determines how many occurrences precede the desired component
        :return: A SAP Node object, that can be used to extract data from Node components in SAP
        """
        try:
            self._sap.window = active_window(self._sap)
            self._sap._component_target_index = target_index
            node_obj = scroll_through_node(self._sap, f'wnd[{self._sap.window}]')

            if not node_obj:
                raise Exception()

            node = Node(node_obj)
            return node

        except:
            raise Exception("Get node failed.")

    def footer_message(self) -> str:
        """
        Get the message text that is in the SAP Footer
        :return: A String with the footer message
        """
        try:
            return self._sap.session.findById("wnd[0]/sbar").text
        except:
            raise Exception("Get footer message failed.")

    def text_at_side(self, field_name, side_index: int, target_index: int = 0,
                     selected_tab: Union[int, str] = 0) -> str:
        """
        This function will return the text next to the text received as a parameter
        :param field_name: The text that you want to search
        :param side_index: Number of components at the side of the respective field_name, with positive numbers the code will go through components at right, if negative it will go through components at left
        :param target_index: Target index, determines how many occurrences precede the desired field
        :param selected_tab: Desired Tab, where this field can be found, the SAP default tab is 0
        :return: A string with the text at the side of the searched text
        """
        self._sap.window = active_window(self._sap)
        self._sap._field_name = field_name
        self._sap._target_index = target_index
        self._sap._side_index = side_index
        if selected_tab != self._sap._selected_tab_id and selected_tab != self._sap._selected_tab_name:
            self._sap.change_active_tab(selected_tab)
        if scroll_through_fields(self._sap, f"wnd[{self._sap.window}]", 'get_text_at_side'):
            return self._sap._found_text
