import win32com.client
from typing import Union


# https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/8f08be87b0194d9882d0382eae798617.html?locale=en-US
class Tree:
    def __init__(self, tree_obj: win32com.client.CDispatch):
        self.tree_obj = tree_obj

    def get_tree_columns(self, *column_text: str) -> Union[dict, list]:
        """
        Return each column content
        :param column_text: Tree list of columns "Field Text"
        :return: A dictionary/list with the desired content, when more than one column is desired, a dictionary with 'header' and 'content' items will be returned
        """
        try:
            header = []
            content = []

            obj_key_values = self.tree_obj.GetAllNodeKeys()
            all_column_names = self.tree_obj.GetColumnNames()
            columns = {}

            for col in all_column_names:
                colName = self.tree_obj.GetColumnTitleFromName(col)
                columns[colName] = col
                if colName in column_text:
                    header.append(colName)

            for i in range(1, len(obj_key_values)):
                active_row = []
                for col in column_text:
                    item = str(self.tree_obj.getItemText(obj_key_values(i), columns[col]))
                    if len(column_text) > 1:
                        active_row.append(item)
                    else:
                        content.append(item)

                if len(column_text) > 1:
                    content.append(active_row)

            if len(column_text) > 1:
                return {'header': header, 'content': content}
            else:
                return content

        except:
            raise Exception("Get Tree Columns Failed.")

    def get_tree_content(self, skip_error: bool = False) -> dict:
        """
        Store all the content from a SAP Tree, the data will be stored and returned in a dictionary with 'header' and
        'content' items
        :param skip_error: Skip this function if occur any error
        :return: A dictionary with 'header' and 'content' items
        """
        try:
            header = []
            content = []

            obj_key_values = self.tree_obj.GetAllNodeKeys()
            all_column_names = self.tree_obj.GetColumnNames()
            columns = {}

            for col in all_column_names:
                colName = self.tree_obj.GetColumnTitleFromName(col)
                columns[colName] = col
                header.append(colName)

            for i in range(1, len(obj_key_values)):
                active_row = []
                for col in columns:
                    item = str(self.tree_obj.getItemText(obj_key_values(i), columns[col]))
                    active_row.append(item)

                content.append(active_row)
            return {'header': header, 'content': content}

        except:
            if not skip_error:
                raise Exception("Get tree content failed.")
