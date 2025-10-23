import win32com


class Node:
    def __init__(self, node_obj: win32com.client.CDispatch):
        self.node_obj = node_obj

    def select_node(self, node_text: str, target_index: int = 0):
        """
        Select a specific Node based on the text inside of it
        :param node_text: The text in the desired Node
        :param target_index: Target index, determines how many occurrences precede the desired field
        """
        parent = self.node_obj.GetAllNodeKeys()
        for item in parent:
            text = self.node_obj.GetNodeTextByKey(item)
            if node_text in text:
                if target_index == 0:
                    self.node_obj.SelectNode(item)
                    return
                else:
                    target_index -= 1

    def expand_selected_node(self):
        """
        Expand the Node selected previously
        """
        node_id = self.node_obj.SelectedNode
        self.node_obj.expandNode(node_id)

    def get_node_content(self) -> list:
        """
        Get all Nodes names in a list format
        :return: A list of string with every Node text
        """
        results = []
        parent = self.node_obj.GetAllNodeKeys()
        for item in parent:
            text = self.node_obj.GetNodeTextByKey(item)
            results.append(text)

        return results
