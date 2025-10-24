import win32com


# https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/2e44c4f890524686977e9729565f7824.html?locale=en-US
class Label:
    def __init__(self, session: win32com.client.CDispatch, window: int):
        self.session = session
        self.wnd = window

    def get_all_screen_labels(self) -> list:
        """
        This function will return each label row in the SAP Screen
        :return: A list with lists
        """
        self.session.findById(f"wnd[{self.wnd}]/usr").VerticalScrollbar.Position = 0
        content = []
        columns = []
        children = self.session.findById(f"wnd[0]/usr").children
        for field in children:
            if field.Type == 'GuiLabel':
                if field.CharLeft not in columns:
                    columns.append(field.CharLeft)

        while True:
            for i in range(2, 100):
                active_row = []
                for c in columns:
                    try:
                        cell = self.session.findById(f"wnd[{self.wnd}]/usr/lbl[{c},{i}]").text.strip()
                        active_row.append(cell)
                    except:
                        pass

                if not all(value is None for value in active_row):
                    content.append(active_row)

            max_scroll = self.session.findById(f"wnd[{self.wnd}]/usr").VerticalScrollbar.Maximum
            pos_scroll = self.session.findById(f"wnd[{self.wnd}]/usr").VerticalScrollbar.Position

            if max_scroll == pos_scroll:
                break
            else:
                self.session.findById(f"wnd[{self.wnd}]").sendVKey(82)
        return content

    def get_label_content(self) -> dict:
        """
        Store all the content from a SAP Label, the data will be stored and returned in a dictionary with
        'header' and 'content' items
        :return: A dictionary with 'header' and 'content' items
        """
        self.session.findById(f"wnd[{self.wnd}]/usr").VerticalScrollbar.Position = 0
        finished_collecting = False
        header = []
        content = []
        columns = []

        children = self.session.findById(f"wnd[0]/usr").children
        for field in children:
            if field.Type == 'GuiLabel':
                if field.CharLeft not in columns:
                    columns.append(field.CharLeft)

        for header_row_index in range(1, 4):
            for c in columns:
                try:
                    cell = self.session.findById(f"wnd[{self.wnd}]/usr/lbl[{c},{header_row_index}]").text.strip()
                    header.append(cell)
                except:
                    pass
            if len(header) > 0:
                break

        while True:
            if finished_collecting:
                break
            for i in range(2, 100):
                active_row = []
                for c in columns:
                    try:
                        cell = self.session.findById(f"wnd[{self.wnd}]/usr/lbl[{c},{i}]").text.strip()
                        active_row.append(cell)
                    except:
                        pass

                if not all(value is None for value in active_row):
                    content.append(active_row)

            max_scroll = self.session.findById(f"wnd[{self.wnd}]/usr").VerticalScrollbar.Maximum
            pos_scroll = self.session.findById(f"wnd[{self.wnd}]/usr").VerticalScrollbar.Position

            if max_scroll == pos_scroll:
                break
            else:
                self.session.findById(f"wnd[{self.wnd}]").sendVKey(82)

        return {'header': header, 'content': content}
