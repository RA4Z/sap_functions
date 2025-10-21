import win32com


class Label:
    def __init__(self, session: win32com.client.CDispatch, window: int):
        self.session = session
        self.wnd = window

    def get_label_content(self) -> dict:
        """
        Store all the visible content from a SAP Label, the data will be stored and returned in a dictionary with
        'header' and 'content' items
        :return: A dictionary with 'header' and 'content' items
        """
        header = []
        content = []
        columns = []
        for c in range(1, 1000):
            try:
                cell = self.session.findById(f"wnd[{self.wnd}]/usr/lbl[{c},{1}]").text
                header.append(cell)
                columns.append(c)
            except:
                pass

        for i in range(2, 100):
            active_row = []
            for c in columns:
                try:
                    cell = self.session.findById(f"wnd[{self.wnd}]/usr/lbl[{c},{i}]").text
                    active_row.append(cell)
                except:
                    pass

            if not all(value is None for value in active_row):
                content.append(active_row)

        return {'header': header, 'content': content}
