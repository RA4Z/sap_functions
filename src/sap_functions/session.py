from typing import Iterator, Literal, Any, TYPE_CHECKING

vKeys = {
    "Enter": 0,
    "F1": 1,
    "F2": 2,
    "F3": 3,
    "F4": 4,
    "F5": 5,
    "F6": 6,
    "F7": 7,
    "F8": 8,
    "F9": 9,
    "F10": 10,
    "Ctrl+S": 11,
    "F12": 12,
    "Shift+F1": 13,
    "Shift+F2": 14,
    "Shift+F3": 15,
    "Shift+F4": 16,
    "Shift+F5": 17,
    "Shift+F6": 18,
    "Shift+F7": 19,
    "Shift+F8": 20,
    "Shift+F9": 21,
    "Shift+Ctrl+0": 22,
    "Shift+F11": 23,
    "Shift+F12": 24,
    "Ctrl+F1": 25,
    "Ctrl+F2": 26,
    "Ctrl+F3": 27,
    "Ctrl+F4": 28,
    "Ctrl+F5": 29,
    "Ctrl+F6": 30,
    "Ctrl+F7": 31,
    "Ctrl+F8": 32,
    "Ctrl+F9": 33,
    "Ctrl+F10": 34,
    "Ctrl+F11": 35,
    "Ctrl+F12": 36,
    "Ctrl+Shift+F1": 37,
    "Ctrl+Shift+F2": 38,
    "Ctrl+Shift+F3": 39,
    "Ctrl+Shift+F4": 40,
    "Ctrl+Shift+F5": 41,
    "Ctrl+Shift+F6": 42,
    "Ctrl+Shift+F7": 43,
    "Ctrl+Shift+F8": 44,
    "Ctrl+Shift+F9": 45,
    "Ctrl+Shift+F10": 46,
    "Ctrl+Shift+F11": 47,
    "Ctrl+Shift+F12": 48
}

VKeyNames = Literal[
    "Enter", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10",
    "Ctrl+S", "F12", "Shift+F1", "Shift+F2", "Shift+F3", "Shift+F4", "Shift+F5",
    "Shift+F6", "Shift+F7", "Shift+F8", "Shift+F9", "Shift+Ctrl+0", "Shift+F11",
    "Shift+F12", "Ctrl+F1", "Ctrl+F2", "Ctrl+F3", "Ctrl+F4", "Ctrl+F5", "Ctrl+F6",
    "Ctrl+F7", "Ctrl+F8", "Ctrl+F9", "Ctrl+F10", "Ctrl+F11", "Ctrl+F12",
    "Ctrl+Shift+F1", "Ctrl+Shift+F2", "Ctrl+Shift+F3", "Ctrl+Shift+F4",
    "Ctrl+Shift+F5", "Ctrl+Shift+F6", "Ctrl+Shift+F7", "Ctrl+Shift+F8",
    "Ctrl+Shift+F9", "Ctrl+Shift+F10", "Ctrl+Shift+F11", "Ctrl+Shift+F12"
]

class SAPGridView:
    toolbarButtonCount: any
    columnOrder: list
    columnCount: int
    rowCount: int
    visibleRowCount: int
    def setCurrentCell(self, row: int, column: str) -> None: """If row and column identify a valid cell, this cell becomes the current cell. Otherwise, an exception is raised."""
    def modifyCheckbox(self, row: int, column: str, operator: bool) -> None: """Flag as true or false an checkbox component in the cell"""
    def triggerModified(self) -> None: """Update Grid to accept checkboxes changes"""
    def doubleClickCurrentCell(self) -> None: """This function emulates a mouse double click on the current cell"""
    def clearSelection(self) -> None: """Calling clearSelection removes all row, column and cell selections"""
    def clickCurrentCell(self) -> None: """This function emulates a mouse click on the current cell"""
    def contextMenu(self) -> None: """Calling contextMenu emulates the context menu request"""
    def selectAll(self) -> None: """This function selects the whole grid content (i.e. all rows and all columns)."""
    def selectContextMenuItem(self, item: str) -> None: """Select an item from the control’s context menu"""
    def selectColumn(self, column: str) -> None: """This function adds the specified column to the collection of the selected columns"""
    def getCellValue(self, row: int, column: str) -> str: """Returns the value of the cell as a string"""


class SAPGuiInfo:
    user: str
    transaction: str
    program: str
    systemName: str
    client: str
    language: str
    guiCodepage: float
    group: str
    isLowSpeedConnection: bool
    messageServer: str
    responseTime: str
    sessionNumber: float
    systemName: str
    systemSessionId: str


class SAPGuiFrameWindow:
    def __init__(self, sap_window_com_object):
        self._sap_window = sap_window_com_object

    def sendVKey(self, key: VKeyNames) -> None:
        """Sends the command corresponding to the specified key in string format."""
        key_code = vKeys.get(key)
        if key_code is None:
            raise ValueError(f"Invalid VKey command '{key}'.")
        self._sap_window.sendVKey(key_code)

    if TYPE_CHECKING:
        name: str
        text: str
        type: str
        def close(self) -> None: ...

    def __getattr__(self, name: str) -> Any:
        return getattr(self._sap_window, name)

class SAPGuiScrollbar:
    position: int
    minimum: int
    maximum: int
    pageSize: int
    range: int


class GuiComponent:
    id: str
    name: str
    type: str
    text: str
    parent: "GuiComponent"
    count: int
    charLeft: str
    children: "GuiComponent"
    horizontalScrollbar: SAPGuiScrollbar
    verticalScrollbar: SAPGuiScrollbar

    def clickCurrentCell(self): ...
    def Select(self) -> None: ...
    def sendVKey(self, key: int) -> None: ...
    def setFocus(self) -> None: ...
    def visualize(self) -> None: ...
    def containerType(self) -> str: ...
    def change(self) -> None: ...
    def press(self) -> None: ...
    def __getitem__(self, index: int) -> "GuiComponent": ...
    def __iter__(self) -> Iterator["GuiComponent"]: ...


class SAPGuiSession:
    def __init__(self, raw_session: Any):
        self._raw_session = raw_session

    @property
    def activeWindow(self) -> SAPGuiFrameWindow:
        """Returns the active window wrapped in our custom class to intercept sendVKey."""
        return SAPGuiFrameWindow(self._raw_session.ActiveWindow)

    if TYPE_CHECKING:
        info: SAPGuiInfo
        isActive: bool
        def CreateSession(self) -> None: ...
        def EndTransaction(self) -> None: ...
        def findById(self, identifier: str) -> GuiComponent: ...
        def startTransaction(self, transaction: str) -> None: ...

    def __getattr__(self, name: str) -> Any:
        return getattr(self._raw_session, name)