from typing import Iterator


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


class SAPGuiWindow:
    name: str
    text: str


class SAPGuiScrollbar:
    position: int
    maximum: int


class GuiComponent:
    id: str
    name: str
    type: str
    text: str
    parent: "GuiComponent"
    count: int
    charLeft: str
    children: "GuiComponent"
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
    info: SAPGuiInfo
    activeWindow: SAPGuiWindow
    isActive: bool
    findById: object
    CreateSession: object
    EndTransaction: object
    startTransaction: object

    def CreateSession(self) -> None: ...
    def EndTransaction(self) -> None: ...
    def findById(self, identifier: str) -> GuiComponent: ...
    def startTransaction(self, transaction: str) -> None: ...
