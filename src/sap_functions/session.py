
class GuiComponent:
    id: str
    name: str
    type: str
    text: str
    parent: "GuiComponent"
    children: tuple["GuiComponent", ...]
    def setFocus(self) -> None: ...
    def visualize(self) -> None: ...
    def containerType(self) -> str: ...
    def change(self) -> None: ...
    def press(self) -> None: ...


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


class SAPGuiSession:
    info: SAPGuiInfo
    activeWindow: SAPGuiWindow
    isActive: bool
    findById: object
    CreateSession: object
    EndTransaction: object
    StartTransaction: object

    def CreateSession(self) -> None: ...
    def EndTransaction(self) -> None: ...
    def findById(self, id: str) -> GuiComponent: ...
    def StartTransaction(self, transaction: str) -> None: ...
