from types import SimpleNamespace
import unittest
from unittest.mock import Mock, call

from src.sap_functions.sap import SAP
from src.sap_functions.session import SAPGuiComponent, SAPGuiSession


class TestSessionProxy(unittest.TestCase):
    def test_component_forwards_methods_and_attributes(self):
        raw = Mock()
        component = SAPGuiComponent(raw)
        component.press()
        component.Select()
        component.setFocus()
        component.Text = "valor"
        raw.press.assert_called_once_with()
        raw.Select.assert_called_once_with()
        raw.setFocus.assert_called_once_with()
        self.assertEqual(raw.Text, "valor")

    def test_component_forwards_collection_operations(self):
        component = SAPGuiComponent(["primeiro", "segundo"])
        self.assertEqual(component[1], "segundo")
        self.assertEqual(list(component), ["primeiro", "segundo"])

    def test_component_send_vkey_accepts_names_and_numbers(self):
        raw = Mock()
        component = SAPGuiComponent(raw)
        component.sendVKey("F8")
        component.sendVKey(4)
        self.assertEqual(raw.sendVKey.call_args_list, [call(8), call(4)])
        with self.assertRaises(ValueError):
            component.sendVKey("invalida")

    def test_insert_variant_uses_real_controls(self):
        raw = Mock()
        raw.ActiveWindow = SimpleNamespace(name="wnd[0]")
        variant = SimpleNamespace(Text="")
        creator = SimpleNamespace(Text="usuario")
        open_button = Mock()
        confirm_button = Mock()
        open_button.press.side_effect = lambda: setattr(raw.ActiveWindow, "name", "wnd[1]")
        confirm_button.press.side_effect = lambda: setattr(raw.ActiveWindow, "name", "wnd[0]")
        controls = {
            "wnd[0]/tbar[1]/btn[17]": open_button,
            "wnd[1]/usr/txtV-LOW": variant,
            "wnd[1]/usr/txtENAME-LOW": creator,
            "wnd[1]/tbar[0]/btn[8]": confirm_button,
        }
        raw.findById.side_effect = controls.__getitem__
        sap = SAP.__new__(SAP)
        sap.session = SAPGuiSession(raw)
        sap.insert_variant("MINHA_VARIANTE")
        open_button.press.assert_called_once_with()
        confirm_button.press.assert_called_once_with()
        self.assertEqual(variant.Text, "MINHA_VARIANTE")
        self.assertEqual(creator.Text, "")
