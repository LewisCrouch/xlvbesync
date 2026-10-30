"""UserForm designer round-trip. Requires desktop Excel and Trust access to the VBA project.

Run: python tests/test_form_roundtrip.py
"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from xlvbesync.core.forms import splice_code


class SpliceTests(unittest.TestCase):
    def test_splice_replaces_only_the_code(self) -> None:
        exported = (
            "VERSION 5.00\r\n"
            "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDemo \r\n"
            '   Caption         =   "Demo Form"\r\n'
            "End\r\n"
            'Attribute VB_Name = "frmDemo"\r\n'
            "\r\n"
            "Option Explicit\r\n"
            "\r\n"
            'Private Sub Old()\r\nEnd Sub\r\n'
        )
        code = 'Option Explicit\r\n\r\nPrivate Sub btnGo_Click()\r\n    MsgBox "clicked"\r\nEnd Sub\r\n'
        out = splice_code(exported, "VERSION 5.00\r\nBegin X\r\nEnd\r\nAttribute VB_Name = \"x\"\r\n\r\n" + code)
        self.assertIn('Caption         =   "Demo Form"', out)
        self.assertIn('MsgBox "clicked"', out)
        self.assertNotIn("Private Sub Old()", out)
        self.assertTrue(out.startswith("VERSION 5.00\r\n"))

    def test_plain_code_is_the_body(self) -> None:
        exported = 'VERSION 5.00\r\nBegin X\r\nEnd\r\nAttribute VB_Name = "frmDemo"\r\n'
        out = splice_code(exported, "Option Explicit\r\n")
        self.assertIn("Option Explicit", out)
        self.assertIn('Attribute VB_Name = "frmDemo"', out)


class FormRoundTripTests(unittest.TestCase):
    def test_move_button_survives_export(self) -> None:
        from xlvbesync.ops.forms import FormHost

        host = FormHost()
        try:
            with tempfile.TemporaryDirectory(prefix="xlvbesync-form-test-") as tmp:
                dest = Path(tmp) / "frmMove.frm"
                host.create(dest, "frmMove")
                opened = host.open(dest)
                model = opened["model"]
                self.assertEqual(model["name"], "frmMove")
                model["children"].append(
                    {
                        "key": "btnGo",
                        "name": "btnGo",
                        "progId": "Forms.CommandButton.1",
                        "type": "CommandButton",
                        "opaque": False,
                        "properties": {
                            "Caption": "Go",
                            "Left": 24,
                            "Top": 36,
                            "Width": 72,
                            "Height": 24,
                        },
                        "children": [],
                    }
                )
                code = (
                    'Option Explicit\r\n\r\n'
                    'Private Sub btnGo_Click()\r\n'
                    '    MsgBox "clicked"\r\n'
                    "End Sub\r\n"
                )
                saved = host.save(opened["session"], model, code)
                self.assertEqual(saved["warnings"], [])
                host.close(opened["session"])

                text = dest.read_text(encoding="utf-8")
                self.assertIn("VERSION 5.00", text)
                self.assertIn('MsgBox "clicked"', text)
                self.assertTrue(dest.with_suffix(".frx").is_file())

                again = host.open(dest)
                try:
                    buttons = [c for c in again["model"]["children"] if c.get("name") == "btnGo"]
                    self.assertEqual(len(buttons), 1, again["model"])
                    props = buttons[0]["properties"]
                    self.assertEqual(props.get("Caption"), "Go", props)
                    self.assertAlmostEqual(float(props["Left"]), 24, delta=0.6)
                    self.assertAlmostEqual(float(props["Top"]), 36, delta=0.6)

                    props["Left"] = 48
                    props["Top"] = 12
                    host.save(again["session"], again["model"], code)
                finally:
                    host.close(again["session"])

                moved = host.open(dest)
                try:
                    button = next(c for c in moved["model"]["children"] if c.get("name") == "btnGo")
                    self.assertAlmostEqual(float(button["properties"]["Left"]), 48, delta=0.6)
                    self.assertAlmostEqual(float(button["properties"]["Top"]), 12, delta=0.6)
                    self.assertIn('MsgBox "clicked"', dest.read_text(encoding="utf-8"))
                finally:
                    host.close(moved["session"])
        finally:
            host.shutdown()


if __name__ == "__main__":
    unittest.main()
