"""Formatter tests. Run: python tests/test_format.py"""

from __future__ import annotations

import unittest

from xlvbesync.core.config import parse_indent
from xlvbesync.core.format import format_vba


def body(text: str, *, indent: str = "    ") -> str:
    formatted = format_vba("Option Explicit\r\n" + text, indent=indent)
    prefix = "Option Explicit\r\n"
    assert formatted.startswith(prefix)
    return formatted[len(prefix) :]


class ParseIndentTests(unittest.TestCase):
    def test_tab_aliases(self) -> None:
        self.assertEqual(parse_indent("tab"), "\t")
        self.assertEqual(parse_indent("TAB"), "\t")
        self.assertEqual(parse_indent("\t"), "\t")

    def test_spaces(self) -> None:
        self.assertEqual(parse_indent(2), "  ")
        self.assertEqual(parse_indent("2"), "  ")
        self.assertEqual(parse_indent("  "), "  ")


class FormatTests(unittest.TestCase):
    def test_one_line_property_does_not_nest(self) -> None:
        src = (
            "Public Property Get Question() As Long : Question = 0 : End Property\r\n"
            "' next\r\n"
            "Public Property Get Warning() As Long : Warning = 1 : End Property\r\n"
            "Public Property Get Danger() As Long : Danger = 2 : End Property\r\n"
        )
        out = body(src, indent="\t")
        self.assertIn("Public Property Get Question() As Long : Question = 0 : End Property", out)
        self.assertIn("Public Property Get Warning() As Long : Warning = 1 : End Property", out)
        self.assertNotIn("\tPublic Property Get Warning()", out)
        self.assertNotIn("\t\tPublic Property Get Danger()", out)

    def test_glues_type_suffix_dollar(self) -> None:
        src = (
            "If m_TempPath = \"\" Then m_TempPath = Environ $(\"Temp\")\r\n"
            "segment = Mid $(segment, 2)\r\n"
            "n = Len(Chr $(10) & Left $(s, 1) & Right $(s, 1) & Trim $(s))\r\n"
        )
        out = body(src)
        self.assertIn("Environ$(\"Temp\")", out)
        self.assertIn("Mid$(segment, 2)", out)
        self.assertIn("Chr$(10)", out)
        self.assertIn("Left$(s, 1)", out)
        self.assertIn("Right$(s, 1)", out)
        self.assertIn("Trim$(s)", out)
        self.assertNotIn(" $", out)

    def test_keeps_then_on_single_line_if(self) -> None:
        src = 'If m_AppDataLocalPath = "" Then m_AppDataLocalPath = Environ$("LocalAppData")\r\n'
        out = body(src, indent="\t")
        self.assertIn("Then m_AppDataLocalPath = Environ$(\"LocalAppData\")", out)
        self.assertTrue(out.startswith("If "))

    def test_file_number_hash_is_not_a_date_literal(self) -> None:
        src = (
            "Open targetFilePath For Binary Access Read As #fileNum\r\n"
            "fileContent = Space$(LOF(fileNum))\r\n"
            "Get #fileNum, , fileContent\r\n"
            "Close #fileNum\r\n"
            "Close #fileNum: On Error GoTo 0\r\n"
        )
        out = body(src, indent="\t")
        self.assertIn("As #fileNum\r\n", out)
        self.assertIn("fileContent = Space$(LOF(fileNum))\r\n", out)
        self.assertIn("Get #fileNum, , fileContent\r\n", out)
        self.assertIn("Close #fileNum\r\n", out)
        self.assertIn("Close #fileNum: On Error GoTo 0\r\n", out)
        self.assertNotIn("#fileNum#", out)

    def test_userform_header_is_not_formatted(self) -> None:
        src = (
            "VERSION 5.00\r\n"
            "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDemo \r\n"
            '   Caption         =   "Demo Form"\r\n'
            "   ClientHeight    =   3015\r\n"
            '   OleObjectBlob   =   "frmDemo.frx":0000\r\n'
            "End\r\n"
            'Attribute VB_Name = "frmDemo"\r\n'
            "Attribute VB_GlobalNameSpace = False\r\n"
            "\r\n"
            "Option Explicit\r\n"
            "\r\n"
            "Private Sub btnGo_Click()\r\n"
            'msgbox "clicked"\r\n'
            "End Sub\r\n"
        )
        out = format_vba(src, indent="    ")
        self.assertIn('   Caption         =   "Demo Form"\r\n', out)
        self.assertIn('   OleObjectBlob   =   "frmDemo.frx":0000\r\n', out)
        self.assertIn("Private Sub btnGo_Click()\r\n", out)
        self.assertIn('    msgbox "clicked"\r\n', out)
        self.assertNotIn("    Caption", out)

    def test_tab_indent_inside_procedure(self) -> None:
        src = (
            "Public Function Foo() As Long\r\n"
            "If x Then\r\n"
            "Foo = 1\r\n"
            "End If\r\n"
            "End Function\r\n"
        )
        out = body(src, indent="\t")
        self.assertIn("\tIf x Then\r\n", out)
        self.assertIn("\t\tFoo = 1\r\n", out)
        self.assertIn("\tEnd If\r\n", out)


if __name__ == "__main__":
    unittest.main()
