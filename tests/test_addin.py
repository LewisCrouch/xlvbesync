"""Unit tests for add-in helpers. Run: python -m unittest tests/test_addin.py"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from xlvbesync.core.addin import (
    AddinHandle,
    discard_backup,
    is_addin,
    replace_file,
)


class AddinTests(unittest.TestCase):
    def test_is_addin(self) -> None:
        self.assertTrue(is_addin("test.xlam"))
        self.assertTrue(is_addin("test.xla"))
        self.assertTrue(is_addin(Path("C:/path/to/MyAddin.XLAM")))
        self.assertFalse(is_addin("test.xlsm"))
        self.assertFalse(is_addin("test.xlsx"))
        self.assertFalse(is_addin("test.bas"))

    def test_addin_handle_stores_path(self) -> None:
        p = Path("C:/AddIns/Tool.xlam")
        handle = AddinHandle(com=None, name="Tool.xlam", was_installed=True, path=p)
        self.assertEqual(handle.name, "Tool.xlam")
        self.assertTrue(handle.was_installed)
        self.assertEqual(handle.path, p)

    def test_replace_file_atomic_with_backup(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            dest = base / "dest.xlam"
            source = base / "source.xlam"

            dest.write_text("initial content", encoding="utf-8")
            source.write_text("updated content", encoding="utf-8")

            backup = replace_file(source, dest, keep_backup=True)
            self.assertIsNotNone(backup)
            self.assertEqual(dest.read_text(encoding="utf-8"), "updated content")
            self.assertTrue(backup.is_file())
            self.assertEqual(backup.read_text(encoding="utf-8"), "initial content")

            discard_backup(backup)
            self.assertFalse(backup.is_file())


if __name__ == "__main__":
    unittest.main()
