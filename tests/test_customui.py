"""CustomUI packing and extraction tests. Run: python -m unittest tests/test_customui.py"""

from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from xlvbesync.core.customui import (
    extract_customui,
    find_customui_dir,
    is_customui_file,
    sync_customui,
)


class CustomUITests(unittest.TestCase):
    def test_is_customui_file(self) -> None:
        self.assertTrue(is_customui_file("src/LIS Tools/customUI/customUI14.xml"))
        self.assertTrue(is_customui_file("src/LIS Tools/customUI/_rels/customUI14.xml.rels"))
        self.assertTrue(is_customui_file("src/LIS Tools/customUI/images/folder_open.png"))
        self.assertTrue(is_customui_file(Path("customUI/customUI.xml")))

        # Ignored files
        self.assertFalse(is_customui_file("src/LIS Tools/customUI/~$customUI14.xml"))
        self.assertFalse(is_customui_file("src/LIS Tools/customUI/icon.tmp"))
        self.assertFalse(is_customui_file("src/LIS Tools/modMain.bas"))
        self.assertFalse(is_customui_file("src/LIS Tools/LIS Tools.xlsm"))

    def test_find_customui_dir(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            self.assertIsNone(find_customui_dir(root))

            cui = root / "customUI"
            cui.mkdir()
            self.assertEqual(find_customui_dir(root), cui)

    def test_sync_and_extract_customui(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            archive = base / "test.xlam"

            # Create minimal archive with _rels/.rels and [Content_Types].xml
            with zipfile.ZipFile(archive, "w") as z:
                z.writestr(
                    "_rels/.rels",
                    b'<?xml version="1.0" encoding="UTF-8"?>'
                    b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                    b'</Relationships>',
                )
                z.writestr(
                    "[Content_Types].xml",
                    b'<?xml version="1.0" encoding="UTF-8"?>'
                    b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                    b'<Default Extension="xml" ContentType="application/xml"/>'
                    b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                    b'</Types>',
                )

            # Create customUI dir with customUI14.xml and an image
            cui_dir = base / "src" / "test" / "customUI"
            (cui_dir / "images").mkdir(parents=True)
            (cui_dir / "customUI14.xml").write_text("<customUI>test</customUI>", encoding="utf-8")
            (cui_dir / "images" / "test.png").write_bytes(b"\x89PNG\r\n\x1a\n")

            # Sync into archive
            res = sync_customui(archive, cui_dir)
            self.assertTrue(res)

            # Verify contents inside archive
            with zipfile.ZipFile(archive, "r") as z:
                names = z.namelist()
                self.assertIn("customUI/customUI14.xml", names)
                self.assertIn("customUI/images/test.png", names)
                self.assertEqual(z.read("customUI/customUI14.xml"), b"<customUI>test</customUI>")

                # Verify relationship was injected into _rels/.rels
                rels = z.read("_rels/.rels").decode("utf-8")
                self.assertIn("customUI/customUI14.xml", rels)

                # Verify png content type was added
                ct = z.read("[Content_Types].xml").decode("utf-8")
                self.assertIn('Extension="png"', ct)

            # Test extraction
            extract_dest = base / "dest"
            extracted = extract_customui(archive, extract_dest)
            self.assertIn("customUI/customUI14.xml", extracted)
            self.assertIn("customUI/images/test.png", extracted)
            self.assertTrue((extract_dest / "customUI" / "customUI14.xml").is_file())
            self.assertEqual(
                (extract_dest / "customUI" / "customUI14.xml").read_text(encoding="utf-8"),
                "<customUI>test</customUI>",
            )


if __name__ == "__main__":
    unittest.main()
