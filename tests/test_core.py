from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path


SCRIPTS = Path(__file__).resolve().parents[1] / "skill" / "scripts"
sys.path.insert(0, str(SCRIPTS))

from combine_sunday_ppt import (  # noqa: E402
    PowerShellDeck,
    chunk_reading,
    image_dimensions,
    split_announcement,
)
from spring_bible import Verse  # noqa: E402


class CoreTests(unittest.TestCase):
    def test_skill_does_not_bundle_powerpoint_files(self) -> None:
        project_root = Path(__file__).resolve().parents[1]
        self.assertEqual(list(project_root.joinpath("skill").rglob("*.pptx")), [])

    def test_reading_uses_at_most_four_verses_per_page(self) -> None:
        verses = [Verse(15, number, f"verse {number}") for number in range(10, 21)]
        pages = chunk_reading(verses)
        self.assertEqual([len(page) for page in pages], [4, 4, 3])

    def test_short_announcement_stays_on_one_page(self) -> None:
        self.assertEqual(split_announcement("短報告。"), ["短報告。"])

    def test_png_dimensions_do_not_require_platform_tools(self) -> None:
        payload = b"\x89PNG\r\n\x1a\n" + b"\x00" * 8 + (640).to_bytes(4, "big") + (360).to_bytes(4, "big")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "image.png"
            path.write_bytes(payload)
            self.assertEqual(image_dimensions(path), (640, 360))

    def test_windows_backend_uses_native_insert_from_file(self) -> None:
        deck = PowerShellDeck(Path("C:/tmp/output.pptx"), 17)
        deck.append_source(Path("C:/tmp/sermon.pptx"), 2)
        script = deck.finish()
        self.assertIn("Slides.InsertFromFile", script)
        self.assertIn("sermon.pptx", script)


if __name__ == "__main__":
    unittest.main()
