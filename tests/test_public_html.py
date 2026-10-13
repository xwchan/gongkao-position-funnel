"""Guardrails for the checked-in, self-contained public page."""
import hashlib
import json
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
PAGE = ROOT / "index.html"
EXPECTED_SHA = "6b3da70baabe31a4a202c6b1179040db7003a956bc85c79cd35e3fb739d9520a"


class PublicPageTests(unittest.TestCase):
    def test_complete_single_file_and_provenance(self):
        content = PAGE.read_text(encoding="utf-8")
        self.assertNotIn("__DATA_JSON__", content)
        match = re.search(r"const SOURCE=(\{.*?\});\s*const R=", content, re.DOTALL)
        self.assertIsNotNone(match)
        data = json.loads(match.group(1))
        self.assertEqual(len(data["rows"]), 19426)
        self.assertEqual(sum(row[6] for row in data["rows"]), 35691)
        self.assertEqual(data["meta"]["sourceSha256"], EXPECTED_SHA)
        self.assertFalse(data["meta"]["demo"])
        self.assertEqual(len({(row[19], row[3]) for row in data["rows"]}), 19426)
        self.assertFalse(re.search(r"<script[^>]+src=|<link[^>]+href=|\bfetch\(|\bXMLHttpRequest\b", content))

    def test_compact_enough_for_github(self):
        self.assertLess(PAGE.stat().st_size, 20 * 1024 * 1024)


if __name__ == "__main__":
    unittest.main()
