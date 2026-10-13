import json
from pathlib import Path
import re
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))
from build import demo_rows, render, workbook_rows  # noqa: E402
from classify import classify  # noqa: E402


class ClassifierTests(unittest.TestCase):
    def test_explicit_graduate_limits(self):
        for note in ["限2027届高校毕业生报考。", "2027届高校毕业生，男性。", "仅招录2027届高校毕业生；"]:
            with self.subTest(note=note):
                self.assertEqual(classify(note), "2027")
        for note in ["限应届高校毕业生报考。", "仅招录应届高校毕业生。", "应届高校毕业生（需核对其他条件）"]:
            with self.subTest(note=note):
                self.assertEqual(classify(note), "fresh")

    def test_conditional_requirements_are_not_limits(self):
        for note in [
            "非应届毕业生须有相关工作经历。", "非2027届高校毕业生需专业从事相关工作。",
            "2027届高校毕业生需通过考试，其他考生须持证。", "不限应届身份。",
            "如考生为应届高校毕业生，请注明身份。", "其他岗位要求以公告为准。",
        ]:
            with self.subTest(note=note):
                self.assertEqual(classify(note), "unspecified")

    def test_non_graduate_only_explicit(self):
        self.assertEqual(classify("招录对象为非应届高校毕业生。"), "non")
        self.assertEqual(classify("限非2027届高校毕业生报考。"), "non")


class BuilderTests(unittest.TestCase):
    def test_demo_render_has_no_real_departments(self):
        rows, sheets = demo_rows()
        self.assertEqual(len(rows), 6)
        self.assertEqual(sum(row[6] for row in rows), 12)
        self.assertEqual([row[20] for row in rows], ["2027", "fresh", "unspecified", "2027", "unspecified", "fresh"])
        with tempfile.TemporaryDirectory() as directory:
            dest = Path(directory) / "demo.html"
            render(rows, sheets, "虚构样例数据", "—", True, dest)
            content = dest.read_text(encoding="utf-8")
            self.assertIn('"demo":true', content)
            self.assertIn("虚构机关甲", content)
            self.assertNotIn("国家税务总局", content)
            self.assertEqual(len(re.findall(r"__DATA_JSON__", content)), 0)

    def test_duplicate_ids_rejected(self):
        rows, sheets = demo_rows()
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(ValueError, "重复"):
                render(rows + [rows[0]], sheets, "虚构样例数据", "—", True, Path(directory) / "x.html")


if __name__ == "__main__":
    unittest.main()
