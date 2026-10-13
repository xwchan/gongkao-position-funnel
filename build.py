#!/usr/bin/env python3
"""Build a self-contained position-funnel report from an .xls workbook or fictional demo data."""
from __future__ import annotations

import argparse
from collections import Counter
from datetime import datetime
import hashlib
import json
from pathlib import Path
import sys
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
from classify import classify  # noqa: E402

TEMPLATE = Path(__file__).resolve().parent / "src" / "index.html"
PROVINCES = (
    "北京市", "天津市", "上海市", "重庆市", "河北省", "山西省", "辽宁省", "吉林省", "黑龙江省",
    "江苏省", "浙江省", "安徽省", "福建省", "江西省", "山东省", "河南省", "湖北省", "湖南省",
    "广东省", "海南省", "四川省", "贵州省", "云南省", "陕西省", "甘肃省", "青海省",
    "内蒙古自治区", "广西壮族自治区", "西藏自治区", "宁夏回族自治区", "新疆维吾尔自治区",
)
# [department, bureau, position, position code, location, province, openings,
#  major, education, degree, political status, work years, project experience,
#  attribute, agency level, distribution, professional test, description, notes,
#  department code, graduate-status classification]
COLUMN_IDS = (1, 2, 4, 8, 20, 11, 12, 13, 14, 15, 16, 17, 5, 9, 6, 18, 7, 22, 0)
EXPECTED_HEADERS = {
    0: "部门代码", 1: "部门名称", 2: "用人司局", 4: "招考职位", 8: "职位代码",
    11: "招考人数", 12: "专业", 13: "学历", 15: "政治面貌", 16: "基层工作最低年限",
    17: "服务基层项目工作经历", 20: "工作地点", 22: "备注",
}


def normalize_province(location: str) -> str:
    return next((province for province in PROVINCES if location.startswith(province)), "未识别地区")


def workbook_rows(path: Path) -> tuple[list[list], list[str]]:
    try:
        import xlrd
    except ImportError as error:
        raise SystemExit("缺少 xlrd，请先运行：python3 -m pip install -r requirements.txt") from error
    if path.suffix.lower() == ".zip":
        with ZipFile(path) as archive:
            members = [entry for entry in archive.infolist() if entry.filename.lower().endswith(".xls")]
            if len(members) != 1:
                raise ValueError("ZIP 中必须恰好包含一份 XLS 招考简章")
            workbook = xlrd.open_workbook(file_contents=archive.read(members[0]), on_demand=True)
    else:
        workbook = xlrd.open_workbook(str(path), on_demand=True)
    rows: list[list] = []
    sheets: list[str] = []
    for sheet in workbook.sheets():
        if sheet.nrows < 3:
            continue
        headers = sheet.row_values(1)
        if any(len(headers) <= col or str(headers[col]).strip() != label for col, label in EXPECTED_HEADERS.items()):
            raise ValueError(f"工作表「{sheet.name}」表头不符合招考简章格式（期望第 2 行为字段名）")
        sheets.append(sheet.name)
        for row_number in range(2, sheet.nrows):
            raw = sheet.row_values(row_number)
            if not str(raw[8]).strip():
                continue
            try:
                openings = int(raw[11])
                if openings < 1 or openings != float(raw[11]):
                    raise ValueError("名额非正整数")
            except (ValueError, TypeError) as error:
                raise ValueError(f"{sheet.name} 第 {row_number + 1} 行招考人数无效：{raw[11]}") from error
            cells = [str(raw[column]).strip() for column in COLUMN_IDS]
            cells[5] = openings
            location = cells[4]
            note = cells[17]
            rows.append(cells[:5] + [normalize_province(location)] + cells[5:] + [classify(note)])
    return rows, sheets


def demo_rows() -> tuple[list[list], list[str]]:
    # Obviously fictional examples: never misrepresent them as an official recruiting list.
    examples = [
        ("虚构机关甲", "综合处", "文书岗", "演示甲001", "北京市", 2, "法学", "本科及以上", "不限", "无限制", "无限制", "限2027届高校毕业生报考。"),
        ("虚构机关乙", "研究室", "研究岗", "演示乙001", "北京市", 1, "经济学", "硕士研究生及以上", "中共党员", "二年", "无限制", "限应届高校毕业生报考。"),
        ("虚构机关丙", "业务处", "综合岗", "演示丙001", "广东省广州市", 3, "不限专业", "本科及以上", "不限", "无限制", "无限制", "请核对具体要求。"),
        ("虚构机关丁", "业务处", "服务岗", "演示丁001", "北京市", 1, "计算机类", "仅限本科", "不限", "无限制", "无限制", "招录对象为2027届高校毕业生。"),
        ("虚构机关戊", "档案处", "信息岗", "演示戊001", "上海市", 4, "计算机科学与技术", "本科及以上", "不限", "二年", "无限制", "非应届毕业生须有相关工作经历。"),
        ("虚构机关己", "事业处", "管理岗", "演示己001", "北京市", 1, "工商管理", "本科及以上", "不限", "无限制", "无限制", "应届高校毕业生，工作地点为北京市。"),
    ]
    rows = []
    for i, (dept, bureau, title, code, location, seats, major, education, politics, years, project, note) in enumerate(examples):
        rows.append([dept, bureau, title, code, location, normalize_province(location), seats,
                     major, education, "与最高学历相对应的学位", politics, years, project,
                     "普通职位", "县（区）级及以下", "其他职位", "否", "虚构数据，仅供体验交互。", note,
                     f"DEMO-{i + 1:03d}", classify(note)])
    return rows, ["虚构样例"]


def render(rows: list[list], sheets: list[str], source: str, modified: str, demo: bool, output: Path, source_sha256: str = "") -> None:
    if not rows:
        raise ValueError("没有可供生成报表的岗位")
    ids = [(row[19], row[3]) for row in rows]
    if len(ids) != len(set(ids)):
        raise ValueError("发现重复的部门代码 + 职位代码，请核对原始数据")
    meta = {
        "source": source, "sheets": sheets, "positions": len(rows),
        "seats": sum(row[6] for row in rows), "provinces": sorted({row[5] for row in rows}),
        "attributes": sorted({row[13] for row in rows}),
        "identityCounts": dict(Counter(row[20] for row in rows)),
        "fileMtime": modified, "demo": demo, "sourceSha256": source_sha256,
    }
    payload = json.dumps({"rows": rows, "meta": meta}, ensure_ascii=False, separators=(",", ":"))
    payload = (payload.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
               .replace("\u2028", "\\u2028").replace("\u2029", "\\u2029"))
    template = TEMPLATE.read_text(encoding="utf-8")
    if template.count("__DATA_JSON__") != 1:
        raise ValueError("页面模板的数据占位符缺失或重复")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(template.replace("__DATA_JSON__", payload), encoding="utf-8")
    print(f"完成：{output}（{len(rows):,} 个岗位，{meta['seats']:,} 个招录名额）")


def main() -> None:
    parser = argparse.ArgumentParser(description="从 2027 国考招考简章 .xls 生成离线数据漏斗 HTML")
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--input", type=Path, help="本地招考简章 .xls 或含单份 .xls 的 .zip（不会被上传）")
    source.add_argument("--demo", action="store_true", help="生成醒目标注的虚构样例（不含真实岗位）")
    parser.add_argument("--output", type=Path, help="输出文件路径，默认 dist/index.html")
    args = parser.parse_args()
    if args.demo:
        rows, sheets = demo_rows()
        render(rows, sheets, "虚构样例数据", "—", True, args.output or Path("dist/demo.html"))
    else:
        path = args.input.expanduser().resolve()
        if not path.is_file() or path.suffix.lower() not in (".xls", ".zip"):
            parser.error("--input 必须是存在的 .xls 或 .zip 文件")
        rows, sheets = workbook_rows(path)
        if path.suffix.lower() == ".zip":
            with ZipFile(path) as archive:
                member = next(info for info in archive.infolist() if info.filename.lower().endswith(".xls"))
                content = archive.read(member)
                modified = datetime(*member.date_time).strftime("%Y-%m-%d %H:%M")
                source_name = path.stem + ".xls"
                digest = hashlib.sha256(content).hexdigest()
        else:
            modified = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
            source_name = path.name
            digest = hashlib.sha256()
            with path.open("rb") as stream:
                for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                    digest.update(chunk)
            digest = digest.hexdigest()
        render(rows, sheets, source_name, modified, False, args.output or Path("dist/index.html"), digest)


if __name__ == "__main__":
    main()
