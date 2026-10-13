#!/usr/bin/env python3
"""Compare every published position field with a local .xls or a ZIP containing it.

Run locally; the original workbook is deliberately not committed to the repository.
"""
from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import sys
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
from classify import classify  # noqa: E402

# Compare the original 19 columns copied to each row, including all position
# descriptions and notes. Remaining columns (province / identity) are derived.
FIELD_MAP = {
    0: 1, 1: 2, 2: 4, 3: 8, 4: 20, 6: 11, 7: 12, 8: 13, 9: 14,
    10: 15, 11: 16, 12: 17, 13: 5, 14: 9, 15: 6, 16: 18,
    17: 7, 18: 22, 19: 0,
}
HEADERS = {
    0: "部门代码", 1: "部门名称", 2: "用人司局", 4: "招考职位", 8: "职位代码",
    11: "招考人数", 12: "专业", 13: "学历", 15: "政治面貌", 16: "基层工作最低年限",
    17: "服务基层项目工作经历", 20: "工作地点", 22: "备注",
}
PROVINCES = (
    "北京市", "天津市", "上海市", "重庆市", "河北省", "山西省", "辽宁省", "吉林省", "黑龙江省",
    "江苏省", "浙江省", "安徽省", "福建省", "江西省", "山东省", "河南省", "湖北省", "湖南省",
    "广东省", "海南省", "四川省", "贵州省", "云南省", "陕西省", "甘肃省", "青海省",
    "内蒙古自治区", "广西壮族自治区", "西藏自治区", "宁夏回族自治区", "新疆维吾尔自治区",
)


def load_workbook(path: Path):
    try:
        import xlrd
    except ImportError as error:
        raise SystemExit("缺少 xlrd，请运行：python3 -m pip install -r requirements.txt") from error
    if path.suffix.lower() == ".zip":
        with ZipFile(path) as archive:
            names = [name for name in archive.namelist() if name.lower().endswith(".xls")]
            if len(names) != 1:
                raise ValueError("ZIP 中必须恰有一个 .xls 文件")
            content = archive.read(names[0])
    elif path.suffix.lower() == ".xls":
        content = path.read_bytes()
    else:
        raise ValueError("数据源只能是 .xls 或仅含一份 .xls 的 .zip")
    return xlrd.open_workbook(file_contents=content, on_demand=True), hashlib.sha256(content).hexdigest()


def embedded_data(page: Path):
    text = page.read_text(encoding="utf-8")
    match = re.search(r"const SOURCE=(\{.*?\});\s*const R=", text, re.DOTALL)
    if not match:
        raise ValueError("未找到嵌入在 HTML 内的岗位数据")
    data = json.loads(match.group(1))
    return data["rows"], data["meta"]


def verify(input_path: Path, page_path: Path) -> list[tuple[str, int, int]]:
    workbook, digest = load_workbook(input_path)
    rows, meta = embedded_data(page_path)
    count = 0
    copies = 0
    counts: list[tuple[str, int, int]] = []
    identities: Counter[str] = Counter()
    provinces: Counter[str] = Counter()
    ids = set()
    for sheet in workbook.sheets():
        if sheet.nrows == 0:
            continue
        header = sheet.row_values(1)
        for index, name in HEADERS.items():
            if header[index] != name:
                raise ValueError(f"{sheet.name} 的 {name} 表头不匹配")
        n, seats = 0, 0
        for line in range(2, sheet.nrows):
            cells = sheet.row_values(line)
            if not str(cells[8]).strip():
                continue
            if count >= len(rows):
                raise ValueError("原表行数多于网页")
            row = rows[count]
            if len(row) != 21:
                raise ValueError(f"第 {count + 1} 个网页岗位字段长度错误")
            for index, column in FIELD_MAP.items():
                expected = int(cells[column]) if index == 6 else str(cells[column]).strip()
                if row[index] != expected:
                    raise ValueError(f"{sheet.name} 第 {line + 1} 行原表字段 {header[column]} 与网页不同")
                copies += 1
            province = next((p for p in PROVINCES if row[4].startswith(p)), "未识别地区")
            if row[5] != province:
                raise ValueError(f"第 {count + 1} 个网页岗位工作省份提取错误")
            if row[20] != classify(str(cells[22])):
                raise ValueError(f"第 {count + 1} 个网页岗位应届身份分类与当前规则不一致")
            key = (row[19], row[3])
            if key in ids:
                raise ValueError(f"岗位部门代码+职位代码重复：{key}")
            ids.add(key)
            count += 1
            n += 1
            seats += row[6]
            identities[row[20]] += 1
            provinces[row[5]] += 1
        counts.append((sheet.name, n, seats))
    if count != len(rows) or sum(seats for _, _, seats in counts) != meta["seats"] or count != meta["positions"]:
        raise ValueError("网页岗位行数或招录人数与 Excel 不一致")
    if [s[0] for s in counts] != meta["sheets"] or dict(identities) != meta["identityCounts"]:
        raise ValueError("网页元数据与原表不一致")
    if sorted(provinces) != meta["provinces"]:
        raise ValueError("网页省份列表与原表不一致")
    print("核对通过：")
    for name, n, seats in counts:
        print(f"  {name}: {n:,} 个岗位 / {seats:,} 个招录名额")
    print(f"  合计：{count:,} 个岗位 / {meta['seats']:,} 个招录名额")
    print(f"  {copies:,} 个原表字段逐个相同；{count * 2:,} 个派生字段与当前规则一致；岗位键无重复。")
    print(f"  应届标记：{dict(identities)}；省份 {len(provinces)} 种。")
    if digest == meta["sourceSha256"]:
        print(f"  源文件与页面记录的 XLS SHA-256 一致：{digest}")
    else:
        print(f"  注意：XLS 二进制哈希不同（当前 {digest}，页面记录 {meta['sourceSha256']}），"
              "但上述每个岗位的19个原表字段完全一致。文件可能被重新保存。")
    return counts


def main() -> None:
    parser = argparse.ArgumentParser(description="核对 HTML 内嵌岗位是否与本地 XLS / ZIP 全部对应")
    parser.add_argument("--input", type=Path, required=True, help=".xls 原件或包含一份 .xls 的 .zip")
    parser.add_argument("--html", type=Path, default=Path("index.html"), help="公开的单文件 HTML")
    args = parser.parse_args()
    verify(args.input.expanduser(), args.html.expanduser())


if __name__ == "__main__":
    main()
