#!/usr/bin/env python3
"""从 2027 届指导手册 docx 中按“节”切出封面、附件表与样式样例。

原文件只读，本脚本只读取、不修改；输出到 2027/模板/。

用法：
    python3 extract_templates.py [--manual <手册docx>] [--out <模板目录>]
"""

from __future__ import annotations

import argparse
import copy
import csv
import json
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

HERE = Path(__file__).resolve().parent
YEAR_DIR = HERE.parent.parent / "assets" / "2027"


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def paragraph_text(p: etree._Element) -> str:
    return "".join(p.itertext()).strip()


def section_ranges(body: etree._Element) -> tuple[list[tuple[int, int]], int | None]:
    """返回 [(起始下标, 结束下标)] 与 body 级 sectPr 的下标。

    每个结束下标的元素是带 pPr/sectPr 的段落；最后一节由 body 级 sectPr 收尾，
    结束下标为其前一个元素。
    """
    children = list(body)
    ranges: list[tuple[int, int]] = []
    start = 0
    final_sect = None
    for i, ch in enumerate(children):
        if ch.tag == q("p"):
            ppr = ch.find(q("pPr"))
            if ppr is not None and ppr.find(q("sectPr")) is not None:
                ranges.append((start, i))
                start = i + 1
        elif ch.tag == q("sectPr"):
            final_sect = i
            ranges.append((start, i - 1))
    return ranges, final_sect


def strip_page_furniture(sect: etree._Element) -> None:
    """去掉页眉页脚与页码设置，让模板保持“无页码、无页眉”。"""
    for tag in ("footerReference", "headerReference", "titlePg", "pgNumType"):
        for el in sect.findall(q(tag)):
            sect.remove(el)


def build_section_docx(
    src: Path, out: Path, first: int, last: int, final_sect: int | None
) -> None:
    with zipfile.ZipFile(src) as zin:
        doc_xml = zin.read("word/document.xml")
        root = etree.fromstring(doc_xml)
        body = root.find(q("body"))
        children = list(body)
        ranges, _ = section_ranges(body)

        # 收集目标节的所有元素
        elems: list[etree._Element] = []
        for sec in range(first, last + 1):
            a, b = ranges[sec]
            elems.extend(children[a : b + 1])

        # 末节的 sectPr：来自该节结束段落，或 body 级 sectPr
        last_end = ranges[last][1]
        last_para = None
        if 0 <= last_end < len(children) and children[last_end].tag == q("p"):
            ppr = children[last_end].find(q("pPr"))
            if ppr is not None and ppr.find(q("sectPr")) is not None:
                last_para = children[last_end]
        if last_para is not None:
            last_sect = last_para.find(q("pPr")).find(q("sectPr"))
        else:
            last_sect = body.find(q("sectPr"))
        if last_sect is None:
            raise RuntimeError(f"第 {last} 节找不到 sectPr")
        final_copy = copy.deepcopy(last_sect)
        strip_page_furniture(final_copy)

        # 清理范围内的分节符：保留中间分节（多页表），删除末节分节符本身
        cleaned: list[etree._Element] = []
        for el in elems:
            if el.tag == q("p"):
                ppr = el.find(q("pPr"))
                sect = ppr.find(q("sectPr")) if ppr is not None else None
                if sect is not None:
                    if el is last_para:
                        ppr.remove(sect)
                        # 末节分节段落若已无内容则丢弃
                        if not paragraph_text(el) and len(el.findall(q("r"))) == 0:
                            continue
                    else:
                        strip_page_furniture(sect)
            cleaned.append(el)

        # 重建 body
        for ch in list(body):
            body.remove(ch)
        for el in cleaned:
            body.append(copy.deepcopy(el))
        body.append(final_copy)

        new_doc = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        out.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = (
                    new_doc
                    if item.filename == "word/document.xml"
                    else zin.read(item.filename)
                )
                zout.writestr(item, data)


def clear_cell_text(cell: etree._Element) -> None:
    """清空单元格内容，保留段落与字体格式。"""
    for p in cell.findall(q("p")):
        runs = p.findall(q("r"))
        for r in runs:
            for t in r.findall(q("t")):
                t.text = ""
            for br in r.findall(q("br")):
                r.remove(br)


def blank_cover_values(docx: Path) -> None:
    """把封面示例里的填写说明替换为空白，保留单元格格式。"""
    with zipfile.ZipFile(docx) as zin:
        root = etree.fromstring(zin.read("word/document.xml"))
        body = root.find(q("body"))
        tbls = body.findall(f".//{q('tbl')}")
        for tbl in tbls:
            rows = tbl.findall(q("tr"))
            for row in rows:
                cells = row.findall(q("tc"))
                if len(cells) != 2:
                    continue
                left = "".join(cells[0].itertext()).strip()
                right = "".join(cells[1].itertext()).strip()
                if left in {
                    "题目",
                    "学院",
                    "年级",
                    "专业",
                    "班级",
                    "学号",
                    "姓名",
                    "指导教师",
                }:
                    if not right:
                        continue
                    clear_cell_text(cells[1])
        new_doc = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        tmp = docx.with_suffix(".tmp.docx")
        with (
            zipfile.ZipFile(docx) as zin,
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout,
        ):
            for item in zin.infolist():
                data = (
                    new_doc
                    if item.filename == "word/document.xml"
                    else zin.read(item.filename)
                )
                zout.writestr(item, data)
        shutil.move(tmp, docx)


def export_class_table(docx: Path, out_csv: Path) -> None:
    """导出“专业、班级信息对照表”为 CSV：学院,专业名称,班级。"""
    from docx import Document

    doc = Document(docx)
    table = doc.tables[-1]
    rows: list[tuple[str, str, str]] = []
    seen: set[tuple[str, str, str]] = set()
    for row in table.rows:
        cells = [c.text.strip() for c in row.cells]
        for offset in (0, 4):
            if len(cells) >= offset + 3:
                trio = (cells[offset], cells[offset + 1], cells[offset + 2])
                if trio[0] and trio[2] and trio not in seen:
                    seen.add(trio)
                    rows.append(trio)
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    with out_csv.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["学院", "专业名称", "班级"])
        writer.writerows(rows)
    print(f"  班级对照表：{len(rows)} 条 -> {out_csv}")


# 节号 -> (子目录, 文件名)。节号以 extract 时的实测值为准。
FORMS = [
    (12, "附件材料", "01-综合成绩评定表"),
    (13, "附件材料", "02-评分表"),
    (14, "附件材料", "03-指导教师评阅表"),
    (15, "附件材料", "04-交叉评阅表"),
    (16, "附件材料", "05-答辩评分表"),
    (17, "附件材料", "06-答辩记录"),
    (18, "附件材料", "07-任务书"),
    (20, "附件材料", "08-开题报告"),
    (22, "附件材料", "09-中期检查表"),
    (23, "附件材料", "10-指导记录表"),
    (24, "附件材料", "11-承诺书"),
]

# 多节合并的表：(起始节, 结束节, 名称)
MULTI = {
    18: (18, 19),  # 任务书跨两节
    20: (20, 21),  # 开题报告跨两节
}

SAMPLES = [
    (4, "样式参考", "摘要（样式样例）"),
    (5, "样式参考", "Abstract（样式样例）"),
    (6, "样式参考", "目录（样式样例）"),
    (7, "样式参考", "正文（样式样例）"),
    (8, "样式参考", "参考文献（样式样例）"),
    (9, "样式参考", "附录（样式样例）"),
    (11, "样式参考", "附件材料目录"),
]

COVERS = [
    (2, "封面", "封面-毕业论文"),
    (3, "封面", "封面-毕业设计"),
    (10, "封面", "封面-附件材料"),
]


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--manual",
        type=Path,
        default=YEAR_DIR / "原文件" / "手册-2027（Word转换版）.docx",
    )
    ap.add_argument("--out", type=Path, default=YEAR_DIR / "模板")
    args = ap.parse_args(argv)

    if not args.manual.is_file():
        print(f"找不到手册转换版：{args.manual}", file=sys.stderr)
        return 2

    with zipfile.ZipFile(args.manual) as zin:
        body = etree.fromstring(zin.read("word/document.xml")).find(q("body"))
    ranges, final_sect = section_ranges(body)
    print(f"手册共 {len(ranges)} 节")

    for sec, sub, name in COVERS + FORMS + SAMPLES:
        first, last = MULTI.get(sec, (sec, sec))
        out = args.out / sub / f"{name}.docx"
        build_section_docx(args.manual, out, first, last, final_sect)
        print(f"  提取 S{first}-S{last} -> {out.relative_to(args.out)}")

    for _sec, sub, name in COVERS:
        blank_cover_values(args.out / sub / f"{name}.docx")

    # 专业班级对照表（最后一节）
    class_docx = args.out / "参考" / "专业班级对照表.docx"
    last_idx = len(ranges) - 1
    build_section_docx(args.manual, class_docx, last_idx, last_idx, final_sect)
    export_class_table(class_docx, args.out / "参考" / "专业班级对照表.csv")

    index = {
        "source": "附件4：重庆财经学院毕业论文（设计）指导手册（2027届）.doc",
        "covers": [f"封面/{n}.docx" for _, _, n in COVERS],
        "forms": [f"附件材料/{n}.docx" for _, _, n in FORMS],
        "samples": [f"样式参考/{n}.docx" for _, _, n in SAMPLES],
        "reference": ["参考/专业班级对照表.docx", "参考/专业班级对照表.csv"],
    }
    (args.out / "模板索引.json").write_text(
        json.dumps(index, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print("完成")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
