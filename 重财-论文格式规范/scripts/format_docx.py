#!/usr/bin/env python3
"""把已有 .docx 调整为重庆财经学院格式（只改格式，不改内容）。

不生成封面、摘要或固定章节；保留原文档结构与文字。
默认：A4/页边距/宋体+Times/固定26磅/标题层级/三线表/底部页码。
可选：--toc 在已有“目录”标题后插入自动目录域。
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "缺少 python-docx。安装：\n"
        "  python3 -m venv thesis-format/.venv\n"
        "  thesis-format/.venv/bin/pip install -r thesis-format/requirements.txt\n"
    )
    raise SystemExit(2)


FONT_SONG = "宋体"
FONT_HEI = "黑体"
FONT_EN = "Times New Roman"
LINE_PT = 26

CN_NUM = "一二三四五六七八九十百"
H1_RE = re.compile(rf"^[{CN_NUM}]+、")
H2_RE = re.compile(rf"^（[{CN_NUM}]+）")
H3_RE = re.compile(r"^\d+(?:\.\d+)*[.、]\s*")
H4_RE = re.compile(r"^（\d+）")
FIG_RE = re.compile(r"^图\s*\d+(?:[-－]\d+)?")
TAB_RE = re.compile(r"^表\s*\d+(?:[-－]\d+)?")
REF_RE = re.compile(r"^\[\d+\]")
NOTE_RE = re.compile(r"^(注|资料来源|数据来源)[:：]")
KW_CN_RE = re.compile(r"^关键词[:：]")
KW_EN_RE = re.compile(r"^Keywords[:：]", re.I)


def set_run_font(run, cn=FONT_SONG, en=FONT_EN, size=12, bold=None) -> None:
    run.font.name = en
    run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:eastAsia"), cn)
    rfonts.set(qn("w:ascii"), en)
    rfonts.set(qn("w:hAnsi"), en)
    rfonts.set(qn("w:cs"), en)
    rfonts.set(qn("w:hint"), "eastAsia")


def set_para(
    para,
    *,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    first_pt=24,
    hanging_pt=None,
    line_pt=LINE_PT,
    before=0,
    after=0,
) -> None:
    pf = para.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line_pt is None:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.line_spacing = 1.0
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line_pt)
    if hanging_pt is not None:
        pf.left_indent = Pt(hanging_pt)
        pf.first_line_indent = Pt(-hanging_pt)
    else:
        pf.first_line_indent = Pt(first_pt)


def clear_para(para) -> None:
    for child in list(para._element):
        if child.tag != qn("w:pPr"):
            para._element.remove(child)


def set_page(section) -> None:
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.header_distance = Cm(1.5)
    section.footer_distance = Cm(1.75)


def set_page_number(section, fmt="decimal") -> None:
    sect = section._sectPr
    el = sect.find(qn("w:pgNumType"))
    if el is None:
        el = OxmlElement("w:pgNumType")
        sect.append(el)
    el.set(qn("w:fmt"), fmt)
    el.set(qn("w:start"), "1")


def add_field_runs(paragraph, instr_text: str, placeholder: str = "",
                   cn=FONT_SONG, size=12) -> None:
    """用标准的分 run 结构写入 Word 域。"""
    def styled_run():
        run = paragraph.add_run()
        set_run_font(run, cn, FONT_EN, size)
        return run

    r = styled_run()
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    r._element.append(fld)
    r = styled_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instr_text
    r._element.append(instr)
    r = styled_run()
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    r._element.append(sep)
    if placeholder:
        r = styled_run()
        t = OxmlElement("w:t")
        t.text = placeholder
        r._element.append(t)
    r = styled_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r._element.append(end)


def add_page_footer(section) -> None:
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    clear_para(p)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_field_runs(p, " PAGE ", "1", cn=FONT_EN, size=12)


def enable_update_fields(doc) -> None:
    settings = doc.settings.element
    el = settings.find(qn("w:updateFields"))
    if el is None:
        el = OxmlElement("w:updateFields")
        settings.append(el)
    el.set(qn("w:val"), "true")


def add_toc_after(para) -> None:
    from docx.text.paragraph import Paragraph

    new_p = OxmlElement("w:p")
    para._element.addnext(new_p)
    p = Paragraph(new_p, para._parent)
    set_para(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_pt=0)
    add_field_runs(
        p,
        r' TOC \o "1-2" \h \z \u ',
        "打开 Word 时自动更新目录",
        cn=FONT_SONG,
        size=12,
    )


def heading_level(para) -> int | None:
    name = (para.style.name or "").lower().replace(" ", "")
    for n in (1, 2, 3, 4):
        if f"heading{n}" in name or f"标题{n}" in name:
            return n
    text = para.text.strip()
    if not text:
        return None
    if H1_RE.match(text):
        return 1
    if H2_RE.match(text):
        return 2
    if H3_RE.match(text) and len(text) <= 80:
        return 3
    if H4_RE.match(text):
        return 4
    return None


def format_heading(para, level: int, doc) -> None:
    style_name = f"Heading {level}"
    try:
        para.style = doc.styles[style_name]
    except KeyError:
        pass
    if level == 1:
        set_para(para, align=WD_ALIGN_PARAGRAPH.CENTER, first_pt=0, before=12, after=6)
        cn, size, bold = FONT_HEI, 15, True
    elif level == 2:
        set_para(para, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_pt=24)
        cn, size, bold = FONT_HEI, 14, False
    elif level == 3:
        set_para(para, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_pt=24)
        cn, size, bold = FONT_HEI, 12, False
    else:
        set_para(para, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_pt=24)
        cn, size, bold = FONT_SONG, 12, True
    for run in para.runs:
        set_run_font(run, cn, FONT_EN, size, bold)


def format_paragraph(para, doc) -> str:
    text = para.text.strip()
    # 图片段落只改行距，不动对齐/缩进/环绕，避免图片被裁或位置被挪
    if para._p.findall(".//" + qn("w:drawing")) or para._p.findall(".//" + qn("w:pict")):
        pf = para.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.line_spacing = 1.0
        return "image"
    if re.fullmatch(r"摘\s*要", text):
        set_para(para, align=WD_ALIGN_PARAGRAPH.CENTER, first_pt=0, before=12, after=12)
        for run in para.runs:
            set_run_font(run, FONT_HEI, FONT_EN, 15, True)
        return "abstract_cn"
    if text.lower() == "abstract":
        set_para(para, align=WD_ALIGN_PARAGRAPH.CENTER, first_pt=0, before=12, after=12)
        for run in para.runs:
            set_run_font(run, FONT_EN, FONT_EN, 15, True)
        return "abstract_en"
    if re.fullmatch(r"目\s*录", text):
        set_para(para, align=WD_ALIGN_PARAGRAPH.CENTER, first_pt=0, before=12, after=12)
        for run in para.runs:
            set_run_font(run, FONT_HEI, FONT_EN, 15, True)
        return "toc_title"
    if text == "参考文献":
        set_para(para, align=WD_ALIGN_PARAGRAPH.CENTER, first_pt=0, before=12, after=6)
        try:
            para.style = doc.styles["Heading 1"]
        except KeyError:
            pass
        for run in para.runs:
            set_run_font(run, FONT_HEI, FONT_EN, 15, True)
        return "ref_title"
    if re.match(r"^附录[A-Z]", text):
        set_para(para, align=WD_ALIGN_PARAGRAPH.CENTER, first_pt=0, before=12, after=6)
        try:
            para.style = doc.styles["Heading 1"]
        except KeyError:
            pass
        for run in para.runs:
            set_run_font(run, FONT_HEI, FONT_EN, 15, True)
        return "appendix_title"

    level = heading_level(para)
    if level:
        format_heading(para, level, doc)
        return f"h{level}"

    if KW_CN_RE.match(text):
        set_para(para, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_pt=0)
        for i, run in enumerate(para.runs):
            cn = FONT_HEI if i == 0 and KW_CN_RE.match(run.text.strip()) else FONT_SONG
            set_run_font(run, cn, FONT_EN, 12, None)
        return "keywords_cn"
    if KW_EN_RE.match(text):
        set_para(para, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_pt=0)
        for run in para.runs:
            set_run_font(run, FONT_EN, FONT_EN, 12, None)
        return "keywords_en"

    if FIG_RE.match(text) or TAB_RE.match(text):
        set_para(para, align=WD_ALIGN_PARAGRAPH.CENTER, first_pt=0)
        for run in para.runs:
            set_run_font(run, FONT_HEI, FONT_EN, 10.5, False)
        return "caption"

    if NOTE_RE.match(text):
        set_para(para, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_pt=24, line_pt=None)
        for run in para.runs:
            set_run_font(run, FONT_SONG, FONT_EN, 9, None)
        return "note"

    if REF_RE.match(text):
        set_para(para, align=WD_ALIGN_PARAGRAPH.JUSTIFY, hanging_pt=21)
        for run in para.runs:
            set_run_font(run, FONT_SONG, FONT_EN, 10.5, None)
        return "reference"

    set_para(para, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_pt=24)
    for run in para.runs:
        set_run_font(run, FONT_SONG, FONT_EN, 12, None)
    return "body"


def set_cell_borders(cell, header_row: bool, last_row: bool) -> None:
    """三线表：首行顶线、表头底线、末行底线；左右与内部横线全部去掉。

    单元格边框优先级高于表格边框，所以外框必须逐格写，否则会被 nil 覆盖。
    """
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    old = tcpr.find(qn("w:tcBorders"))
    if old is not None:
        tcpr.remove(old)
    # 去掉原有底纹，三线表表头不铺色
    for shd in tcpr.findall(qn("w:shd")):
        tcpr.remove(shd)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        if edge == "top":
            on = header_row
        elif edge == "bottom":
            on = header_row or last_row
        else:
            on = False
        if on:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "8")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")
        else:
            el.set(qn("w:val"), "nil")
        borders.append(el)
    tcpr.append(borders)


def format_table(table) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # 跨页续表时重复表头
    if table.rows:
        trpr = table.rows[0]._tr.get_or_add_trPr()
        if trpr.find(qn("w:tblHeader")) is None:
            header = OxmlElement("w:tblHeader")
            header.set(qn("w:val"), "true")
            trpr.append(header)
    tblpr = table._tbl.tblPr
    old = tblpr.find(qn("w:tblBorders"))
    if old is not None:
        tblpr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge, sz in (("top", "12"), ("bottom", "12"),
                     ("left", "nil"), ("right", "nil"),
                     ("insideH", "nil"), ("insideV", "nil")):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil" if sz == "nil" else "single")
        if sz != "nil":
            el.set(qn("w:sz"), sz)
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")
        borders.append(el)
    tblpr.append(borders)

    n_rows = len(table.rows)
    for ri, row in enumerate(table.rows):
        header = ri == 0
        last = ri == n_rows - 1
        for cell in row.cells:
            set_cell_borders(cell, header, last)
            for p in cell.paragraphs:
                p.paragraph_format.first_line_indent = Pt(0)
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p.paragraph_format.line_spacing = 1.0
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if header else WD_ALIGN_PARAGRAPH.LEFT
                for run in p.runs:
                    set_run_font(run, FONT_SONG, FONT_EN, 10.5, True if header else None)


def blacken_toc(doc) -> None:
    """把 TOC 标题与条目改成黑色。

    pandoc 生成的目录在 w:sdt 里，不在 doc.paragraphs 的顶层列表中；
    Word 默认还会用 Hyperlink / TOC 样式把目录染成蓝色。这里直接遍历
    document.xml 的所有段落，命中 TOC 样式后覆盖 run 颜色。
    """
    for p in doc.element.body.iter(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        style = ppr.find(qn("w:pStyle")) if ppr is not None else None
        style_id = style.get(qn("w:val"), "") if style is not None else ""
        if not style_id.lower().startswith("toc"):
            continue
        for run in p.iter(qn("w:r")):
            rpr = run.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                run.insert(0, rpr)
            rstyle = rpr.find(qn("w:rStyle"))
            if rstyle is not None and rstyle.get(qn("w:val")) == "Hyperlink":
                rpr.remove(rstyle)
            color = rpr.find(qn("w:color"))
            if color is None:
                color = OxmlElement("w:color")
                rpr.append(color)
            for attr in ("themeColor", "themeShade", "themeTint"):
                color.attrib.pop(qn(f"w:{attr}"), None)
            color.set(qn("w:val"), "000000")


def find_toc_anchor(doc):
    for p in doc.paragraphs:
        if p.text.strip() == "[[TOC]]" or re.fullmatch(r"目\s*录", p.text.strip()):
            return p
    return None


def unique_output(path: Path) -> Path:
    if not path.exists():
        return path
    for i in range(2, 1000):
        cand = path.with_name(f"{path.stem}_{i}{path.suffix}")
        if not cand.exists():
            return cand
    raise RuntimeError(f"无法生成不冲突的输出名：{path}")


def format_docx(src: Path, dst: Path, *, toc=False, page_numbers=True,
                three_line_tables=True, page_format="decimal") -> dict:
    doc = Document(src)
    stats = {"body": 0, "headings": 0, "tables": 0, "references": 0}

    try:
        normal = doc.styles["Normal"]
        normal.font.name = FONT_EN
        normal.font.size = Pt(12)
        rpr = normal.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        rfonts.set(qn("w:eastAsia"), FONT_SONG)
        rfonts.set(qn("w:ascii"), FONT_EN)
        rfonts.set(qn("w:hAnsi"), FONT_EN)
        pf = normal.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Pt(24)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(LINE_PT)
    except KeyError:
        pass

    for section in doc.sections:
        set_page(section)
        if page_numbers:
            set_page_number(section, page_format)
            add_page_footer(section)

    for p in doc.paragraphs:
        kind = format_paragraph(p, doc)
        if kind.startswith("h"):
            stats["headings"] += 1
        elif kind == "reference":
            stats["references"] += 1
        else:
            stats["body"] += 1

    if three_line_tables:
        for table in doc.tables:
            format_table(table)
            stats["tables"] += 1

    blacken_toc(doc)

    if toc:
        anchor = find_toc_anchor(doc)
        if anchor is None:
            print("WARN 未找到“目录”标题或 [[TOC]] 占位符，未插入目录。")
        else:
            if anchor.text.strip() == "[[TOC]]":
                clear_para(anchor)
                add_toc_after(anchor)
                anchor._element.getparent().remove(anchor._element)
            else:
                add_toc_after(anchor)

    enable_update_fields(doc)
    dst = unique_output(dst)
    doc.save(dst)
    stats["output"] = str(dst)
    return stats


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="按重庆财经学院格式调整已有 docx（只改格式，不改内容）")
    ap.add_argument("input", type=Path)
    ap.add_argument("-o", "--output", type=Path)
    ap.add_argument("--toc", action="store_true", help="在已有“目录”标题后插入自动目录域")
    ap.add_argument("--no-page-numbers", action="store_true")
    ap.add_argument("--no-three-line-tables", action="store_true")
    ap.add_argument("--page-number-format", choices=("decimal", "upperRoman"), default="decimal")
    args = ap.parse_args(argv)

    if not args.input.is_file():
        print(f"输入不存在：{args.input}", file=sys.stderr)
        return 2
    dst = args.output or args.input.with_name(args.input.stem + "_cfec.docx")
    stats = format_docx(
        args.input,
        dst,
        toc=args.toc,
        page_numbers=not args.no_page_numbers,
        three_line_tables=not args.no_three_line_tables,
        page_format=args.page_number_format,
    )
    print(f"OK   {stats['output']}")
    print(
        f"     标题 {stats['headings']} / 正文段 {stats['body']} / "
        f"表格 {stats['tables']} / 参考文献 {stats['references']}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
