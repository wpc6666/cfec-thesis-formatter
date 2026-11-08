#!/usr/bin/env python3
"""2027 届重财论文：封面、前置部分、正文、附件表的生成与填充。

只做格式与装配，不改写论文内容。格式依据：
Skill 内 assets/format.yaml 与《指导手册（2027届）》。
"""

from __future__ import annotations

import copy
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import yaml
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

HERE = Path(__file__).resolve().parent
SKILL_DIR = HERE.parent.parent
ASSETS = SKILL_DIR / "assets"
YEAR_DIR = ASSETS / "2027"
sys.path.insert(0, str(HERE))

from docx_merge import merge_docx  # noqa: E402

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

SONG = "宋体"
HEI = "黑体"
TNR = "Times New Roman"

CN_DIGITS = "一二三四五六七八九十"


# ---------------------------------------------------------------- 基础格式


def cn_number(n: int) -> str:
    """1 -> 一，11 -> 十一，21 -> 二十一。"""
    if n <= 0:
        return str(n)
    if n <= 10:
        return CN_DIGITS[n - 1]
    if n < 20:
        return "十" + CN_DIGITS[n - 11]
    tens, ones = divmod(n, 10)
    return CN_DIGITS[tens - 1] + "十" + (CN_DIGITS[ones - 1] if ones else "")


def set_run_font(
    run, cn: str = SONG, en: str = TNR, size: float = 12, bold: bool | None = False
) -> None:
    run.font.name = en
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.italic = False
    run.underline = False
    if bold is not None:
        run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), en)
    rfonts.set(qn("w:hAnsi"), en)
    rfonts.set(qn("w:eastAsia"), cn)
    rfonts.set(qn("w:cs"), en)


def _drop(ppr, tag: str) -> None:
    for el in ppr.findall(qn(f"w:{tag}")):
        ppr.remove(el)


def set_paragraph(
    para,
    *,
    align: str = "both",
    first_line_chars: int | None = None,
    left_chars: int | None = None,
    hanging_chars: int | None = None,
    line_pt: float | None = 26,
    before_pt: float = 0,
    after_pt: float = 0,
    outline: int | None = None,
    keep_with_next: bool = False,
) -> None:
    pf = para.paragraph_format
    para.alignment = ALIGN.get(align, WD_ALIGN_PARAGRAPH.JUSTIFY)
    pf.space_before = Pt(before_pt)
    pf.space_after = Pt(after_pt)
    if line_pt is None:
        pf.line_spacing = None
    else:
        pf.line_spacing = Pt(line_pt)
    pf.keep_with_next = keep_with_next
    ppr = para._p.get_or_add_pPr()
    # 先用 python-docx 的 setter 生成 w:ind（保证 OOXML 顺序），再补字符单位
    pf.first_line_indent = Pt(0)
    pf.left_indent = Pt(0)
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    for attr in (
        "firstLine",
        "firstLineChars",
        "left",
        "leftChars",
        "hanging",
        "hangingChars",
    ):
        ind.attrib.pop(qn(f"w:{attr}"), None)
    if first_line_chars:
        ind.set(qn("w:firstLineChars"), str(first_line_chars * 100))
        ind.set(qn("w:firstLine"), str(int(first_line_chars * 240)))
    elif left_chars is not None:
        ind.set(qn("w:leftChars"), str(left_chars * 100))
        ind.set(qn("w:left"), str(int(left_chars * 240)))
        ind.set(qn("w:firstLine"), "0")
    elif hanging_chars:
        ind.set(qn("w:hangingChars"), str(hanging_chars * 100))
        ind.set(qn("w:hanging"), str(int(hanging_chars * 210)))
        ind.set(qn("w:firstLine"), "0")
    else:
        ind.set(qn("w:firstLine"), "0")
        ind.set(qn("w:firstLineChars"), "0")

    _drop(ppr, "outlineLvl")
    if outline is not None:
        lvl = OxmlElement("w:outlineLvl")
        lvl.set(qn("w:val"), str(outline))
        ppr.append(lvl)
    if line_pt is not None:
        spacing = ppr.find(qn("w:spacing"))
        if spacing is not None:
            spacing.set(qn("w:line"), str(int(line_pt * 20)))
            spacing.set(qn("w:lineRule"), "exact")


def p_text(el) -> str:
    """只取 w:t 的文本。python-docx 的 CT_P.text / itertext 会把同一文字按层级重复计数。"""
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def no_auto_space(para) -> None:
    """关闭中英文自动间距，避免“附录A”“图3-1”被渲染成“附录 A”“图 3-1”。"""
    ppr = para._p.get_or_add_pPr()
    for tag in ("autoSpaceDE", "autoSpaceDN"):
        for old in ppr.findall(qn(f"w:{tag}")):
            ppr.remove(old)
        el = OxmlElement(f"w:{tag}")
        el.set(qn("w:val"), "0")
        ppr.insert(0, el)


def clear_paragraph(para) -> None:
    for child in list(para._p):
        if child.tag in (
            qn("w:r"),
            qn("w:hyperlink"),
            qn("w:bookmarkStart"),
            qn("w:bookmarkEnd"),
        ):
            para._p.remove(child)


def write_paragraph(para, text: str, **fmt) -> None:
    clear_paragraph(para)
    run = para.add_run(text)
    set_run_font(
        run,
        cn=fmt.pop("cn", SONG),
        en=fmt.pop("en", TNR),
        size=fmt.pop("size", 12),
        bold=fmt.pop("bold", False),
    )
    set_paragraph(para, **fmt)


def set_page(section) -> None:
    section.page_width = Pt(595.3)  # 21 cm
    section.page_height = Pt(841.9)  # 29.7 cm
    section.left_margin = Pt(79.4)  # 2.8 cm
    section.right_margin = Pt(70.9)  # 2.5 cm
    section.top_margin = Pt(70.9)
    section.bottom_margin = Pt(70.9)
    section.header_distance = Pt(42.5)  # 1.5 cm
    section.footer_distance = Pt(49.6)  # 1.75 cm


def set_page_numbering(section, fmt: str | None, start: int | None = 1) -> None:
    sect = section._sectPr
    for old in sect.findall(qn("w:pgNumType")):
        sect.remove(old)
    if fmt is None:
        return
    pg = OxmlElement("w:pgNumType")
    pg.set(qn("w:fmt"), fmt)
    if start is not None:
        pg.set(qn("w:start"), str(start))
    cols = sect.find(qn("w:cols"))
    if cols is not None:
        cols.addprevious(pg)
    else:
        sect.append(pg)


def add_page_footer(section) -> None:
    section.footer.is_linked_to_previous = False
    footer = section.footer
    para = footer.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for child in list(para._p):
        if child.tag == qn("w:r"):
            para._p.remove(child)
    run = para.add_run()
    set_run_font(run, cn=TNR, en=TNR, size=12, bold=False)
    f1 = OxmlElement("w:fldChar")
    f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = " PAGE "
    f2 = OxmlElement("w:fldChar")
    f2.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t")
    t.text = "1"
    f3 = OxmlElement("w:fldChar")
    f3.set(qn("w:fldCharType"), "end")
    for el in (f1, it, f2, t, f3):
        run._r.append(copy.deepcopy(el))


def enable_update_fields(doc) -> None:
    settings = doc.settings.element
    for old in settings.findall(qn("w:updateFields")):
        settings.remove(old)
    el = OxmlElement("w:updateFields")
    el.set(qn("w:val"), "true")
    settings.append(el)


def add_toc(doc) -> None:
    para = doc.add_paragraph()
    set_paragraph(para, align="left", line_pt=26)
    run = para.add_run()
    f1 = OxmlElement("w:fldChar")
    f1.set(qn("w:fldCharType"), "begin")
    f1.set(qn("w:dirty"), "true")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = ' TOC \\o "1-2" \\z \\u '
    f2 = OxmlElement("w:fldChar")
    f2.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t")
    t.text = "打开文档后按 Ctrl+A、F9 更新目录"
    f3 = OxmlElement("w:fldChar")
    f3.set(qn("w:fldCharType"), "end")
    for el in (f1, it, f2, t, f3):
        run._r.append(copy.deepcopy(el))


# ---------------------------------------------------------------- Markdown


def read_markdown(path: Path) -> tuple[dict, str]:
    text = path.read_text(encoding="utf-8")
    meta: dict = {}
    body = text
    if text.startswith("---"):
        parts = text.split("---", 2)
        if len(parts) == 3:
            meta = yaml.safe_load(parts[1]) or {}
            body = parts[2].lstrip("\n")
    return meta, body


def normalize_keywords(raw: str, *, english: bool = False) -> str:
    items = [x.strip() for x in re.split(r"[；;，,]", str(raw or "")) if x.strip()]
    items = [x.rstrip("。；;，,、 ") for x in items]
    return ("; " if english else "；").join(items)


def paragraphs_of(text: str) -> list[str]:
    chunks = [x.strip() for x in re.split(r"\n\s*\n", str(text or "").strip())]
    return [x for x in chunks if x]


# ---------------------------------------------------------------- 封面


def _set_cell(
    cell, text: str, *, size: float = 14, bold: bool = False, align: str = "center"
) -> None:
    para = cell.paragraphs[0]
    for child in list(para._p):
        if child.tag in (qn("w:r"), qn("w:hyperlink")):
            para._p.remove(child)
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    run = para.add_run(text)
    set_run_font(run, cn=SONG, en=TNR, size=size, bold=bold)
    set_paragraph(para, align=align, line_pt=26)


def _table_pairs(table) -> list[tuple[str, object]]:
    pairs = []
    for row in table.rows:
        seen = set()
        cells = []
        for c in row.cells:
            if id(c._tc) in seen:
                continue
            seen.add(id(c._tc))
            cells.append(c)
        for i in range(len(cells) - 1):
            label = cells[i].text.strip().rstrip("：:")
            if label in {
                "题目",
                "学院",
                "年级",
                "专业",
                "班级",
                "学号",
                "姓名",
                "指导教师",
            }:
                pairs.append((label, cells[i + 1]))
    return pairs


def build_cover(template: Path, info: dict, out: Path) -> Path:
    doc = Document(template)
    values = {
        "题目": info.get("title", ""),
        "学院": info.get("college", ""),
        "年级": info.get("grade", ""),
        "专业": info.get("major", ""),
        "班级": info.get("class_name", ""),
        "学号": info.get("student_id", ""),
        "姓名": info.get("name", ""),
        "指导教师": info.get("advisor", ""),
    }
    for table in doc.tables:
        for label, cell in _table_pairs(table):
            if label == "题目":
                _set_cell(cell, values["题目"], size=14, bold=True)
            else:
                _set_cell(cell, values[label], size=14, bold=False)
    date_text = info.get("date") or f"{info.get('year', '2027')}年5月"
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if re.fullmatch(r"\d{4}年\d{1,2}月", cell.text.strip()):
                    _set_cell(cell, date_text, size=15, bold=False)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return out


# ---------------------------------------------------------------- 前置部分


def _new_doc() -> Document:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = TNR
    style.font.size = Pt(12)
    set_page(doc.sections[0])
    return doc


def _title_block(doc, text: str, *, cn: bool = True) -> None:
    para = doc.add_paragraph()
    run = para.add_run(text)
    set_run_font(run, cn=HEI if cn else TNR, en=TNR, size=15, bold=True)
    set_paragraph(para, align="center", line_pt=26, outline=None)


def _abstract_page(doc, title: str, body: str, keywords: str, *, cn: bool) -> None:
    _title_block(doc, title, cn=cn)
    for text in paragraphs_of(body):
        text = re.sub(r"\s*\n\s*", "" if cn else " ", text)
        para = doc.add_paragraph()
        run = para.add_run(text)
        set_run_font(run, cn=SONG if cn else TNR, en=TNR, size=12, bold=False)
        _format_runs(para, SONG, TNR, 12, False)
        set_paragraph(para, align="both", first_line_chars=2, line_pt=26)
    doc.add_paragraph()
    para = doc.add_paragraph()
    label = para.add_run("关键词：" if cn else "Keywords: ")
    set_run_font(label, cn=HEI if cn else TNR, en=TNR, size=12, bold=not cn)
    value = para.add_run(keywords)
    set_run_font(value, cn=SONG if cn else TNR, en=TNR, size=12, bold=False)
    set_paragraph(para, align="both", line_pt=26)


def build_front(info: dict, out: Path, *, include_toc: bool = True) -> Path:
    doc = _new_doc()
    section = doc.sections[0]
    add_page_footer(section)
    set_page_numbering(section, "upperRoman", 1)

    _abstract_page(
        doc,
        "摘  要",
        info.get("abstract_cn", ""),
        normalize_keywords(info.get("keywords_cn", "")),
        cn=True,
    )
    doc.add_page_break()
    _abstract_page(
        doc,
        "Abstract",
        info.get("abstract_en", ""),
        normalize_keywords(info.get("keywords_en", ""), english=True),
        cn=False,
    )
    if include_toc:
        doc.add_page_break()
        _title_block(doc, "目  录")
        doc.add_paragraph()
        add_toc(doc)
    enable_update_fields(doc)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return out


def build_single_abstract(info: dict, out: Path, *, cn: bool) -> Path:
    doc = _new_doc()
    section = doc.sections[0]
    add_page_footer(section)
    set_page_numbering(section, "upperRoman", 1)
    if cn:
        _abstract_page(
            doc,
            "摘  要",
            info.get("abstract_cn", ""),
            normalize_keywords(info.get("keywords_cn", "")),
            cn=True,
        )
    else:
        _abstract_page(
            doc,
            "Abstract",
            info.get("abstract_en", ""),
            normalize_keywords(info.get("keywords_en", ""), english=True),
            cn=False,
        )
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return out


# ---------------------------------------------------------------- 正文


def _reference_doc() -> Path:
    """pandoc 参考样式，随 skill 一起分发，不依赖外部工具目录。"""
    ref = ASSETS / "thesis-reference.docx"
    if not ref.is_file():
        raise RuntimeError(f"缺少参考样式文件：{ref}")
    return ref


def pandoc_markdown(body_md: str, out_docx: Path) -> None:
    ref = _reference_doc()
    with tempfile.TemporaryDirectory(prefix="cfec2027-") as td:
        md = Path(td) / "body.md"
        md.write_text(body_md, encoding="utf-8")
        subprocess.run(
            [
                "pandoc",
                str(md),
                "-o",
                str(out_docx),
                "-f",
                "markdown+tex_math_single_backslash",
                "--reference-doc",
                str(ref),
            ],
            check=True,
        )


SPECIAL_H1 = re.compile(r"^(参考文献|致谢|附录\s*[A-Z])$")

HEADING_SPECS = {
    "Heading 1": (HEI, 15, False),
    "Heading 2": (HEI, 14, False),
    "Heading 3": (HEI, 12, False),
    "Heading 4": (SONG, 12, True),
}


def set_heading_styles(doc) -> None:
    """把标题字体放进样式，而不是 run 直接格式。

    Word 更新目录时会把标题 run 的直接字体复制到目录条目；去掉直接字体后，
    目录才会使用 TOC1/TOC2 样式（宋体、小四）。
    """
    for name, (cn, size, bold) in HEADING_SPECS.items():
        try:
            style = doc.styles[name]
        except KeyError:
            continue
        style.font.name = TNR
        style.font.size = Pt(size)
        style.font.bold = bold
        style.font.color.rgb = RGBColor(0, 0, 0)
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            rfonts.attrib.pop(qn(f"w:{attr}"), None)
        rfonts.set(qn("w:ascii"), TNR)
        rfonts.set(qn("w:hAnsi"), TNR)
        rfonts.set(qn("w:eastAsia"), cn)


def strip_heading_run_fonts(doc) -> None:
    for para in doc.paragraphs:
        if para.style.name not in HEADING_SPECS:
            continue
        for run in para.runs:
            rpr = run._r.find(qn("w:rPr"))
            if rpr is None:
                continue
            rfonts = rpr.find(qn("w:rFonts"))
            if rfonts is not None:
                rpr.remove(rfonts)
CAPTION = re.compile(r"^(图|表)\s*(\d+[-－]\d+)\s*(.*)$")


def _heading_level(para) -> int | None:
    name = (para.style.name or "").lower()
    m = re.match(r"heading\s*(\d)", name)
    if m:
        return int(m.group(1))
    ppr = para._p.find(qn("w:pPr"))
    if ppr is not None:
        lvl = ppr.find(qn("w:outlineLvl"))
        if lvl is not None:
            return int(lvl.get(qn("w:val"))) + 1
    return None


def _format_runs(para, cn: str, en: str, size: float, bold: bool) -> None:
    for run in para.runs:
        set_run_font(run, cn=cn, en=en, size=size, bold=bold)


def _page_break_before(para) -> None:
    prev = para._p.getprevious()
    if prev is not None:
        for br in prev.iter(qn("w:br")):
            if br.get(qn("w:type")) == "page":
                return
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    para._p.addprevious(p)


def _insert_empty_before(para) -> None:
    prev = para._p.getprevious()
    if prev is not None and prev.tag == qn("w:p"):
        text = p_text(prev).strip()
        if not text:
            return
    new_p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    rpr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "24")
    rpr.append(sz)
    ppr.append(rpr)
    new_p.append(ppr)
    para._p.addprevious(new_p)


def _three_line_table(table) -> None:
    tbl_pr = table._tbl.tblPr
    for old in tbl_pr.findall(qn("w:tblBorders")):
        tbl_pr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge, sz in (
        ("top", 12),
        ("bottom", 12),
        ("left", 0),
        ("right", 0),
        ("insideH", 0),
        ("insideV", 0),
    ):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single" if sz else "nil")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tbl_pr.append(borders)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 表头下横线
    if table.rows:
        for cell in table.rows[0].cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            for old in tc_pr.findall(qn("w:tcBorders")):
                tc_pr.remove(old)
            tcb = OxmlElement("w:tcBorders")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "6")
            bottom.set(qn("w:space"), "0")
            bottom.set(qn("w:color"), "000000")
            tcb.append(bottom)
            tc_pr.append(tcb)

    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                set_paragraph(
                    para, align="center", line_pt=None, before_pt=0, after_pt=0
                )
                _format_runs(para, SONG, TNR, 10.5, False)


def format_body(docx_path: Path) -> None:
    doc = Document(docx_path)
    section = doc.sections[0]
    set_page(section)
    add_page_footer(section)
    set_page_numbering(section, "decimal", 1)

    chapter = 0
    sub = 0
    third = 0
    fourth = 0
    in_refs = False
    in_appendix = False
    appendix_title_pending = False
    seen_content = False

    for para in doc.paragraphs:
        text = para.text.strip()
        level = _heading_level(para)

        if level == 1 and text:
            if SPECIAL_H1.match(text):
                in_refs = text == "参考文献"
                m = re.match(r"^附录\s*([A-Z])", text)
                in_appendix = bool(m)
                title = text if not m else f"附录{m.group(1)}"
                write_paragraph(
                    para,
                    title,
                    cn=HEI,
                    en=TNR,
                    size=15,
                    bold=False,
                    align="center",
                    line_pt=26,
                    outline=0,
                    keep_with_next=True,
                )
                if in_refs or in_appendix:
                    _page_break_before(para)
                if in_appendix:
                    no_auto_space(para)
                    appendix_title_pending = True
                if seen_content and not (in_refs or in_appendix):
                    _insert_empty_before(para)
            else:
                chapter += 1
                sub = third = fourth = 0
                in_refs = False
                in_appendix = False
                clean = re.sub(r"^[一二三四五六七八九十百]+[、．.]\s*", "", text)
                numbered = f"{cn_number(chapter)}、{clean}"
                write_paragraph(
                    para,
                    numbered,
                    cn=HEI,
                    en=TNR,
                    size=15,
                    bold=False,
                    align="center",
                    line_pt=26,
                    outline=0,
                    keep_with_next=True,
                )
                if seen_content:
                    _insert_empty_before(para)
            seen_content = True
            continue

        if appendix_title_pending and text and level is None:
            appendix_title_pending = False
            write_paragraph(
                para,
                text,
                cn=HEI,
                en=TNR,
                size=14,
                bold=False,
                align="center",
                line_pt=26,
                outline=None,
                keep_with_next=True,
            )
            seen_content = True
            continue
        if text:
            appendix_title_pending = False

        if level == 2 and text:
            sub += 1
            third = fourth = 0
            if in_appendix:
                clean = re.sub(r"^（[一二三四五六七八九十]+）\s*", "", text)
                write_paragraph(
                    para,
                    clean,
                    cn=HEI,
                    en=TNR,
                    size=14,
                    bold=False,
                    align="center",
                    line_pt=26,
                    outline=None,
                    keep_with_next=True,
                )
            else:
                clean = re.sub(r"^（[一二三四五六七八九十]+）\s*", "", text)
                write_paragraph(
                    para,
                    f"（{cn_number(sub)}）{clean}",
                    cn=HEI,
                    en=TNR,
                    size=14,
                    bold=False,
                    align="both",
                    first_line_chars=2,
                    line_pt=26,
                    outline=1,
                    keep_with_next=True,
                )
            seen_content = True
            continue

        if level == 3 and text:
            third += 1
            fourth = 0
            clean = re.sub(r"^\d+[.．、]\s*", "", text)
            write_paragraph(
                para,
                f"{third}.{clean}",
                cn=HEI,
                en=TNR,
                size=12,
                bold=False,
                align="both",
                first_line_chars=2,
                line_pt=26,
                outline=2,
                keep_with_next=True,
            )
            seen_content = True
            continue

        if level == 4 and text:
            fourth += 1
            clean = re.sub(r"^（\d+）\s*", "", text)
            write_paragraph(
                para,
                f"（{fourth}）{clean}",
                cn=SONG,
                en=TNR,
                size=12,
                bold=True,
                align="both",
                first_line_chars=2,
                line_pt=26,
                outline=None,
                keep_with_next=True,
            )
            seen_content = True
            continue

        if para._p.findall(".//" + qn("w:drawing")):
            set_paragraph(para, align="center", line_pt=None)
            continue

        if not text:
            continue

        if "caption" in (para.style.name or "").lower():
            m = CAPTION.match(text)
            caption = f"{m.group(1)}{m.group(2)}  {m.group(3)}".rstrip() if m else text
            write_paragraph(para, caption, cn=HEI, en=TNR, size=10.5, bold=False,
                            align="center", line_pt=None, before_pt=0, after_pt=0)
            no_auto_space(para)
            continue

        m = CAPTION.match(text)
        if m:
            label, num, title = m.groups()
            write_paragraph(
                para,
                f"{label}{num}  {title}".rstrip(),
                cn=HEI,
                en=TNR,
                size=10.5,
                bold=False,
                align="center",
                line_pt=None,
                before_pt=0,
                after_pt=0,
            )
            no_auto_space(para)
            continue

        if re.match(r"^(数据来源|资料来源|注)\s*[：:]", text):
            write_paragraph(para, text, cn=SONG, en=TNR, size=9, bold=False,
                            align="both", first_line_chars=2, line_pt=None)
            continue

        if in_refs:
            clean = re.sub(r"^\[(\d+)\]\s*", r"[\1]\t", text)
            write_paragraph(
                para,
                clean,
                cn=SONG,
                en=TNR,
                size=10.5,
                bold=False,
                align="both",
                hanging_chars=2,
                line_pt=26,
            )
            continue

        _format_runs(para, SONG, TNR, 12, False)
        set_paragraph(para, align="both", first_line_chars=2, line_pt=26)

    for table in doc.tables:
        _three_line_table(table)

    set_heading_styles(doc)
    strip_heading_run_fonts(doc)
    enable_update_fields(doc)
    doc.save(docx_path)


def build_body(body_md: str, out: Path) -> Path:
    pandoc_markdown(body_md, out)
    format_body(out)
    return out


def split_body(body_docx: Path, out_dir: Path) -> dict[str, Path]:
    """按“参考文献 / 附录”切出正文、参考文献、附录三个文件。"""
    doc = Document(body_docx)
    body = doc.element.body
    children = list(body)
    anchors: dict[str, int] = {}
    for i, ch in enumerate(children):
        if ch.tag != qn("w:p"):
            continue
        text = p_text(ch).strip()
        if text == "参考文献" and "refs" not in anchors:
            anchors["refs"] = i
        elif re.match(r"^附录\s*[A-Z]$", text) and "appendix" not in anchors:
            anchors["appendix"] = i
    out_dir.mkdir(parents=True, exist_ok=True)
    names = ["正文", "参考文献", "附录"]
    starts = [
        0,
        anchors.get("refs", len(children)),
        anchors.get("appendix", len(children)),
    ]
    ends = [starts[1], starts[2], len(children)]
    out: dict[str, Path] = {}
    for index, name in enumerate(names):
        start, end = starts[index], ends[index]
        if start >= end:
            continue
        target = out_dir / f"{name}.docx"
        shutil.copy2(body_docx, target)
        d = Document(target)
        b = d.element.body
        for i, ch in enumerate(list(b)):
            if i < start or i >= end:
                b.remove(ch)
        d.save(target)
        out[name] = target
    return out


# ---------------------------------------------------------------- 附件表


FORM_LABELS = {"姓名", "学号", "专业", "题目", "学院", "年级", "班级"}


def fill_form(template: Path, info: dict, out: Path) -> Path:
    values = {
        "姓名": info.get("name", ""),
        "学号": info.get("student_id", ""),
        "专业": info.get("major", ""),
        "题目": info.get("title", ""),
        "学院": info.get("college", ""),
        "年级": info.get("grade", ""),
        "班级": info.get("class_name", ""),
    }
    doc = Document(template)
    for table in doc.tables:
        for row in table.rows:
            seen = set()
            cells = []
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                cells.append(c)
            for i, cell in enumerate(cells[:-1]):
                label = cell.text.strip().rstrip("：:")
                if label in FORM_LABELS and not cells[i + 1].text.strip():
                    target = cells[i + 1]
                    align = "center"
                    if label != "题目" and target.width and target.width.cm > 6:
                        align = "left"
                    _set_cell(
                        target, values.get(label, ""), size=14, bold=False, align=align
                    )
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return out


# ---------------------------------------------------------------- 完整论文


def fix_toc_styles(docx_path: Path) -> None:
    """目录条目：宋体/TNR 小四、黑色、点状前导符、页码右对齐。"""
    doc = Document(docx_path)
    for name, left_chars in (("toc 1", 0), ("toc 2", 2)):
        try:
            style = doc.styles[name]
        except KeyError:
            continue
        style.font.name = TNR
        style.font.size = Pt(12)
        style.font.bold = False
        style.font.color.rgb = RGBColor(0, 0, 0)
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            rfonts.attrib.pop(qn(f"w:{attr}"), None)
        rfonts.set(qn("w:ascii"), TNR)
        rfonts.set(qn("w:hAnsi"), TNR)
        rfonts.set(qn("w:eastAsia"), SONG)
        pf = style.paragraph_format
        pf.left_indent = Pt(12 * left_chars)
        pf.first_line_indent = Pt(0)
        pf.line_spacing = Pt(26)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.tab_stops.add_tab_stop(Pt(415), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        if left_chars:
            ppr = style.element.find(qn("w:pPr"))
            ind = ppr.find(qn("w:ind")) if ppr is not None else None
            if ind is not None:
                ind.set(qn("w:leftChars"), str(left_chars * 100))
    doc.save(docx_path)


def build_full(cover: Path, front: Path, body: Path, out: Path) -> Path:
    tmp1 = out.with_name("_merge1.docx")
    tmp2 = out.with_name("_merge2.docx")
    merge_docx(cover, front, tmp1)
    merge_docx(tmp1, body, tmp2)
    shutil.move(tmp2, out)
    merged = Document(out)
    set_heading_styles(merged)
    strip_heading_run_fonts(merged)
    merged.save(out)
    fix_toc_styles(out)
    tmp1.unlink(missing_ok=True)
    return out
