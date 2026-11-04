#!/usr/bin/env python3
"""从 Markdown 生成符合重庆财经学院格式的 docx。

流程：
  1. 读取 YAML 元数据，拼出封面 / 中英文摘要 / 目录 / 正文；
  2. 用 pandoc 默认 reference.docx 生成带正确样式的参考文档；
  3. pandoc 生成 docx；
  4. 后处理：页面设置、TOC 域、罗马/阿拉伯页码、页脚页码。
"""
from __future__ import annotations

import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

import yaml
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ref_zh  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def rid(tag: str) -> str:
    return f"{{{R}}}{tag}"


def get_or_make(parent: etree._Element, tag: str) -> etree._Element:
    el = parent.find(q(tag))
    if el is None:
        el = etree.SubElement(parent, q(tag))
    return el


def set_rpr(
    rpr: etree._Element,
    *,
    font_cn: str = "宋体",
    font_en: str = "Times New Roman",
    size_pt: float = 12,
    bold: bool = False,
) -> None:
    for child in list(rpr):
        if child.tag in {q("rFonts"), q("sz"), q("szCs"), q("b"), q("bCs"), q("lang")}:
            rpr.remove(child)
    rfonts = etree.SubElement(rpr, q("rFonts"))
    rfonts.set(q("ascii"), font_en)
    rfonts.set(q("hAnsi"), font_en)
    rfonts.set(q("eastAsia"), font_cn)
    rfonts.set(q("cs"), font_en)
    rfonts.set(q("hint"), "eastAsia")
    sz = etree.SubElement(rpr, q("sz"))
    sz.set(q("val"), str(int(round(size_pt * 2))))
    szcs = etree.SubElement(rpr, q("szCs"))
    szcs.set(q("val"), str(int(round(size_pt * 2))))
    if bold:
        etree.SubElement(rpr, q("b"))
        etree.SubElement(rpr, q("bCs"))
    lang = etree.SubElement(rpr, q("lang"))
    lang.set(q("val"), "en-US")
    lang.set(q("eastAsia"), "zh-CN")


def set_ppr(
    ppr: etree._Element,
    *,
    align: str = "both",
    first_line: int | None = None,
    hanging: int | None = None,
    left: int | None = None,
    line: int | None = 520,
    outline: int | None = None,
) -> None:
    for child in list(ppr):
        if child.tag in {q("jc"), q("ind"), q("spacing"), q("outlineLvl")}:
            ppr.remove(child)
    jc = etree.SubElement(ppr, q("jc"))
    jc.set(q("val"), align)
    if first_line is not None or hanging is not None or left is not None:
        ind = etree.SubElement(ppr, q("ind"))
        if first_line is not None:
            ind.set(q("firstLine"), str(first_line))
        if hanging is not None:
            ind.set(q("hanging"), str(hanging))
        if left is not None:
            ind.set(q("left"), str(left))
    spacing = etree.SubElement(ppr, q("spacing"))
    spacing.set(q("before"), "0")
    spacing.set(q("after"), "0")
    if line is not None:
        spacing.set(q("line"), str(line))
        spacing.set(q("lineRule"), "exact")
    else:
        spacing.set(q("lineRule"), "auto")
    if outline is not None:
        ol = etree.SubElement(ppr, q("outlineLvl"))
        ol.set(q("val"), str(outline))


def ensure_style(styles: etree._Element, style_id: str, name: str,
                 based_on: str = "Normal") -> etree._Element:
    for style in styles.findall(q("style")):
        if style.get(q("styleId")) == style_id:
            return style
    style = etree.SubElement(styles, q("style"))
    style.set(q("type"), "paragraph")
    style.set(q("styleId"), style_id)
    n = etree.SubElement(style, q("name"))
    n.set(q("val"), name)
    b = etree.SubElement(style, q("basedOn"))
    b.set(q("val"), based_on)
    etree.SubElement(style, q("qFormat"))
    return style


STYLE_SPECS = {
    "Normal": dict(align="both", first_line=480, line=520, font_cn="宋体",
                   font_en="Times New Roman", size_pt=12, bold=False),
    "BodyText": dict(align="both", first_line=480, line=520, font_cn="宋体",
                     font_en="Times New Roman", size_pt=12, bold=False),
    "Heading1": dict(align="center", first_line=0, line=520, outline=0,
                     font_cn="黑体", font_en="Times New Roman", size_pt=15, bold=True),
    "Heading2": dict(align="both", first_line=480, line=520, outline=1,
                     font_cn="黑体", font_en="Times New Roman", size_pt=14, bold=False),
    "Heading3": dict(align="both", first_line=480, line=520, outline=2,
                     font_cn="黑体", font_en="Times New Roman", size_pt=12, bold=False),
    "Heading4": dict(align="both", first_line=480, line=520, outline=9,
                     font_cn="宋体", font_en="Times New Roman", size_pt=12, bold=True),
    "Title": dict(align="center", first_line=0, line=520, outline=9,
                  font_cn="黑体", font_en="Times New Roman", size_pt=15, bold=True),
    "CoverHeader": dict(align="center", first_line=0, line=520,
                        font_cn="宋体", font_en="Times New Roman", size_pt=14, bold=False),
    "CoverTitle": dict(align="center", first_line=0, line=520,
                       font_cn="宋体", font_en="Times New Roman", size_pt=14, bold=True),
    "CoverItem": dict(align="center", first_line=0, line=520,
                      font_cn="宋体", font_en="Times New Roman", size_pt=14, bold=False),
    "AbstractTitle": dict(align="center", first_line=0, line=520,
                          font_cn="黑体", font_en="Times New Roman", size_pt=15, bold=True),
    "AbstractBodyCN": dict(align="both", first_line=480, line=520,
                           font_cn="宋体", font_en="Times New Roman", size_pt=12, bold=False),
    "AbstractBodyEN": dict(align="both", first_line=480, line=520,
                           font_cn="Times New Roman", font_en="Times New Roman",
                           size_pt=12, bold=False),
    "KeywordsCN": dict(align="both", first_line=0, line=520,
                       font_cn="宋体", font_en="Times New Roman", size_pt=12, bold=False),
    "KeywordsEN": dict(align="both", first_line=0, line=520,
                       font_cn="Times New Roman", font_en="Times New Roman",
                       size_pt=12, bold=False),
    "TOCTitle": dict(align="center", first_line=0, line=520,
                     font_cn="黑体", font_en="Times New Roman", size_pt=15, bold=True),
    "ReferenceTitle": dict(align="center", first_line=0, line=520, outline=0,
                           font_cn="黑体", font_en="Times New Roman", size_pt=15, bold=True),
    "ReferenceItem": dict(align="both", hanging=420, line=520,
                          font_cn="宋体", font_en="Times New Roman",
                          size_pt=10.5, bold=False),
    "AppendixTitle": dict(align="center", first_line=0, line=520,
                          font_cn="黑体", font_en="Times New Roman",
                          size_pt=14, bold=False),
    "Caption": dict(align="center", first_line=0, line=520,
                    font_cn="黑体", font_en="Times New Roman",
                    size_pt=10.5, bold=False),
    "TableText": dict(align="left", first_line=0, line=None,
                      font_cn="宋体", font_en="Times New Roman",
                      size_pt=10.5, bold=False),
    "Note": dict(align="both", first_line=360, line=None,
                 font_cn="宋体", font_en="Times New Roman",
                 size_pt=9, bold=False),
}


def patch_styles(data: bytes) -> bytes:
    root = etree.fromstring(data)
    for style_id, spec in STYLE_SPECS.items():
        style = ensure_style(root, style_id, style_id)
        ppr = get_or_make(style, "pPr")
        set_ppr(
            ppr,
            align=spec["align"],
            first_line=spec.get("first_line"),
            hanging=spec.get("hanging"),
            line=spec.get("line"),
            outline=spec.get("outline"),
        )
        rpr = get_or_make(style, "rPr")
        set_rpr(
            rpr,
            font_cn=spec["font_cn"],
            font_en=spec["font_en"],
            size_pt=spec["size_pt"],
            bold=spec["bold"],
        )

    # TOC 条目：点状前导符 + 右对齐页码
    for style_id, left in (("TOC1", 0), ("TOC2", 480)):
        style = ensure_style(root, style_id, style_id)
        ppr = get_or_make(style, "pPr")
        set_ppr(ppr, align="both", left=left, first_line=0, line=520)
        tabs = get_or_make(ppr, "tabs")
        tab = etree.SubElement(tabs, q("tab"))
        tab.set(q("val"), "right")
        tab.set(q("leader"), "dot")
        tab.set(q("pos"), "8306")
        rpr = get_or_make(style, "rPr")
        set_rpr(rpr, font_cn="宋体", font_en="Times New Roman", size_pt=12)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_settings(data: bytes) -> bytes:
    root = etree.fromstring(data)
    for old in root.findall(q("updateFields")):
        root.remove(old)
    el = etree.SubElement(root, q("updateFields"))
    el.set(q("val"), "true")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_sectpr(front: bool) -> etree._Element:
    sect = etree.Element(q("sectPr"))
    footer_id = "rIdDevFooterDefault"
    if front:
        ref_default = etree.SubElement(sect, q("footerReference"))
        ref_default.set(q("type"), "default")
        ref_default.set(rid("id"), footer_id)
        ref_first = etree.SubElement(sect, q("footerReference"))
        ref_first.set(q("type"), "first")
        ref_first.set(rid("id"), "rIdDevFooterFirst")
    else:
        ref_default = etree.SubElement(sect, q("footerReference"))
        ref_default.set(q("type"), "default")
        ref_default.set(rid("id"), footer_id)

    pgsz = etree.SubElement(sect, q("pgSz"))
    pgsz.set(q("w"), "11906")
    pgsz.set(q("h"), "16838")
    pgmar = etree.SubElement(sect, q("pgMar"))
    pgmar.set(q("top"), "1418")
    pgmar.set(q("right"), "1418")
    pgmar.set(q("bottom"), "1418")
    pgmar.set(q("left"), "1588")
    pgmar.set(q("header"), "851")
    pgmar.set(q("footer"), "992")
    pgmar.set(q("gutter"), "0")
    pgnum = etree.SubElement(sect, q("pgNumType"))
    if front:
        pgnum.set(q("fmt"), "upperRoman")
        pgnum.set(q("start"), "0")
    else:
        pgnum.set(q("fmt"), "decimal")
        pgnum.set(q("start"), "1")
    etree.SubElement(sect, q("cols"))
    if front:
        etree.SubElement(sect, q("titlePg"))
    etree.SubElement(sect, q("docGrid"))
    return sect


def paragraph_text(p: etree._Element) -> str:
    return "".join(p.itertext()).strip()


def replace_placeholder(body: etree._Element, marker: str,
                        replacement: etree._Element) -> bool:
    for p in body.findall(q("p")):
        if paragraph_text(p) == marker:
            p.addnext(replacement)
            body.remove(p)
            return True
    return False


def _force_title_layout(body: etree._Element) -> None:
    """给摘要/目录标题加直接格式，避免 Word 对自定义样式居中失效。"""
    for p in body.findall(q("p")):
        ppr = p.find(q("pPr"))
        if ppr is None:
            continue
        style = ppr.find(q("pStyle"))
        if style is None or style.get(q("val")) not in {"AbstractTitle", "TOCTitle"}:
            continue
        for tag in ("ind", "jc"):
            old = ppr.find(q(tag))
            if old is not None:
                ppr.remove(old)
        ind = etree.SubElement(ppr, q("ind"))
        ind.set(q("firstLine"), "0")
        jc = etree.SubElement(ppr, q("jc"))
        jc.set(q("val"), "center")


def make_toc_paragraph() -> etree._Element:
    p = etree.Element(q("p"))
    ppr = etree.SubElement(p, q("pPr"))
    pstyle = etree.SubElement(ppr, q("pStyle"))
    pstyle.set(q("val"), "TOC1")
    fld = etree.SubElement(p, q("fldSimple"))
    fld.set(q("instr"), ' TOC \\o "1-2" \\h \\z \\u ')
    fld.set(q("dirty"), "true")
    r = etree.SubElement(fld, q("r"))
    t = etree.SubElement(r, q("t"))
    t.text = "打开 Word 时自动更新目录"
    return p


def patch_document(data: bytes) -> bytes:
    root = etree.fromstring(data)
    body = root.find(q("body"))
    if body is None:
        raise RuntimeError("document.xml 缺少 w:body")

    _force_title_layout(body)

    if not replace_placeholder(body, "[[TOC]]", make_toc_paragraph()):
        raise RuntimeError("找不到目录占位符 [[TOC]]")

    front_break = etree.Element(q("p"))
    ppr = etree.SubElement(front_break, q("pPr"))
    ppr.append(make_sectpr(front=True))
    if not replace_placeholder(body, "[[BODY_START]]", front_break):
        raise RuntimeError("找不到正文占位符 [[BODY_START]]")

    final = body.find(q("sectPr"))
    if final is None:
        final = etree.SubElement(body, q("sectPr"))
    body.replace(final, make_sectpr(front=False))

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def footer_xml(with_page_number: bool) -> bytes:
    root = etree.Element(q("ftr"), nsmap={"w": W, "r": R})
    p = etree.SubElement(root, q("p"))
    ppr = etree.SubElement(p, q("pPr"))
    jc = etree.SubElement(ppr, q("jc"))
    jc.set(q("val"), "center")
    rpr = etree.SubElement(ppr, q("rPr"))
    set_rpr(rpr, font_cn="Times New Roman", font_en="Times New Roman", size_pt=12)
    if with_page_number:
        r1 = etree.SubElement(p, q("r"))
        fld1 = etree.SubElement(r1, q("fldChar"))
        fld1.set(q("fldCharType"), "begin")
        r2 = etree.SubElement(p, q("r"))
        instr = etree.SubElement(r2, q("instrText"))
        instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        instr.text = " PAGE "
        r3 = etree.SubElement(p, q("r"))
        fld3 = etree.SubElement(r3, q("fldChar"))
        fld3.set(q("fldCharType"), "separate")
        r4 = etree.SubElement(p, q("r"))
        t4 = etree.SubElement(r4, q("t"))
        t4.text = "1"
        r5 = etree.SubElement(p, q("r"))
        fld5 = etree.SubElement(r5, q("fldChar"))
        fld5.set(q("fldCharType"), "end")
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_content_types(data: bytes) -> bytes:
    root = etree.fromstring(data)
    existing = {el.get("PartName") for el in root.findall(f"{{{CT}}}Override")}
    for part in ("/word/footer1.xml", "/word/footer2.xml"):
        if part in existing:
            continue
        el = etree.SubElement(root, f"{{{CT}}}Override")
        el.set("PartName", part)
        el.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        )
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_rels(data: bytes) -> bytes:
    root = etree.fromstring(data)
    existing = {el.get("Id") for el in root.findall(f"{{{REL}}}Relationship")}
    for rel_id, target in (
        ("rIdDevFooterDefault", "footer1.xml"),
        ("rIdDevFooterFirst", "footer2.xml"),
    ):
        if rel_id in existing:
            continue
        el = etree.SubElement(root, f"{{{REL}}}Relationship")
        el.set("Id", rel_id)
        el.set(
            "Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        )
        el.set("Target", target)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def read_front_matter(text: str) -> tuple[dict, str]:
    if text.startswith("---"):
        try:
            _, meta_text, body = text.split("---", 2)
            meta = yaml.safe_load(meta_text) or {}
            return meta, body.lstrip("\n")
        except ValueError:
            pass
    return {}, text


def custom(style: str, content: str) -> str:
    return f"::: {{custom-style=\"{style}\"}}\n{content}\n:::\n"


def page_break() -> str:
    return '```{=openxml}\n<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n```\n'


def build_source(meta: dict, body: str) -> str:
    title = meta.get("title", "")
    college = meta.get("college", "")
    grade = meta.get("grade", "")
    major = meta.get("major", "")
    student_id = meta.get("student_id", "")
    name = meta.get("name", "")
    advisor = meta.get("advisor", "")
    abstract_cn = meta.get("abstract_cn", "")
    keywords_cn = meta.get("keywords_cn", "")
    abstract_en = meta.get("abstract_en", "")
    keywords_en = meta.get("keywords_en", "")

    parts: list[str] = []
    parts.append(custom("CoverHeader", "2027届本科毕业论文（设计）"))
    parts.append(custom("CoverTitle", title))
    for label, value in (
        ("学院", college),
        ("年级", grade),
        ("专业", major),
        ("学号", student_id),
        ("姓名", name),
        ("指导教师", advisor),
        ("", "2027年5月"),
    ):
        parts.append(custom("CoverItem", f"{label}　{value}" if label else value))
    parts.append(page_break())

    # HTML 实体避免 Markdown 把连续两个半角空格折叠成一个。
    parts.append(custom("AbstractTitle", "摘&#32;&#32;要"))
    parts.append(custom("AbstractBodyCN", str(abstract_cn).strip()))
    parts.append(custom("KeywordsCN", f"关键词：{keywords_cn}"))
    parts.append(page_break())

    parts.append(custom("AbstractTitle", "Abstract"))
    parts.append(custom("AbstractBodyEN", str(abstract_en).strip()))
    parts.append(custom("KeywordsEN", f"Keywords: {keywords_en}"))
    parts.append(page_break())

    parts.append(custom("TOCTitle", "目&#32;&#32;录"))
    parts.append("[[TOC]]\n")
    parts.append("[[BODY_START]]\n")

    body = re.sub(
        r"(?m)^# (参考文献|附录[^\n]*)",
        lambda m: page_break() + f"# {m.group(1)}",
        body,
    )
    parts.append(body)
    return "\n".join(parts)


def build_reference(target: Path, cache_dir: Path) -> None:
    cache_dir.mkdir(parents=True, exist_ok=True)
    if target.is_file() and target.stat().st_mtime >= Path(__file__).stat().st_mtime:
        return
    with tempfile.TemporaryDirectory(prefix="devconfig-thesis-ref-") as td:
        base = Path(td) / "default.docx"
        with base.open("wb") as fh:
            subprocess.run(
                ["pandoc", "--print-default-data-file", "reference.docx"],
                stdout=fh,
                check=True,
            )
        out = Path(td) / "ref.docx"
        with zipfile.ZipFile(base) as zin, zipfile.ZipFile(
            out, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/styles.xml":
                    data = patch_styles(data)
                elif item.filename == "word/settings.xml":
                    data = patch_settings(data)
                elif item.filename == "word/fontTable.xml":
                    data = ref_zh.patch_font_table(data)
                zout.writestr(item, data)
        shutil.copy2(out, target)


def postprocess(docx: Path) -> None:
    with tempfile.TemporaryDirectory(prefix="devconfig-thesis-post-") as td:
        tmp = Path(td) / "out.docx"
        with zipfile.ZipFile(docx) as zin, zipfile.ZipFile(
            tmp, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            names = set(zin.namelist())
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/styles.xml":
                    data = patch_styles(data)
                elif item.filename == "word/settings.xml":
                    data = patch_settings(data)
                elif item.filename == "word/fontTable.xml":
                    data = ref_zh.patch_font_table(data)
                elif item.filename == "word/document.xml":
                    data = patch_document(data)
                elif item.filename == "[Content_Types].xml":
                    data = patch_content_types(data)
                elif item.filename == "word/_rels/document.xml.rels":
                    data = patch_rels(data)
                elif item.filename == "word/footer1.xml":
                    data = footer_xml(True)
                elif item.filename == "word/footer2.xml":
                    data = footer_xml(False)
                zout.writestr(item, data)

            if "word/footer1.xml" not in names:
                zout.writestr("word/footer1.xml", footer_xml(True))
            if "word/footer2.xml" not in names:
                zout.writestr("word/footer2.xml", footer_xml(False))
        shutil.copy2(tmp, docx)


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        sys.stderr.write("用法：python3 build_word.py 输入.md 输出.docx\n")
        return 2
    source = Path(argv[1])
    output = Path(argv[2])
    text = source.read_text(encoding="utf-8")
    meta, body = read_front_matter(text)
    markdown = build_source(meta, body)

    cache_dir = Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache")) / "devconfig"
    ref = cache_dir / "thesis-reference.docx"
    build_reference(ref, cache_dir)

    with tempfile.TemporaryDirectory(prefix="devconfig-thesis-src-") as td:
        md = Path(td) / "source.md"
        md.write_text(markdown, encoding="utf-8")
        output.parent.mkdir(parents=True, exist_ok=True)
        subprocess.run(
            [
                "pandoc",
                str(md),
                "-f",
                "markdown+tex_math_single_backslash",
                "-o",
                str(output),
                "--reference-doc",
                str(ref),
            ],
            check=True,
        )

    postprocess(output)
    print(f"OK   Word 生成：{output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
