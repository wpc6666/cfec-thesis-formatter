#!/usr/bin/env python3
"""把 pandoc 默认 reference.docx 修成 Mac/Windows 双端中文字体参考文档。

用法：
    python3 ref_zh.py 默认reference.docx 输出reference.docx
    python3 ref_zh.py --verify 输出文档.docx

原理：
    1. 清掉 styles.xml 里的 *Theme 字体属性，改成显式 w:eastAsia。
    2. 正文用「宋体」，标题用「黑体」。
    3. 在每个 rPr 里写入 w:hint="eastAsia" 和 w:lang eastAsia="zh-CN"。
    4. 在 fontTable.xml 里给 5 个 Windows 规范字体写 Mac 替代名 w:altName。

这样 Windows 命中宋体/黑体/仿宋/楷体/微软雅黑；Mac 缺字体时 Word 用
Songti SC / Heiti SC / Kaiti SC / PingFang SC 兜底。
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

try:
    from lxml import etree
except ImportError:  # pragma: no cover - 只在缺依赖时触发
    sys.stderr.write(
        "缺少 lxml。安装：python3 -m pip install lxml\n"
    )
    raise SystemExit(2) from None

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

BODY_FONT = "宋体"
HEADING_FONT = "黑体"

# Windows 规范名 -> Mac 替代名。Word 只在主字体缺失时使用 altName。
FONT_MAP = {
    "宋体": "Songti SC",
    "黑体": "Heiti SC",
    "仿宋": "Songti SC",
    "楷体": "Kaiti SC",
    "微软雅黑": "PingFang SC",
}

LATIN_FONT = "Times New Roman"


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def _style_id(rpr: etree._Element) -> str | None:
    node = rpr.getparent()
    while node is not None:
        if node.tag == q("style"):
            return node.get(q("styleId"))
        node = node.getparent()
    return None


def _east_font(rpr: etree._Element) -> str:
    sid = _style_id(rpr)
    if sid and (
        sid.lower().startswith("heading")
        or sid in {"Title", "Subtitle", "AbstractTitle"}
    ):
        return HEADING_FONT
    return BODY_FONT


def _set_rfonts(rfonts: etree._Element, east: str) -> None:
    for attr in ("asciiTheme", "eastAsiaTheme", "hAnsiTheme", "cstheme"):
        rfonts.attrib.pop(q(attr), None)
    rfonts.set(q("ascii"), LATIN_FONT)
    rfonts.set(q("hAnsi"), LATIN_FONT)
    rfonts.set(q("eastAsia"), east)
    rfonts.set(q("cs"), LATIN_FONT)
    rfonts.set(q("hint"), "eastAsia")


def patch_styles(data: bytes) -> bytes:
    root = etree.fromstring(data)
    for rpr in root.iter(q("rPr")):
        rfonts = rpr.find(q("rFonts"))
        if rfonts is None:
            rfonts = etree.Element(q("rFonts"))
            rpr.insert(0, rfonts)
        _set_rfonts(rfonts, _east_font(rpr))

        lang = rpr.find(q("lang"))
        if lang is None:
            lang = etree.SubElement(rpr, q("lang"))
        lang.set(q("val"), "en-US")
        lang.set(q("eastAsia"), "zh-CN")

    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def patch_font_table(data: bytes) -> bytes:
    root = etree.fromstring(data)
    existing = {
        f.get(q("name")): f for f in root.findall(q("font")) if f.get(q("name"))
    }
    for name, alt in FONT_MAP.items():
        font = existing.get(name)
        if font is None:
            font = etree.SubElement(root, q("font"))
            font.set(q("name"), name)
            for tag, val in (
                ("charset", "86"),
                ("family", "auto"),
                ("pitch", "variable"),
            ):
                child = etree.SubElement(font, q(tag))
                child.set(q("val"), val)
            existing[name] = font
        alt_name = font.find(q("altName"))
        if alt_name is None:
            alt_name = etree.SubElement(font, q("altName"))
        alt_name.set(q("val"), alt)

    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def patch_reference(src: Path, dst: Path) -> None:
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(
        dst, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/styles.xml":
                data = patch_styles(data)
            elif item.filename == "word/fontTable.xml":
                data = patch_font_table(data)
            zout.writestr(item, data)


def _count(root: etree._Element, tag: str, attr: str, value: str) -> int:
    return sum(
        1
        for el in root.iter(q(tag))
        if el.get(q(attr)) == value
    )


def _count_attr(root: etree._Element, tag: str, attr: str) -> int:
    return sum(1 for el in root.iter(q(tag)) if q(attr) in el.attrib)


def verify(docx: Path) -> list[str]:
    problems: list[str] = []
    with zipfile.ZipFile(docx) as z:
        names = set(z.namelist())
        if "word/styles.xml" not in names:
            return ["docx 缺少 word/styles.xml"]
        if "word/fontTable.xml" not in names:
            return ["docx 缺少 word/fontTable.xml"]

        styles = etree.fromstring(z.read("word/styles.xml"))
        table = etree.fromstring(z.read("word/fontTable.xml"))

        body = _count(styles, "rFonts", "eastAsia", BODY_FONT)
        head = _count(styles, "rFonts", "eastAsia", HEADING_FONT)
        theme = _count_attr(styles, "rFonts", "eastAsiaTheme")
        lang = _count(styles, "lang", "eastAsia", "zh-CN")

        if body == 0:
            problems.append(f'styles.xml 没有 eastAsia="{BODY_FONT}"')
        if head == 0:
            problems.append(f'styles.xml 没有 eastAsia="{HEADING_FONT}"')
        if theme != 0:
            problems.append(f"styles.xml 仍残留 {theme} 个 eastAsiaTheme")
        if lang == 0:
            problems.append('styles.xml 没有 lang eastAsia="zh-CN"')

        fonts = {
            f.get(q("name")): f for f in table.findall(q("font")) if f.get(q("name"))
        }
        for name, alt in FONT_MAP.items():
            font = fonts.get(name)
            if font is None:
                problems.append(f"fontTable.xml 缺少字体条目：{name}")
                continue
            got = font.find(q("altName"))
            if got is None or got.get(q("val")) != alt:
                problems.append(
                    f"fontTable.xml {name} 的 altName 期望 {alt}，"
                    f"实际 {got.get(q('val')) if got is not None else '缺失'}"
                )

        for part in sorted(names):
            if not (
                part.startswith("word/header")
                or part.startswith("word/footer")
                or part == "word/footnotes.xml"
                or part == "word/endnotes.xml"
            ):
                continue
            root = etree.fromstring(z.read(part))
            n = _count_attr(root, "rFonts", "eastAsiaTheme")
            if n:
                problems.append(f"{part} 仍残留 {n} 个 eastAsiaTheme")

    if not problems:
        print(
            f"OK docx 字体断言通过：宋体 {body} 处 / 黑体 {head} 处 / "
            f"theme 残留 {theme} / altName {len(FONT_MAP)}/{len(FONT_MAP)}"
        )
    return problems


def main(argv: list[str]) -> int:
    if len(argv) == 3 and argv[1] == "--verify":
        docx = Path(argv[2])
        if not docx.is_file():
            sys.stderr.write(f"文件不存在：{docx}\n")
            return 2
        problems = verify(docx)
        if problems:
            sys.stderr.write("FAIL docx 字体断言：\n")
            for p in problems:
                sys.stderr.write(f"  - {p}\n")
            return 1
        return 0

    if len(argv) != 3 or argv[1] in {"-h", "--help"}:
        sys.stderr.write(
            "用法：\n"
            "  python3 ref_zh.py <默认reference.docx> <输出reference.docx>\n"
            "  python3 ref_zh.py --verify <输出文档.docx>\n"
        )
        return 0 if argv[1:2] in (["-h"], ["--help"]) else 2

    src, dst = Path(argv[1]), Path(argv[2])
    if not src.is_file():
        sys.stderr.write(f"源文件不存在：{src}\n")
        return 2
    patch_reference(src, dst)
    print(f"OK reference 已生成：{dst}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
