#!/usr/bin/env python3
"""按“节”合并两个 docx，保留各自的页面设置、页脚与页码。

docxcompose 会把后一个文档的分节符并掉（封面与正文变成同一节），
这里直接操作 OOXML：把 base 的末节 sectPr 变成段落级分节符，
再把 part 的正文与末节 sectPr 接上，同时迁移样式、编号、页脚和图片关系。
"""

from __future__ import annotations

import copy
import posixpath
import re
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

REL_FOOTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
)
REL_HEADER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
)

# 这些部件由 base 提供或单独合并，不能再作为独立部件导入
SKIP_RELTYPES = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument",
}
CT_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
CT_HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def r_attr(tag: str) -> str:
    return f"{{{R}}}{tag}"


def parse(data: bytes):
    return etree.fromstring(data)


def to_bytes(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def convert_last_section_to_break(body, sect_pr) -> None:
    """把 body 级 sectPr 移到文末一个 1 磅空段落里，形成分节符。"""
    body.remove(sect_pr)
    p = etree.SubElement(body, q("p"))
    ppr = etree.SubElement(p, q("pPr"))
    spacing = etree.SubElement(ppr, q("spacing"))
    spacing.set(q("before"), "0")
    spacing.set(q("after"), "0")
    spacing.set(q("line"), "20")
    spacing.set(q("lineRule"), "exact")
    rpr = etree.SubElement(ppr, q("rPr"))
    sz = etree.SubElement(rpr, q("sz"))
    sz.set(q("val"), "2")
    szcs = etree.SubElement(rpr, q("szCs"))
    szcs.set(q("val"), "2")
    ppr.append(sect_pr)


def content_type_for_part(name: str) -> tuple[str, str] | None:
    if re.search(r"/footer\d*\.xml$", name):
        return "Override", CT_FOOTER
    if re.search(r"/header\d*\.xml$", name):
        return "Override", CT_HEADER
    return None


def import_relationships(
    part_zip,
    imported: dict[str, bytes],
    base_rels: etree._Element,
    part_rels: etree._Element | None,
    reserved: set[str] | None = None,
) -> dict[str, str]:
    """把 part 的 document.xml.rels 迁移到 base，返回 旧 rId -> 新 rId。"""
    if part_rels is None:
        return {}

    existing_ids = {el.get("Id") for el in base_rels}
    used_targets = set(imported) | (reserved or set())
    counter = [1]

    def next_id() -> str:
        while True:
            rid = f"rIdPart{counter[0]}"
            counter[0] += 1
            if rid not in existing_ids:
                existing_ids.add(rid)
                return rid

    mapping: dict[str, str] = {}
    for rel in part_rels.findall(f"{{{REL}}}Relationship"):
        old = rel.get("Id")
        target = rel.get("Target", "")
        rtype = rel.get("Type", "")
        if rtype in SKIP_RELTYPES:
            continue
        external = rel.get("TargetMode") == "External"
        new_id = next_id()
        if external:
            new_rel = etree.SubElement(base_rels, f"{{{REL}}}Relationship")
            new_rel.set("Id", new_id)
            new_rel.set("Type", rtype)
            new_rel.set("Target", target)
            new_rel.set("TargetMode", "External")
            mapping[old] = new_id
            continue

        src_name = posixpath.normpath(
            posixpath.join("word", target.lstrip("/"))
            if not target.startswith("/")
            else target.lstrip("/")
        )
        if src_name not in part_zip.namelist():
            # 目标缺失时保留原关系，避免整体失败
            continue
        data = part_zip.read(src_name)
        # 同名同内容复用；同名不同内容改名，避免覆盖 base 的页脚/图片
        if imported.get(src_name) == data:
            new_name = src_name
        elif rtype in (REL_FOOTER, REL_HEADER):
            stem, ext = posixpath.splitext(src_name)
            prefix = "footer" if rtype == REL_FOOTER else "header"
            n = 1
            while f"word/{prefix}{n}{ext}" in used_targets:
                n += 1
            new_name = f"word/{prefix}{n}{ext}"
        elif src_name in used_targets:
            stem, ext = posixpath.splitext(src_name)
            n = 1
            while f"{stem}_part{n}{ext}" in used_targets:
                n += 1
            new_name = f"{stem}_part{n}{ext}"
        else:
            new_name = src_name
        imported[new_name] = data
        used_targets.add(new_name)

        new_rel = etree.SubElement(base_rels, f"{{{REL}}}Relationship")
        new_rel.set("Id", new_id)
        new_rel.set("Type", rtype)
        new_rel.set("Target", posixpath.relpath(new_name, "word"))
        mapping[old] = new_id
    return mapping


def remap_ids(el, mapping: dict[str, str]) -> None:
    for node in el.iter():
        for attr in (r_attr("id"), r_attr("embed"), r_attr("link")):
            value = node.get(attr)
            if value in mapping:
                node.set(attr, mapping[value])


def merge_styles(base_root, part_root) -> None:
    if part_root is None:
        return
    existing = {s.get(q("styleId")) for s in base_root.findall(q("style"))}
    for style in part_root.findall(q("style")):
        sid = style.get(q("styleId"))
        if sid and sid not in existing:
            base_root.append(copy.deepcopy(style))
            existing.add(sid)


def merge_numbering(base_root, part_root, appended_elems) -> None:
    """迁移编号定义并改写正文中的 numId。"""
    if part_root is None:
        return
    abs_ids = [
        int(a.get(q("abstractNumId"))) for a in part_root.findall(q("abstractNum"))
    ]
    num_ids = [int(n.get(q("numId"))) for n in part_root.findall(q("num"))]
    abs_offset = (
        max(
            [
                int(a.get(q("abstractNumId")))
                for a in base_root.findall(q("abstractNum"))
            ]
            + [-1]
        )
        + 1
    )
    num_offset = (
        max([int(n.get(q("numId"))) for n in base_root.findall(q("num"))] + [0]) + 1
    )
    abs_map = {old: old + abs_offset for old in abs_ids}
    num_map = {old: old + num_offset for old in num_ids}

    for abstract in part_root.findall(q("abstractNum")):
        new = copy.deepcopy(abstract)
        old = int(new.get(q("abstractNumId")))
        new.set(q("abstractNumId"), str(abs_map[old]))
        for level in new.findall(q("lvl")):
            for num_id in level.findall(q("numId")):
                num_id.set(q("val"), str(abs_map[int(num_id.get(q("val")))]))
        base_root.append(new)
    for num in part_root.findall(q("num")):
        new = copy.deepcopy(num)
        old = int(new.get(q("numId")))
        new.set(q("numId"), str(num_map[old]))
        ref = new.find(q("abstractNumId"))
        if ref is not None:
            ref.set(q("val"), str(abs_map[int(ref.get(q("val")))]))
        base_root.append(new)

    for el in appended_elems:
        for node in el.iter(q("numId")):
            old = int(node.get(q("val")))
            if old in num_map:
                node.set(q("val"), str(num_map[old]))


DEFAULT_CONTENT_TYPES = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "wmf": "image/x-wmf",
    "emf": "image/x-emf",
    "svg": "image/svg+xml",
    "bin": "application/vnd.openxmlformats-officedocument.oleObject",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
}


def add_content_types(ct_root, parts: dict[str, bytes]) -> None:
    existing_override = {
        el.get("PartName") for el in ct_root.findall(f"{{{CT}}}Override")
    }
    existing_default = {
        el.get("Extension") for el in ct_root.findall(f"{{{CT}}}Default")
    }
    for name in parts:
        part_name = "/" + name
        if part_name in existing_override:
            continue
        if re.search(r"/footer[\w-]*\.xml$", name):
            el = etree.SubElement(ct_root, f"{{{CT}}}Override")
            el.set("PartName", part_name)
            el.set("ContentType", CT_FOOTER)
            existing_override.add(part_name)
        elif re.search(r"/header[\w-]*\.xml$", name):
            el = etree.SubElement(ct_root, f"{{{CT}}}Override")
            el.set("PartName", part_name)
            el.set("ContentType", CT_HEADER)
            existing_override.add(part_name)
        else:
            ext = posixpath.splitext(name)[1].lstrip(".").lower()
            if ext and ext not in existing_default:
                el = etree.SubElement(ct_root, f"{{{CT}}}Default")
                el.set("Extension", ext)
                el.set(
                    "ContentType",
                    DEFAULT_CONTENT_TYPES.get(ext, "application/octet-stream"),
                )
                existing_default.add(ext)


def merge_docx(base_path: Path, part_path: Path, out_path: Path) -> None:
    with zipfile.ZipFile(base_path) as bz:
        base_parts = {item.filename: bz.read(item.filename) for item in bz.infolist()}
        base_order = [item.filename for item in bz.infolist()]
    with zipfile.ZipFile(part_path) as pz:
        part_parts = {item.filename: pz.read(item.filename) for item in pz.infolist()}

    base_doc = parse(base_parts["word/document.xml"])
    base_body = base_doc.find(q("body"))
    base_sect = base_body.find(q("sectPr"))
    if base_sect is None:
        raise RuntimeError(f"{base_path} 缺少 sectPr")
    convert_last_section_to_break(base_body, base_sect)

    part_doc = parse(part_parts["word/document.xml"])
    part_body = part_doc.find(q("body"))
    part_sect = part_body.find(q("sectPr"))
    part_elems = [el for el in part_body if el.tag != q("sectPr")]

    # 关系迁移：footer/header/media/hyperlink
    base_rels_name = "word/_rels/document.xml.rels"
    base_rels = (
        parse(base_parts[base_rels_name])
        if base_rels_name in base_parts
        else etree.Element(f"{{{REL}}}Relationships")
    )
    part_rels = None
    if "word/_rels/document.xml.rels" in part_parts:
        part_rels = parse(part_parts["word/_rels/document.xml.rels"])
    with zipfile.ZipFile(part_path) as pz:
        imported: dict[str, bytes] = {}
        mapping = import_relationships(
            pz, imported, base_rels, part_rels, set(base_parts)
        )

    for el in part_elems:
        remap_ids(el, mapping)
        base_body.append(copy.deepcopy(el))

    # 页码/页眉关系在 sectPr 中
    if part_sect is not None:
        remap_ids(part_sect, mapping)
        base_body.append(copy.deepcopy(part_sect))

    # 样式
    if "word/styles.xml" in part_parts:
        base_styles = parse(base_parts["word/styles.xml"])
        merge_styles(base_styles, parse(part_parts["word/styles.xml"]))
        base_parts["word/styles.xml"] = to_bytes(base_styles)
    # 编号
    if "word/numbering.xml" in part_parts and "word/numbering.xml" in base_parts:
        base_numbering = parse(base_parts["word/numbering.xml"])
        merge_numbering(
            base_numbering,
            parse(part_parts["word/numbering.xml"]),
            part_elems,
        )
        base_parts["word/numbering.xml"] = to_bytes(base_numbering)

    # 设置：打开时更新目录
    settings_name = "word/settings.xml"
    if settings_name in base_parts:
        settings = parse(base_parts[settings_name])
        for old in settings.findall(q("updateFields")):
            settings.remove(old)
        upd = etree.SubElement(settings, q("updateFields"))
        upd.set(q("val"), "true")
        base_parts[settings_name] = to_bytes(settings)

    base_parts["word/document.xml"] = to_bytes(base_doc)
    base_parts["word/_rels/document.xml.rels"] = to_bytes(base_rels)
    base_parts.update(imported)

    # 内容类型
    ct = parse(base_parts["[Content_Types].xml"])
    add_content_types(ct, imported)
    base_parts["[Content_Types].xml"] = to_bytes(ct)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        written = set()
        for name in base_order:
            zout.writestr(name, base_parts[name])
            written.add(name)
        for name, data in base_parts.items():
            if name not in written:
                zout.writestr(name, data)


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 4:
        raise SystemExit("用法：merge_docx.py base.docx part.docx out.docx")
    merge_docx(Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3]))
