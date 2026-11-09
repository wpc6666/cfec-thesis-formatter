#!/usr/bin/env python3
"""把一份课程报告 docx 打包成多级项目文件夹。

输入：报告 docx + 学校/课程格式模版 docx（+ 可选的课程要求 docx）。
输出：
    <报告名>/
    ├── 00-原文件/          报告、模版、要求原件
    ├── 01-模板/
    │   ├── 封面/           从模版切出的封面
    │   └── 附件材料/       从模版切出的附件表
    ├── 02-封面/            已填写的封面
    ├── 03-输出/            正文.docx、完整报告.docx
    ├── 04-附件材料/        每张表一个子文件夹
    └── 05-测试/            完整性检查、格式自检、测试说明

用法：
    python3 make_report_project.py --report 报告.docx --template 模版.docx \
        [--requirements 要求.docx] [--out 输出根目录] [--info 信息.yaml]
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from pathlib import Path

import yaml
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

HERE = Path(__file__).resolve().parent
SKILL_DIR = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "2027"))

import format_docx  # noqa: E402
from docx_merge import merge_docx  # noqa: E402
from extract_templates import build_section_docx, section_ranges  # noqa: E402
from pdf_preview import export_pdf  # noqa: E402

SONG = "宋体"
TNR = "Times New Roman"


def safe_name(text: str) -> str:
    text = re.sub(r"[/\\:*?\"<>|]", "-", str(text)).strip()
    if text in ("", ".", ".."):
        return "课程报告"
    return text[:80]


def copy_into(src: Path, dst: Path) -> Path:
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    return dst


def set_cell_text(cell, text: str) -> None:
    para = cell.paragraphs[0]
    for child in list(para._p):
        if child.tag in (qn("w:r"), qn("w:hyperlink")):
            para._p.remove(child)
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    run = para.add_run(text)
    run.font.name = TNR
    run.font.size = Pt(12)
    run.bold = False
    run._element.get_or_add_rPr()
    rfonts = run._element.rPr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        run._element.rPr.insert(0, rfonts)
    rfonts.set(qn("w:eastAsia"), SONG)
    rfonts.set(qn("w:ascii"), TNR)
    rfonts.set(qn("w:hAnsi"), TNR)
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT


def unique_cells(row):
    out = []
    seen = set()
    for cell in row.cells:
        if id(cell._tc) in seen:
            continue
        seen.add(id(cell._tc))
        out.append(cell)
    return out


COVER_LABELS = {
    "课程名称": "course",
    "项目名称": "project",
    "学生学院": "college",
    "学院": "college",
    "年级": "grade",
    "学生专业": "major",
    "专业": "major",
    "班级": "class_name",
}


def fill_cover(cover_path: Path, info: dict) -> None:
    doc = Document(cover_path)
    if not doc.tables:
        doc.save(cover_path)
        return
    table = doc.tables[0]
    for row in table.rows:
        cells = unique_cells(row)
        for i, cell in enumerate(cells[:-1]):
            label = cell.text.strip().rstrip("：:")
            key = COVER_LABELS.get(label)
            if key and info.get(key):
                set_cell_text(cells[i + 1], str(info[key]))
            elif label in ("学生姓名", "姓名"):
                members = info.get("members") or []
                target = cells[i + 1]
                if members and isinstance(members[0], dict):
                    set_cell_text(target, str(members[0].get("name", "")))
                else:
                    set_cell_text(target, "")
            elif label == "学号":
                members = info.get("members") or []
                target = cells[i + 1]
                if members and isinstance(members[0], dict):
                    set_cell_text(target, str(members[0].get("id", "")))
                else:
                    set_cell_text(target, "")
    # 清掉模版里的 ××× 占位
    for row in table.rows:
        for cell in unique_cells(row):
            if "×××" in cell.text:
                set_cell_text(cell, "")
    doc.save(cover_path)


def extract_template_parts(template: Path, out_dir: Path) -> tuple[Path, Path]:
    import zipfile

    from lxml import etree

    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with zipfile.ZipFile(template) as z:
        body = etree.fromstring(z.read("word/document.xml")).find(W + "body")
    ranges, final_sect = section_ranges(body)
    # 单节模板会把整份内容同时当成封面和评价表，导致成品重复。
    if len(ranges) < 2:
        raise RuntimeError(
            f"模版至少要有「封面」和「评价表」两节，当前只有 {len(ranges)} 节：{template}")

    cover = out_dir / "封面" / "封面模板.docx"
    build_section_docx(template, cover, 0, 0, final_sect)

    attach = out_dir / "附件材料" / "课程报告基础信息及评价表.docx"
    last = len(ranges) - 1
    build_section_docx(template, attach, last, last, final_sect)
    return cover, attach


def strip_title_page(report: Path, out: Path) -> None:
    doc = Document(report)
    body = doc.element.body
    children = list(body)
    cut = None
    for i, ch in enumerate(children):
        if ch.tag == qn("w:p") and any(
            br.get(qn("w:type")) == "page" for br in ch.iter(qn("w:br"))
        ):
            cut = i
            break
    if cut is not None:
        for ch in children[: cut + 1]:
            body.remove(ch)
    doc.save(out)


def set_body_margins(path: Path) -> None:
    doc = Document(path)
    for section in doc.sections:
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
    doc.save(path)


def read_info(path: Path | None, report: Path) -> dict:
    info: dict = {}
    if path and path.is_file():
        info = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    doc = Document(report)
    texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    if "course" not in info and texts:
        course = texts[0].replace("课程报告", "").strip()
        info["course"] = course.strip("《》 ").strip()
    if "project" not in info and len(texts) > 1:
        info["project"] = texts[1]
    info.setdefault("members", [])
    return info


def write_info(info: dict, path: Path) -> None:
    path.write_text(
        yaml.safe_dump(info, allow_unicode=True, sort_keys=False), encoding="utf-8"
    )


def run_checks(project: Path, has_template: bool) -> list[tuple[str, bool, str]]:
    checks = []

    def check(label: str, rel: str, required: bool = True) -> None:
        p = project / rel
        if required:
            checks.append((label, p.exists(), rel))

    check("原文件：原始报告", "00-原文件/原始报告.docx")
    check("输出：正文", "03-输出/正文.docx")
    check("测试：格式自检", "05-测试/格式自检.md")
    if has_template:
        for label, rel in (
            ("原文件：格式模版", "00-原文件/格式模版.docx"),
            ("模板：封面", "01-模板/封面/封面模板.docx"),
            ("模板：附件表", "01-模板/附件材料/课程报告基础信息及评价表.docx"),
            ("已填封面", "02-封面/封面.docx"),
            ("输出：完整报告", "03-输出/完整报告.docx"),
            ("附件材料：评价表", "04-附件材料/01-课程报告基础信息及评价表/课程报告基础信息及评价表.docx"),
        ):
            check(label, rel)
        full = project / "03-输出/完整报告.docx"
        if full.is_file():
            try:
                doc = Document(full)
                checks.append(("完整报告可打开", True, f"{len(doc.paragraphs)} 段 / {len(doc.tables)} 表"))
                checks.append(("完整报告至少三节", len(doc.sections) >= 3, f"{len(doc.sections)} 节"))
            except Exception as exc:  # noqa: BLE001
                checks.append(("完整报告可打开", False, str(exc)))
    else:
        for label, rel in (
            ("模板：未提供说明", "01-模板/说明.md"),
            ("封面：未提供说明", "02-封面/说明.md"),
            ("附件：未提供说明", "04-附件材料/说明.md"),
        ):
            check(label, rel)
    return checks


def write_test_files(project: Path, info: dict, checks, warnings,
                     has_template: bool) -> None:
    test_dir = project / "05-测试"
    test_dir.mkdir(parents=True, exist_ok=True)
    failed = [c for c in checks if not c[1]]
    rows = ["| 检查项 | 结果 | 说明 |", "|---|---|---|"]
    for label, ok, detail in checks:
        rows.append(f"| {label} | {'通过' if ok else '缺失'} | {detail} |")
    (test_dir / "完整性检查.md").write_text(
        "\n".join(
            [
                "# 完整性检查",
                "",
                f"- 课程：{info.get('course', '')}",
                f"- 项目：{info.get('project', '')}",
                f"- 模式：{'使用课程模版' if has_template else '不使用模板'}",
                f"- 结论：{'全部通过' if not failed else f'{len(failed)} 项缺失'}",
                "",
                *rows,
                "",
                "## 说明",
                "",
                "- 检查文件存在、docx 能打开。",
                "- 封面姓名、学号、成绩、签字需人工填写或核对。",
                "",
            ]
        ),
        encoding="utf-8",
    )
    target = "03-输出/完整报告.docx" if has_template else "03-输出/正文.docx"
    (test_dir / "测试说明.md").write_text(
        "# 测试说明\n\n"
        f"1. 打开 `{target}`，检查正文与分页。\n"
        + ("2. 检查 `04-附件材料/` 下的评价表。\n" if has_template else
           "2. 本次未提供模板，没有封面和附件表，封面/成绩/签字需另附。\n")
        + "3. PDF 预览：已授权 Word 时自动生成；未授权先运行 "
        "`bash <skill目录>/authorize-word.sh`。\n",
        encoding="utf-8",
    )
    target_path = project / target
    pdf_name = "完整报告预览.pdf" if has_template else "正文预览.pdf"
    ok, how = export_pdf(target_path, test_dir / pdf_name)
    if not ok:
        warnings.append(
            f"未生成 PDF 预览：{how}。运行 authorize-word.sh 授权 Word 后可自动导出。"
        )
    if warnings:
        (test_dir / "提醒.md").write_text(
            "# 提醒\n\n" + "\n".join(f"- {w}" for w in warnings) + "\n",
            encoding="utf-8",
        )


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="课程报告 → 多级项目文件夹（给模版 / 不给模版）")
    ap.add_argument("--report", type=Path, required=True)
    ap.add_argument("--template", type=Path, help="课程/学校格式模版 docx")
    ap.add_argument("--no-template", action="store_true", help="明确不使用任何模板")
    ap.add_argument("--requirements", type=Path, help="课程要求 docx（不是模板）")
    ap.add_argument("--out", type=Path, default=Path.cwd())
    ap.add_argument("--info", type=Path, help="可选 YAML：course/project/college/grade/major/class_name/members")
    args = ap.parse_args(argv)

    if args.template and args.no_template:
        print("--template 与 --no-template 只能选一个。", file=sys.stderr)
        return 2
    if not args.template and not args.no_template:
        print(
            "未指定模板。请先确认：\n"
            "  1. 用课程/学校模版：--template 模版.docx\n"
            "  2. 不用模板：--no-template\n"
            "  3. 用 2027 论文模版：改用 make-thesis-2027.sh\n"
            "不会默认套用论文模板。",
            file=sys.stderr,
        )
        return 2
    if not args.report.is_file():
        print(f"报告不存在：{args.report}", file=sys.stderr)
        return 2
    if args.template and not args.template.is_file():
        print(f"模板不存在：{args.template}", file=sys.stderr)
        return 2
    if args.info and not args.info.is_file():
        print(f"--info 文件不存在：{args.info}", file=sys.stderr)
        return 2
    if args.requirements and not args.requirements.is_file():
        print(f"--requirements 文件不存在：{args.requirements}", file=sys.stderr)
        return 2
    has_template = bool(args.template)

    info = read_info(args.info, args.report)
    project = args.out / safe_name(info.get("project") or args.report.stem)
    if project.is_file():
        print(f"目标路径已是文件，未覆盖：{project}", file=sys.stderr)
        return 3
    if project.exists() and any(project.iterdir()):
        print(f"目标目录已存在且非空，未覆盖：{project}", file=sys.stderr)
        return 3

    dirs = {
        "origin": project / "00-原文件",
        "templates": project / "01-模板",
        "covers": project / "02-封面",
        "output": project / "03-输出",
        "attachments": project / "04-附件材料",
        "test": project / "05-测试",
        "tmp": project / "05-测试" / "中间文件",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)

    # 00 原文件
    copy_into(args.report, dirs["origin"] / "原始报告.docx")
    if args.template:
        copy_into(args.template, dirs["origin"] / "格式模版.docx")
    if args.requirements:
        copy_into(args.requirements, dirs["origin"] / "课程报告要求.docx")
    write_info(info, dirs["origin"] / "报告信息.yaml")
    note_lines = [
        "# 原文件（只读）",
        "",
        "- `原始报告.docx`：你提供的课程报告，未修改。",
        f"- 模式：{'使用课程模版' if has_template else '不使用模板'}。",
    ]
    if args.template:
        note_lines.append("- `格式模版.docx`：课程/学校格式模版。")
    if args.requirements:
        note_lines.append("- `课程报告要求.docx`：课程要求与评分标准，不是模板。")
    note_lines.append("- `报告信息.yaml`：本次生成使用的字段。")
    (dirs["origin"] / "说明.md").write_text("\n".join(note_lines) + "\n", encoding="utf-8")

    generated: list[str] = []
    warnings: list[str] = []

    if has_template:
        # 01 模板：只放真正的模板，不把“课程报告要求”当模板
        copy_into(args.template, dirs["templates"] / "格式模版.docx")
        cover_tpl, attach_tpl = extract_template_parts(args.template, dirs["templates"])
        generated += [str(p.relative_to(project)) for p in (cover_tpl, attach_tpl)]

        # 02 封面
        cover = dirs["covers"] / "封面.docx"
        shutil.copy2(cover_tpl, cover)
        fill_cover(cover, info)
        generated.append(str(cover.relative_to(project)))

        # 03 输出：去标题页 → 刷格式 → 合并封面与附件
        body_raw = dirs["tmp"] / "正文（去标题页）.docx"
        strip_title_page(args.report, body_raw)
        body = dirs["output"] / "正文.docx"
        format_docx.format_docx(
            body_raw, body, toc=False, page_numbers=True, three_line_tables=True
        )
        set_body_margins(body)
        generated.append(str(body.relative_to(project)))

        merged1 = dirs["tmp"] / "封面+正文.docx"
        full = dirs["output"] / "完整报告.docx"
        merge_docx(cover, body, merged1)
        merge_docx(merged1, attach_tpl, full)
        generated.append(str(full.relative_to(project)))

        # 04 附件材料
        att_dir = dirs["attachments"] / "01-课程报告基础信息及评价表"
        att_dir.mkdir(exist_ok=True)
        copy_into(attach_tpl, att_dir / "课程报告基础信息及评价表.docx")
        generated.append(
            str((att_dir / "课程报告基础信息及评价表.docx").relative_to(project))
        )
        if args.requirements:
            req_dir = dirs["attachments"] / "02-课程报告要求及评分标准"
            req_dir.mkdir(exist_ok=True)
            copy_into(args.requirements, req_dir / "课程报告要求及评分标准.docx")
            generated.append(
                str((req_dir / "课程报告要求及评分标准.docx").relative_to(project))
            )
        (dirs["attachments"] / "说明.md").write_text(
            "# 附件材料\n\n"
            "- `01-课程报告基础信息及评价表/`：从格式模版切出的组内评价表。\n"
            "- `02-课程报告要求及评分标准/`：课程要求文件（如有）。\n",
            encoding="utf-8",
        )
        if not info.get("members"):
            warnings.append(
                "封面学生姓名与学号为空：课程报告是小组作业，请在 02-封面/封面.docx 里补成员。"
            )

    else:
        # 不给模板：保留原报告结构，不生成封面/附件，也不套用论文模板
        (dirs["templates"] / "说明.md").write_text(
            "# 未提供模板\n\n"
            "本次没有提供课程/学校模版，所以没有生成封面、附件表，也没有套用论文模板。\n"
            "如果要用 2027 论文模版，请改用 `make-thesis-2027.sh`。\n",
            encoding="utf-8",
        )
        (dirs["covers"] / "说明.md").write_text(
            "# 未生成封面\n\n未提供模板，封面保持原报告自己的标题页。\n",
            encoding="utf-8",
        )
        (dirs["attachments"] / "说明.md").write_text(
            "# 未生成附件材料\n\n未提供模板，不生成课程评价表或其他附件。\n",
            encoding="utf-8",
        )
        body = dirs["output"] / "正文.docx"
        format_docx.format_docx(
            args.report, body, toc=False, page_numbers=True, three_line_tables=True
        )
        set_body_margins(body)
        generated.append(str(body.relative_to(project)))
        warnings.append("未提供模板：没有封面和附件表，按原报告标题页输出。")

    # 05 测试
    steps = (
        [
            "1. 打开 `02-封面/封面.docx`，补学生姓名、学号、班级、报告成绩。",
            "2. 打开 `03-输出/完整报告.docx`，检查封面、正文、附件表顺序。",
            "3. 检查图表是否跨页、分页是否合理。",
            "4. 打印前核对课程报告要求与命名。",
        ]
        if has_template
        else [
            "1. 打开 `03-输出/正文.docx`，检查正文与分页。",
            "2. 本次未提供模板，没有封面和附件表；封面/成绩/签字需另附。",
            "3. 检查图表是否跨页、分页是否合理。",
            "4. 打印前核对课程报告要求与命名。",
        ]
    )
    (dirs["test"] / "格式自检.md").write_text(
        "\n".join(
            [
                "# 生成报告（课程报告）",
                "",
                f"- 课程：{info.get('course', '')}",
                f"- 项目：{info.get('project', '')}",
                f"- 模式：{'使用课程模版' if has_template else '不使用模板'}",
                "",
                "## 已生成",
                "",
                *[f"- `{x}`" for x in generated],
                "",
                "## 需要你做的事",
                "",
                *steps,
                "",
            ]
        ),
        encoding="utf-8",
    )
    checks = run_checks(project, has_template)
    write_test_files(project, info, checks, warnings, has_template)

    print(f"OK   {project}")
    for w in warnings:
        print(f"提醒 {w}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
