#!/usr/bin/env python3
"""一键生成 2027 届重财毕业论文项目文件夹。

输入：一份带 YAML 头的 Markdown（题目、学院、专业、班级、学号、姓名、
指导教师、中英文摘要与关键词，正文用 # / ## / ### 标题）。
输出：项目文件夹，含原文件、模板、论文（封面到附录）、附件材料（每表一个子文件夹）。

用法：
    python3 make_project.py 论文.md [--out 输出目录] [--type 毕业论文|毕业设计]
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL_DIR = HERE.parent.parent
ASSETS = SKILL_DIR / "assets"
YEAR_DIR = ASSETS / "2027"
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent))

import thesis2027 as T  # noqa: E402
from pdf_preview import export_pdf  # noqa: E402

FORM_FOLDERS = [
    ("01-综合成绩评定表", "01-综合成绩评定表.docx"),
    ("02-评分表", "02-评分表.docx"),
    ("03-指导教师评阅表", "03-指导教师评阅表.docx"),
    ("04-交叉评阅表", "04-交叉评阅表.docx"),
    ("05-答辩评分表", "05-答辩评分表.docx"),
    ("06-答辩记录", "06-答辩记录.docx"),
    ("07-任务书", "07-任务书.docx"),
    ("08-开题报告", "08-开题报告.docx"),
    ("09-中期检查表", "09-中期检查表.docx"),
    ("10-指导记录表", "10-指导记录表.docx"),
    ("11-承诺书", "11-承诺书.docx"),
]

REQUIRED = [
    "title",
    "college",
    "grade",
    "major",
    "class_name",
    "student_id",
    "name",
    "advisor",
    "abstract_cn",
    "keywords_cn",
]


def safe_name(text: str) -> str:
    text = re.sub(r"[/\\:*?\"<>|]", "-", str(text)).strip()
    return text[:60] or "未命名论文"


def copy_tree(src: Path, dst: Path) -> None:
    dst.mkdir(parents=True, exist_ok=True)
    for item in src.iterdir():
        if item.name in {".DS_Store", "__pycache__"}:
            continue
        target = dst / item.name
        if item.is_dir():
            shutil.copytree(item, target, dirs_exist_ok=True,
                            ignore=shutil.ignore_patterns(".DS_Store", "__pycache__"))
        else:
            shutil.copy2(item, target)


def write_info_yaml(info: dict, path: Path) -> None:
    import yaml

    path.write_text(
        yaml.safe_dump(info, allow_unicode=True, sort_keys=False), encoding="utf-8"
    )


def build_report(
    info: dict, project: Path, generated: list[str], warnings: list[str]
) -> str:
    lines = [
        "# 生成报告（2027 届）",
        "",
        f"- 题目：{info.get('title', '')}",
        f"- 类型：{info.get('type', '毕业论文')}",
        f"- 学院 / 专业 / 班级：{info.get('college', '')} / {info.get('major', '')} / {info.get('class_name', '')}",
        f"- 学号 / 姓名 / 指导教师：{info.get('student_id', '')} / {info.get('name', '')} / {info.get('advisor', '')}",
        "",
        "## 已生成",
        "",
    ]
    lines += [f"- `{x}`" for x in generated]
    lines += [
        "",
        "## 需要你在 Word 里做的事",
        "",
        "1. 打开 `03-输出/完整论文.docx`，按 `Ctrl+A` 再按 `F9`，更新目录页码。",
        "2. 检查封面题目是否超出行宽、是否自动换行。",
        "3. 逐页核对附件表中的姓名、学号、专业、题目。",
        "4. 图、表、公式的编号按章手写（`图1-1  图题` / `表1-1  表题`），生成器只负责排版。",
        "5. 打印：正文封面 A3 白色平纹纸左侧胶装；附件材料 A4 长尾夹左侧竖装。",
        "",
    ]
    if warnings:
        lines += ["## 提醒", ""] + [f"- {w}" for w in warnings] + [""]
    lines += [
        "## 未自动验证",
        "",
        "- 目录页码：需 Word 更新域后才准确。",
        "- 字体是否安装：封面用到方正小标宋，若本机没有会替换；以学校电脑打印为准。",
        "- 打印幅面与装订：需人工确认。",
        "",
    ]
    return "\n".join(lines)


def collect_checks(project: Path, info: dict) -> list[tuple[str, bool, str]]:
    """检查关键目录与文件是否齐全、docx 是否能打开。"""
    from docx import Document

    checks: list[tuple[str, bool, str]] = []

    def path_check(label: str, rel: str) -> None:
        p = project / rel
        checks.append((label, p.exists(), rel))

    path_check("原文件：学校手册", "00-原文件/附件4：重庆财经学院毕业论文（设计）指导手册（2027届）.doc")
    path_check("原文件：原稿副本", "00-原文件/论文信息.yaml")
    path_check("模板：毕业论文封面", "01-模板/封面/封面-毕业论文.docx")
    path_check("模板：毕业设计封面", "01-模板/封面/封面-毕业设计.docx")
    path_check("模板：附件材料封面", "01-模板/封面/封面-附件材料.docx")
    for _, filename in FORM_FOLDERS:
        path_check(f"模板：{filename}", f"01-模板/附件材料/{filename}")
    path_check("已填封面：正文", "02-封面/正文封面.docx")
    path_check("已填封面：附件材料", "02-封面/附件材料封面.docx")
    path_check("输出：完整论文", "03-输出/完整论文.docx")
    for prefix, name in (
        ("01", "中文摘要"),
        ("02", "英文摘要"),
        ("03", "正文"),
        ("04", "参考文献"),
        ("05", "附录"),
    ):
        path_check(f"输出：{name}", f"03-输出/{prefix}-{name}.docx")
    for folder, filename in FORM_FOLDERS:
        path_check(f"附件材料：{folder}", f"04-附件材料/{folder}/{filename}")
    path_check("测试：格式自检", "05-测试/格式自检.md")

    full = project / "03-输出/完整论文.docx"
    if full.is_file():
        try:
            doc = Document(full)
            checks.append(("完整论文可打开", True, f"{len(doc.paragraphs)} 段 / {len(doc.tables)} 表"))
            checks.append(("完整论文三节", len(doc.sections) == 3, f"{len(doc.sections)} 节"))
            has_toc = any("TOC" in (el.text or "") for el in doc.element.body.iter() if el.tag.endswith("instrText"))
            checks.append(("目录域存在", has_toc, "TOC 域"))
        except Exception as exc:  # noqa: BLE001
            checks.append(("完整论文可打开", False, str(exc)))

    return checks


def write_test_files(project: Path, info: dict, checks: list[tuple[str, bool, str]],
                     warnings: list[str]) -> None:
    test_dir = project / "05-测试"
    test_dir.mkdir(parents=True, exist_ok=True)
    rows = ["| 检查项 | 结果 | 说明 |", "|---|---|---|"]
    for label, ok, detail in checks:
        rows.append(f"| {label} | {'通过' if ok else '缺失'} | {detail} |")
    failed = [c for c in checks if not c[1]]
    lines = [
        "# 完整性检查",
        "",
        f"- 题目：{info.get('title', '')}",
        f"- 类型：{info.get('type', '毕业论文')}",
        f"- 结论：{'全部通过' if not failed else f'{len(failed)} 项缺失'}",
        "",
        *rows,
        "",
        "## 说明",
        "",
        "- 本检查只验证文件存在、docx 能打开、完整论文为三节且含目录域。",
        "- 目录页码、字体替换、打印幅面需在 Word 中人工确认。",
        "",
    ]
    if warnings:
        lines += ["## 提醒", ""] + [f"- {w}" for w in warnings] + [""]
    (test_dir / "完整性检查.md").write_text("\n".join(lines), encoding="utf-8")

    (test_dir / "测试说明.md").write_text(
        "# 测试说明\n\n"
        "1. 打开 `03-输出/完整论文.docx`，按 `Ctrl+A`、`F9` 更新目录页码。\n"
        "2. 检查封面题目、学院、年级、专业、班级、学号、姓名、指导教师。\n"
        "3. 逐张打开 `04-附件材料/` 下的表，核对基本信息。\n"
        "4. 查看 `完整性检查.md` 是否有“缺失”。\n"
        "5. PDF 预览：已授权 Word 时自动生成；未授权先运行 "
        "`bash <skill目录>/authorize-word.sh`。\n",
        encoding="utf-8",
    )

    full = project / "03-输出/完整论文.docx"
    ok, how = export_pdf(full, test_dir / "完整论文预览.pdf")
    if not ok:
        warnings.append(
            f"未生成 PDF 预览：{how}。运行 authorize-word.sh 授权 Word 后可自动导出。"
        )
        lines += ["## 提醒", ""] + [f"- {w}" for w in warnings] + [""]
        (test_dir / "完整性检查.md").write_text("\n".join(lines), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="一键生成 2027 届重财毕业论文项目")
    ap.add_argument("input", type=Path, help="带 YAML 头的论文 Markdown")
    ap.add_argument(
        "--out", type=Path, default=Path.cwd(), help="输出根目录（默认当前目录）"
    )
    ap.add_argument("--type", choices=["毕业论文", "毕业设计"], default=None)
    ap.add_argument("--templates", type=Path, default=YEAR_DIR / "模板")
    ap.add_argument("--manual", type=Path, default=YEAR_DIR / "原文件")
    ap.add_argument("--latex", action="store_true", help="同时生成正文 LaTeX（不编译 PDF）")
    args = ap.parse_args(argv)

    if not args.input.is_file():
        print(f"输入不存在：{args.input}", file=sys.stderr)
        return 2
    if shutil.which("pandoc") is None:
        print("找不到 pandoc，无法生成 Word 正文", file=sys.stderr)
        return 2
    if not (args.templates / "封面" / "封面-毕业论文.docx").is_file():
        print(f"模板不完整：{args.templates}", file=sys.stderr)
        return 2

    meta, body = T.read_markdown(args.input)
    info = {
        k: meta.get(k, "")
        for k in [
            "title",
            "college",
            "grade",
            "major",
            "class_name",
            "student_id",
            "name",
            "advisor",
            "abstract_cn",
            "keywords_cn",
            "abstract_en",
            "keywords_en",
            "date",
            "year",
        ]
    }
    info["type"] = args.type or meta.get("type", "毕业论文")
    info["year"] = str(info["year"] or "2027")

    warnings = [f"缺少字段：{k}" for k in REQUIRED if not str(info.get(k, "")).strip()]
    if not str(info.get("abstract_en", "")).strip():
        warnings.append("没有英文摘要，Abstract 页会只有标题和关键词")

    project = args.out / safe_name(info["title"] or args.input.stem)
    if project.exists() and any(project.iterdir()):
        print(f"目标目录已存在且非空，未覆盖：{project}", file=sys.stderr)
        return 3

    dirs = {
        "origin": project / "00-原文件",
        "templates": project / "01-模板",
        "covers": project / "02-封面",
        "output": project / "03-输出",
        "forms": project / "04-附件材料",
        "test": project / "05-测试",
        "tmp": project / "05-测试" / "中间文件",
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)

    # 00 原文件：只复制，不改
    for item in args.manual.iterdir():
        if item.is_file():
            shutil.copy2(item, dirs["origin"] / item.name)
    shutil.copy2(args.input, dirs["origin"] / f"原稿-{args.input.name}")
    write_info_yaml(info, dirs["origin"] / "论文信息.yaml")
    (dirs["origin"] / "说明.md").write_text(
        "# 原文件（只读）\n\n"
        "- `附件4：……（2027届）.doc`：学校原手册，未做任何修改。\n"
        "- `手册-2027（Word转换版）.docx`：仅用于提取模板的转换副本。\n"
        "- `原稿-*.md`：你提供的论文源文件副本。\n"
        "- `论文信息.yaml`：本次生成使用的字段。\n",
        encoding="utf-8",
    )

    # 01 模板：整棵目录复制，含 封面/ 与 附件材料/ 子目录
    copy_tree(args.templates, dirs["templates"])
    generated: list[str] = []

    # 02 封面：已填写的正文封面与附件材料封面
    cover_tpl = args.templates / "封面" / f"封面-{info['type']}.docx"
    cover = T.build_cover(cover_tpl, info, dirs["covers"] / "正文封面.docx")
    generated.append(str(cover.relative_to(project)))
    form_cover = T.build_cover(
        args.templates / "封面" / "封面-附件材料.docx",
        info,
        dirs["covers"] / "附件材料封面.docx",
    )
    generated.append(str(form_cover.relative_to(project)))

    # 03 输出：分章文件与完整论文
    cn = T.build_single_abstract(info, dirs["output"] / "01-中文摘要.docx", cn=True)
    en = T.build_single_abstract(info, dirs["output"] / "02-英文摘要.docx", cn=False)
    generated += [str(cn.relative_to(project)), str(en.relative_to(project))]

    body_docx = T.build_body(body, dirs["tmp"] / "正文混排.docx")
    split_dir = dirs["tmp"] / "拆分"
    split_dir.mkdir(parents=True, exist_ok=True)
    parts = T.split_body(body_docx, split_dir)
    for i, name in enumerate(["正文", "参考文献", "附录"], start=3):
        if name in parts:
            target = dirs["output"] / f"{i:02d}-{name}.docx"
            shutil.move(str(parts[name]), target)
            generated.append(str(target.relative_to(project)))

    front = T.build_front(info, dirs["tmp"] / "前置（摘要+目录）.docx")
    full = T.build_full(cover, front, body_docx, dirs["output"] / "完整论文.docx")
    generated.append(str(full.relative_to(project)))

    # 04 附件材料：目录页 + 每张表一个子文件夹
    form_root = dirs["forms"]
    toc_dir = form_root / "00-附件材料目录"
    toc_dir.mkdir(exist_ok=True)
    shutil.copy2(
        args.templates / "样式参考" / "附件材料目录.docx",
        toc_dir / "附件材料目录.docx",
    )
    generated.append(str((toc_dir / "附件材料目录.docx").relative_to(project)))

    for folder, filename in FORM_FOLDERS:
        d = form_root / folder
        d.mkdir(exist_ok=True)
        filled = T.fill_form(args.templates / "附件材料" / filename, info, d / filename)
        generated.append(str(filled.relative_to(project)))

    for folder, note in (
        ("12-查重检测报告", "查重检测报告由学校指定系统导出，模板以系统为准；本文件夹留空占位。"),
        ("13-AIGC检测报告", "AIGC 检测报告由学校指定系统导出，模板以系统为准；本文件夹留空占位。"),
    ):
        d = form_root / folder
        d.mkdir(exist_ok=True)
        (d / "说明.md").write_text(f"# {folder}\n\n{note}\n", encoding="utf-8")

    # 可选：正文 LaTeX
    if args.latex:
        tex = dirs["output"] / "正文.tex"
        md = dirs["tmp"] / "正文.md"
        md.write_text(body, encoding="utf-8")
        template = ASSETS / "body-template.tex"
        fix_tex = SKILL_DIR / "scripts" / "fix_tex.py"
        try:
            subprocess.run(
                ["pandoc", str(md), "-o", str(tex),
                 "-f", "markdown+tex_math_single_backslash",
                 "--template", str(template), "--standalone"],
                check=True,
            )
            subprocess.run([sys.executable, str(fix_tex), str(tex)], check=True)
            generated.append(str(tex.relative_to(project)))
        except subprocess.CalledProcessError:
            warnings.append("LaTeX 生成失败，已跳过（正文 docx 不受影响）")

    # 05 测试：格式自检、完整性检查、测试说明
    report = dirs["test"] / "格式自检.md"
    report.write_text(build_report(info, project, generated, warnings), encoding="utf-8")
    checks = collect_checks(project, info)
    write_test_files(project, info, checks, warnings)

    print(f"OK   {project}")
    for w in warnings:
        print(f"提醒 {w}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
