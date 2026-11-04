#!/usr/bin/env python3
"""把 Markdown 正文按重财格式输出 DOCX / PDF（format-only）。

不添加封面、摘要或固定章节；只应用字体、行距、标题、三线表、页码。
--toc 时插入可选目录（只到二级）。
"""
from __future__ import annotations

import argparse
import os
from pathlib import Path
import subprocess
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parent))

import format_docx  # noqa: E402
from build_word import build_reference  # noqa: E402

SCRIPT_DIR = Path(__file__).resolve().parent
BODY_TEMPLATE = SCRIPT_DIR.parent / "assets" / "body-template.tex"
FIX_TEX = SCRIPT_DIR / "fix_tex.py"


def strip_front_matter(text: str) -> str:
    if text.startswith("---"):
        parts = text.split("---", 2)
        if len(parts) == 3:
            return parts[2].lstrip()
    return text


def build_docx(src: Path, dst: Path, *, toc: bool, page_numbers: bool,
               three_line_tables: bool = True) -> Path:
    text = strip_front_matter(src.read_text(encoding="utf-8"))
    cache = Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache")) / "devconfig"
    ref = cache / "thesis-reference.docx"
    build_reference(ref, cache)

    with tempfile.TemporaryDirectory(prefix="devconfig-format-md-") as td:
        md = Path(td) / "body.md"
        raw = Path(td) / "raw.docx"
        md.write_text(text, encoding="utf-8")
        cmd = [
            "pandoc",
            str(md),
            "-o",
            str(raw),
            "-f",
            "markdown+tex_math_single_backslash",
            "--reference-doc",
            str(ref),
        ]
        if toc:
            cmd += ["--toc", "--toc-depth=2", "-M", "toc-title=目  录"]
        subprocess.run(cmd, check=True)
        format_docx.format_docx(
            raw,
            dst,
            toc=False,
            page_numbers=page_numbers,
            three_line_tables=three_line_tables,
        )
    return dst


def build_pdf(src: Path, dst: Path, *, toc: bool, page_numbers: bool = True) -> Path:
    text = strip_front_matter(src.read_text(encoding="utf-8"))
    dst.parent.mkdir(parents=True, exist_ok=True)
    tex = dst.with_suffix(".tex")

    with tempfile.TemporaryDirectory(prefix="devconfig-format-pdf-") as td:
        md = Path(td) / "body.md"
        md.write_text(text, encoding="utf-8")
        cmd = [
            "pandoc",
            str(md),
            "-o",
            str(tex),
            "-f",
            "markdown+tex_math_single_backslash",
            "--template",
            str(BODY_TEMPLATE),
            "--standalone",
        ]
        if toc:
            cmd += ["--toc", "--toc-depth=2", "-M", "toc-title=目  录"]
        if page_numbers:
            cmd += ["-M", "page_numbers=true"]
        subprocess.run(cmd, check=True)
        subprocess.run([sys.executable, str(FIX_TEX), str(tex)], check=True)

    if shutil_which("latexmk"):
        subprocess.run(
            ["latexmk", "-xelatex", "-interaction=nonstopmode", "-halt-on-error", tex.name],
            cwd=dst.parent,
            check=True,
        )
    else:
        for _ in range(2):
            subprocess.run(
                ["xelatex", "-interaction=nonstopmode", "-halt-on-error", tex.name],
                cwd=dst.parent,
                check=True,
            )
    if not dst.is_file():
        raise RuntimeError(f"PDF 未生成：{dst}")
    return dst


def shutil_which(cmd: str) -> bool:
    from shutil import which
    return which(cmd) is not None


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="按重财格式输出 Markdown（format-only）")
    ap.add_argument("input", type=Path)
    ap.add_argument("-o", "--output", type=Path)
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--toc", action="store_true")
    ap.add_argument("--no-page-numbers", action="store_true")
    ap.add_argument("--no-three-line-tables", action="store_true")
    args = ap.parse_args(argv)

    if not args.input.is_file():
        print(f"输入不存在：{args.input}", file=sys.stderr)
        return 2
    if args.pdf:
        out = args.output or args.input.with_suffix(".pdf")
        build_pdf(args.input, out, toc=args.toc, page_numbers=not args.no_page_numbers)
    else:
        out = args.output or args.input.with_name(args.input.stem + "_cfec.docx")
        build_docx(
            args.input,
            out,
            toc=args.toc,
            page_numbers=not args.no_page_numbers,
            three_line_tables=not args.no_three_line_tables,
        )
    print(f"OK   {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
