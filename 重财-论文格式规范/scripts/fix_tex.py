#!/usr/bin/env python3
"""把 pandoc 生成的未编号参考文献/附录标题加入 LaTeX 目录。"""
from __future__ import annotations

import re
import sys
from pathlib import Path


def fix(text: str) -> str:
    def repl(match: re.Match[str]) -> str:
        title = match.group(1)
        marker = f"\\addcontentsline{{toc}}{{section}}{{{title}}}"
        if marker in text:
            return match.group(0)
        return (
            "\\clearpage\n"
            f"\\section*{{{title}}}\n"
            f"{marker}"
        )

    return re.sub(r"\\section\*\{([^}]*)\}", repl, text)


def main(argv: list[str]) -> int:
    if len(argv) != 2:
        sys.stderr.write("用法：python3 fix_tex.py <thesis.tex>\n")
        return 2
    path = Path(argv[1])
    path.write_text(fix(path.read_text(encoding="utf-8")), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
