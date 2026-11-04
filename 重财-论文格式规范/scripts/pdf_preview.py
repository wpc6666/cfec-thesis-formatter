#!/usr/bin/env python3
"""把 docx 导出成 PDF 预览：优先 Word，其次 soffice。

Word 路径会先检查 macOS 自动化授权；未授权时返回提示，
由调用方写入测试报告，并提示运行 authorize-word.sh。
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
WORD_APP = Path("/Applications/Microsoft Word.app")


def _run(cmd: list[str], timeout: int) -> subprocess.CompletedProcess:
    return subprocess.run(
        cmd, capture_output=True, text=True, timeout=timeout
    )


def export_pdf(docx: Path, pdf: Path, timeout: int = 180) -> tuple[bool, str]:
    docx = docx.resolve()
    pdf = pdf.resolve()
    if not docx.is_file():
        return False, f"输入不存在：{docx}"
    pdf.parent.mkdir(parents=True, exist_ok=True)
    if pdf.is_file():
        pdf.unlink()

    errors: list[str] = []

    if sys.platform == "darwin" and WORD_APP.is_dir():
        script = HERE / "word-export.sh"
        if script.is_file():
            try:
                result = _run(["bash", str(script), str(docx), str(pdf)], timeout)
                if result.returncode == 0 and pdf.is_file() and pdf.stat().st_size > 0:
                    return True, "Word"
                tail = (result.stderr or result.stdout).strip().splitlines()
                errors.append(tail[-1] if tail else "Word 导出失败")
            except subprocess.TimeoutExpired:
                errors.append("Word 导出超时（可能停在授权弹窗）")
            except OSError as exc:
                errors.append(f"Word 调用失败：{exc}")

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            result = _run(
                [
                    soffice,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(pdf.parent),
                    str(docx),
                ],
                timeout,
            )
            produced = pdf.parent / f"{docx.stem}.pdf"
            if result.returncode == 0 and produced.is_file():
                if produced.resolve() != pdf:
                    shutil.move(str(produced), str(pdf))
                if pdf.is_file() and pdf.stat().st_size > 0:
                    return True, "soffice"
            errors.append("soffice 导出失败")
        except subprocess.TimeoutExpired:
            errors.append("soffice 导出超时")
        except OSError as exc:
            errors.append(f"soffice 调用失败：{exc}")

    if not errors:
        errors.append(
            "未找到可用的 Word/soffice；如用 Word，请先运行 authorize-word.sh 授权"
        )
    return False, "；".join(errors)
