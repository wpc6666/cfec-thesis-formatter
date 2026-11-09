#!/usr/bin/env bash
# ============================================================
# make-thesis-2027.sh — 一键生成 2027 届重财毕业论文完整项目
#
# Skill 自带模板与原文件，换电脑只需安装本 Skill 和 pandoc。
#
# 用法：
#   bash <skill目录>/make-thesis-2027.sh 论文.md [--out 输出目录]
#        [--type 毕业论文|毕业设计] [--latex]
#
# 输出：<输出目录>/<题目>/
#   00-原文件  01-模板  02-封面  03-输出  04-附件材料  05-测试
# ============================================================
set -euo pipefail

SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd -P)"
VENV="${CFEC_VENV:-${SKILL_DIR}/.venv}"
PY="${VENV}/bin/python"

if [[ ! -x "${PY}" ]] || ! "${PY}" -c 'import docx, yaml, lxml' >/dev/null 2>&1; then
  printf -- '->   创建/修复隔离环境 %s\n' "${VENV}"
  [[ -x "${PY}" ]] || python3 -m venv "${VENV}"
  "${VENV}/bin/pip" install -q -r "${SKILL_DIR}/scripts/requirements.txt"
fi

if ! command -v pandoc >/dev/null 2>&1; then
  printf 'ERR  找不到 pandoc。\n' >&2
  printf '     macOS: brew install pandoc\n' >&2
  printf '     Ubuntu/Debian: sudo apt install pandoc\n' >&2
  exit 2
fi

exec "${PY}" "${SKILL_DIR}/scripts/2027/make_project.py" "$@"
