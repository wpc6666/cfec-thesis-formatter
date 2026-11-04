#!/usr/bin/env bash
# ============================================================
# format-only.sh — 只给已有 docx 刷格式，不生成封面
#
# 用法：
#   bash <skill目录>/format-only.sh 输入.docx [-o 输出.docx]
#        [--toc] [--no-page-numbers] [--no-three-line-tables]
#        [--page-number-format decimal|upperRoman]
# ============================================================
set -euo pipefail

SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd -P)"
VENV="${CFEC_VENV:-${SKILL_DIR}/.venv}"
PY="${VENV}/bin/python"

if [[ ! -x "${PY}" ]]; then
  printf -- '->   首次运行，创建隔离环境 %s\n' "${VENV}"
  python3 -m venv "${VENV}"
  "${VENV}/bin/pip" install -q -r "${SKILL_DIR}/scripts/requirements.txt"
fi

exec "${PY}" "${SKILL_DIR}/scripts/format_docx.py" "$@"
