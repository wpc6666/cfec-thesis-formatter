#!/usr/bin/env bash
# ============================================================
# make-report-project.sh — 课程报告 → 多级项目文件夹
#
# 用法：
#   bash <skill目录>/make-report-project.sh \
#     --report 报告.docx --template 格式模版.docx \
#     [--requirements 课程要求.docx] [--out 输出目录] [--info 信息.yaml]
#
#   bash <skill目录>/make-report-project.sh \
#     --report 报告.docx --no-template [--out 输出目录]
#
#   --template 与 --no-template 二选一；都不给会拒绝执行。
#
# 输出：<报告名>/
#   00-原文件 01-模板 02-封面 03-输出 04-附件材料 05-测试
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

exec "${PY}" "${SKILL_DIR}/scripts/make_report_project.py" "$@"
