#!/usr/bin/env bash
# ============================================================
# make-thesis.sh — 按重财格式调整已有文档（format-only）
#
# 默认：只改字体/字号/行距/标题/三线表/页码，不生成封面、摘要或固定章节。
#
# 用法：
#   bash make-thesis.sh --docx 输入.docx [输出.docx]
#   bash make-thesis.sh --md   输入.md   [输出.docx]
#   bash make-thesis.sh --pdf  输入.md   [输出.pdf]
#   bash make-thesis.sh --full ...       旧的整篇论文生成器（可选）
#
# 可选：--toc / --no-page-numbers / --no-three-line-tables
# ============================================================
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd -P)"
VENV_DIR="${CFEC_VENV:-${SCRIPT_DIR}/.venv}"
PY=""

ok()   { printf 'OK   %s\n' "$*"; }
info() { printf -- '->   %s\n' "$*"; }
err()  { printf 'ERR  %s\n' "$*" >&2; }
die()  { err "$*"; exit 1; }

usage() {
  cat <<'EOF'
make-thesis.sh — 重庆财经学院格式器（只调格式，不改内容）

用法：
  bash make-thesis.sh --docx 输入.docx [输出.docx]
  bash make-thesis.sh --md   输入.md   [输出.docx]
  bash make-thesis.sh --pdf  输入.md   [输出.pdf]
  bash make-thesis.sh --full ...       旧的整篇论文生成器（可选）

选项：
  --toc                    插入/更新目录（只到二级）
  --no-page-numbers        不加页码
  --no-three-line-tables   不转换三线表
  --page-number-format      decimal（默认）或 upperRoman
  --help

说明：
  - 不会生成封面、摘要、参考文献或固定章节；只处理已有内容。
  - DOCX 输出默认加 _cfec 后缀，不覆盖原文件。
  - 首次使用自动创建 .venv 并安装 python-docx / PyYAML / lxml。
EOF
}

ensure_python() {
  if [[ -x "${VENV_DIR}/bin/python" ]]; then
    PY="${VENV_DIR}/bin/python"
  else
    info "首次使用：创建隔离 venv 并安装 python-docx/PyYAML/lxml"
    python3 -m venv "${VENV_DIR}" || die "创建 venv 失败：${VENV_DIR}"
    "${VENV_DIR}/bin/pip" install -q -r "${SCRIPT_DIR}/scripts/requirements.txt" \
      || die "安装依赖失败；检查网络后重试"
    PY="${VENV_DIR}/bin/python"
  fi
  "$PY" -c 'import docx, yaml, lxml' >/dev/null 2>&1 \
    || die "依赖不完整：$PY -m pip install -r '${SCRIPT_DIR}/scripts/requirements.txt'"
}

if [[ "${1:-}" == "--full" ]]; then
  shift
  exec bash "${SCRIPT_DIR}/scripts/make-full-thesis.sh" "$@"
fi

MODE=""
INPUT=""
OUTPUT=""
TOC=0
PAGE_NUMBERS=1
THREE_LINE=1
PAGE_FORMAT="decimal"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --docx) MODE=docx; shift ;;
    --md) MODE=md; shift ;;
    --pdf) MODE=pdf; shift ;;
    --toc) TOC=1; shift ;;
    --no-page-numbers) PAGE_NUMBERS=0; shift ;;
    --no-three-line-tables) THREE_LINE=0; shift ;;
    --page-number-format) shift; PAGE_FORMAT="${1:-decimal}"; shift ;;
    -o|--output) shift; OUTPUT="${1:-}"; shift ;;
    --help|-h) usage; exit 0 ;;
    -*) die "未知参数：$1（--help 查看用法）" ;;
    *)
      if [[ -z "$INPUT" ]]; then INPUT="$1"
      elif [[ -z "$OUTPUT" ]]; then OUTPUT="$1"
      else die "多余参数：$1"
      fi
      shift ;;
  esac
done

[[ -n "$INPUT" ]] || { usage >&2; exit 2; }
[[ -f "$INPUT" ]] || die "输入不存在：${INPUT}"

if [[ -z "$MODE" ]]; then
  case "$INPUT" in
    *.docx) MODE=docx ;;
    *.md|*.markdown) MODE=md ;;
    *) die "无法判断类型，请用 --docx / --md / --pdf" ;;
  esac
fi

ensure_python

case "$MODE" in
  docx)
    args=("$PY" "${SCRIPT_DIR}/scripts/format_docx.py" "$INPUT")
    [[ -n "$OUTPUT" ]] && args+=(-o "$OUTPUT")
    (( TOC )) && args+=(--toc)
    (( PAGE_NUMBERS )) || args+=(--no-page-numbers)
    (( THREE_LINE )) || args+=(--no-three-line-tables)
    args+=(--page-number-format "$PAGE_FORMAT")
    info "format-only DOCX"
    "${args[@]}"
    ;;
  md|pdf)
    args=("$PY" "${SCRIPT_DIR}/scripts/format_md.py" "$INPUT")
    [[ -n "$OUTPUT" ]] && args+=(-o "$OUTPUT")
    (( TOC )) && args+=(--toc)
    (( PAGE_NUMBERS )) || args+=(--no-page-numbers)
    (( THREE_LINE )) || args+=(--no-three-line-tables)
    [[ "$MODE" == pdf ]] && args+=(--pdf)
    info "format-only ${MODE}"
    "${args[@]}"
    ;;
esac
