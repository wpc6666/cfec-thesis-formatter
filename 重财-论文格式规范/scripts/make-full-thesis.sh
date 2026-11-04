#!/usr/bin/env bash
# ============================================================
# make-full-thesis.sh — 从 Markdown 生成 Word / LaTeX / PDF
#
# 用法：
#   bash make-full-thesis.sh --word  输入.md [输出.docx]
#   bash make-full-thesis.sh --latex 输入.md [输出.tex]
#   bash make-full-thesis.sh --pdf   输入.md [输出.pdf]
#   bash make-full-thesis.sh --all   输入.md [输出基名]
#
# 依赖：
#   pandoc、python3（PyYAML/lxml）、xelatex + latexmk（PDF）
# ============================================================
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd -P)"
TEMPLATE="${SCRIPT_DIR}/../assets/thesis-template.tex"
BUILD_WORD="${SCRIPT_DIR}/build_word.py"
FIX_TEX="${SCRIPT_DIR}/fix_tex.py"

ok()   { printf 'OK   %s\n' "$*"; }
info() { printf -- '->   %s\n' "$*"; }
err()  { printf 'ERR  %s\n' "$*" >&2; }
die()  { err "$*"; exit 1; }

usage() {
  cat <<'EOF'
make-full-thesis.sh — 重庆财经学院 2027 届论文格式生成器

用法：
  bash make-full-thesis.sh --word  输入.md [输出.docx]
  bash make-full-thesis.sh --latex 输入.md [输出.tex]
  bash make-full-thesis.sh --pdf   输入.md [输出.pdf]
  bash make-full-thesis.sh --all   输入.md [输出基名]
  bash make-full-thesis.sh --help

输入：Markdown + YAML 元数据（见 example.md）
规范：README.md / format.yaml
EOF
}

need() {
  command -v "$1" >/dev/null 2>&1 || die "未找到 $1"
}

check_common() {
  need pandoc
  need python3
  python3 -c 'import yaml' >/dev/null 2>&1 || die "python3 缺 PyYAML"
}

check_pdf() {
  command -v latexmk >/dev/null 2>&1 || command -v xelatex >/dev/null 2>&1 \
    || die "PDF 需要 xelatex 或 latexmk"
}

build_latex() {
  local input="$1" output="$2"
  check_common
  info "pandoc ${input} -> ${output}"
  pandoc "$input" -o "$output" \
    -f markdown+tex_math_single_backslash \
    --template="$TEMPLATE" \
    --standalone
  python3 "$FIX_TEX" "$output"
  ok "LaTeX 生成：${output}"
}

build_pdf() {
  local input="$1" output="$2"
  check_common
  check_pdf
  local dir base tex
  dir="$(cd "$(dirname "$output")" && pwd -P)"
  base="$(basename "${output%.pdf}")"
  tex="${dir}/${base}.tex"
  build_latex "$input" "$tex"
  info "编译 PDF（自动跑两遍目录）"
  if command -v latexmk >/dev/null 2>&1; then
    (cd "$dir" && latexmk -xelatex -interaction=nonstopmode -halt-on-error "${base}.tex" >/dev/null) \
      || die "latexmk 编译失败：${tex}"
  else
    (cd "$dir" && xelatex -interaction=nonstopmode -halt-on-error "${base}.tex" >/dev/null) \
      || die "xelatex 第一遍失败：${tex}"
    (cd "$dir" && xelatex -interaction=nonstopmode -halt-on-error "${base}.tex" >/dev/null) \
      || die "xelatex 第二遍失败：${tex}"
  fi
  [[ -f "$output" ]] || die "PDF 未生成：${output}"
  ok "PDF 生成：${output}"
}

build_word() {
  local input="$1" output="$2"
  check_common
  python3 -c 'import lxml' >/dev/null 2>&1 || die "python3 缺 lxml"
  info "生成 Word：${input} -> ${output}"
  python3 "$BUILD_WORD" "$input" "$output"
  ok "Word 生成：${output}"
}

MODE=""
INPUT=""
OUTPUT=""

while [[ $# -gt 0 ]]; do
  case "$1" in
    --word) MODE=word; shift ;;
    --latex) MODE=latex; shift ;;
    --pdf) MODE=pdf; shift ;;
    --all) MODE=all; shift ;;
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

[[ -n "$MODE" && -n "$INPUT" ]] || { usage >&2; exit 2; }
[[ -f "$INPUT" ]] || die "输入不存在：${INPUT}"

stem="$(basename "${INPUT%.*}")"
dir="$(cd "$(dirname "$INPUT")" && pwd -P)"

case "$MODE" in
  word)
    build_word "$INPUT" "${OUTPUT:-${dir}/${stem}.docx}"
    ;;
  latex)
    build_latex "$INPUT" "${OUTPUT:-${dir}/${stem}.tex}"
    ;;
  pdf)
    build_pdf "$INPUT" "${OUTPUT:-${dir}/${stem}.pdf}"
    ;;
  all)
    base="${OUTPUT:-${dir}/${stem}}"
    build_word "$INPUT" "${base}.docx"
    build_latex "$INPUT" "${base}.tex"
    build_pdf "$INPUT" "${base}.pdf"
    ;;
esac
