#!/usr/bin/env bash
# ============================================================
# word-export.sh — 用 Microsoft Word 无提示导出 PDF
#
# macOS 不允许脚本静默写入 TCC 授权；AppleEvents 权限需要用户一次性
# 在「系统设置 → 隐私与安全性 → 自动化」里允许运行脚本的 App（终端 / Claude Code / ChatGPT / Codex）控制 Word。
#
# 文件访问则不该反复弹窗：本脚本先把 docx 复制进 Word 自己的容器，
# 在容器内打开、导出 PDF，再用普通文件复制把 PDF 放回目标路径。
#
# 用法：
#   bash word-export.sh 输入.docx [输出.pdf]
#   bash word-export.sh --check
#   bash word-export.sh --permissions
# ============================================================
set -euo pipefail

WORD_APP="/Applications/Microsoft Word.app"
WORD_DATA="${HOME}/Library/Containers/com.microsoft.Word/Data"
WORK="${WORD_DATA}/Documents/CFECWord"

ok()   { printf 'OK   %s\n' "$*"; }
info() { printf -- '->   %s\n' "$*"; }
err()  { printf 'ERR  %s\n' "$*" >&2; }
die()  { err "$*"; exit 1; }

usage() {
  cat <<'EOF'
word-export.sh — 通过 Microsoft Word 导出 PDF

用法：
  bash word-export.sh 输入.docx [输出.pdf]
  bash word-export.sh --check
  bash word-export.sh --permissions
  bash word-export.sh --help

说明：
  - docx 会先复制到 Word 容器，Word 不需要访问 /tmp 或 iCloud 原路径。
  - AppleEvents 权限只能在系统设置里授权一次；脚本不会绕过 TCC。
  - 默认输出到输入文件同目录的 <文件名>.pdf。
EOF
}

check_word() {
  [[ -d "$WORD_APP" ]] || die "找不到 Microsoft Word.app"
  mkdir -p "$WORK"
  [[ -w "$WORK" ]] || die "Word 容器不可写：$WORK"
  local ver
  ver="$(osascript -e 'tell application "Microsoft Word" to get version' 2>&1)" \
    || die "AppleEvents 未授权：请在 系统设置 → 隐私与安全性 → 自动化 中允许运行脚本的 App（终端 / Claude Code / ChatGPT / Codex）控制 Microsoft Word。原始错误：$ver"
  ok "Word ${ver}；容器可写"
}

open_permissions() {
  info "打开「自动化」设置；请只授权一次“运行脚本的 App → Microsoft Word”"
  open "x-apple.systempreferences:com.apple.preference.security?Privacy_Automation" 2>/dev/null || true
  info "不需要授予「完全磁盘访问权限」；文件已放进 Word 自己的容器"
}

export_pdf() {
  local input="$1" output="$2"
  [[ -f "$input" ]] || die "输入不存在：$input"
  case "$input" in
    *.docx|*.doc) ;;
    *) die "只支持 .docx / .doc：$input" ;;
  esac

  check_word
  mkdir -p "$WORK"

  local base stem staged pdf
  base="$(basename "$input")"
  stem="${base%.*}"
  staged="${WORK}/${base}"
  pdf="${WORK}/${stem}-$$.pdf"
  cp "$input" "$staged" || die "复制 docx 到 Word 容器失败：$staged"

  info "Word 容器内打开：$staged"
  if ! osascript - "$staged" "$pdf" <<'APPLESCRIPT'
on run argv
  set inPath to item 1 of argv
  set outPath to item 2 of argv
  set targetName to name of (info for POSIX file inPath)
  tell application "Microsoft Word"
    activate
    open POSIX file inPath
    delay 2
    set d to active document
    if name of d is not targetName then
      set d to document targetName
    end if
    save as d file name outPath file format format PDF
    close d saving no
  end tell
end run
APPLESCRIPT
  then
    die "Word 导出失败；若系统弹出授权框，请允许一次后重试"
  fi

  [[ -s "$pdf" ]] || die "Word 未生成 PDF：$pdf"
  mkdir -p "$(dirname "$output")"
  cp "$pdf" "$output" || die "复制 PDF 到目标路径失败：$output"
  ok "已导出：$output"
}

case "${1:-}" in
  --check) check_word ;;
  --permissions) open_permissions ;;
  --help|-h) usage ;;
  "")
    usage >&2
    exit 2
    ;;
  -*)
    die "未知参数：$1（--help 查看用法）"
    ;;
  *)
    input="$1"
    if [[ -n "${2:-}" ]]; then
      output="$2"
    else
      output="${input%.*}.pdf"
    fi
    export_pdf "$input" "$output"
    ;;
esac
