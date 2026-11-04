#!/usr/bin/env bash
# ============================================================
# authorize-word.sh — 一次性授权 Word 自动化（macOS）
#
# macOS 的“自动化”权限属于 TCC，脚本不能静默勾选；
# 本脚本会主动触发一次系统授权弹窗，并在失败时打开设置面板。
#
# 用法：
#   bash <skill目录>/authorize-word.sh
# ============================================================
set -euo pipefail

SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd -P)"
EXPORT="${SKILL_DIR}/scripts/word-export.sh"

if [[ "$(uname -s)" != "Darwin" ]]; then
  printf '这不是 macOS，不需要 AppleEvents 授权。\n'
  exit 0
fi

if [[ ! -d "/Applications/Microsoft Word.app" ]]; then
  printf 'ERR  未安装 Microsoft Word，无法使用 Word 自动导出。\n' >&2
  printf '     可改用含 soffice 的办公套件导出 PDF。\n' >&2
  exit 2
fi

printf -- '->   触发 Word 自动化授权（首次会弹窗，请点“允许”）…\n'
if bash "${EXPORT}" --check; then
  printf 'OK   Word 自动化已授权，生成器可直接导出 PDF。\n'
  exit 0
fi

printf -- '->   未授权。正在打开「系统设置 → 隐私与安全性 → 自动化」…\n'
bash "${EXPORT}" --permissions || true
cat <<'EOF'

请手动完成一次：
  系统设置 → 隐私与安全性 → 自动化
  → 找到运行本脚本的 App（终端 / Claude Code / ChatGPT / Codex）
  → 勾选它下面的 Microsoft Word

勾选后重新运行：
  bash <skill目录>/authorize-word.sh

说明：macOS 不允许脚本静默写入 TCC 授权；这里只能触发弹窗并打开设置，
不能代替你点“允许”。授权一次后，后续导出 PDF 不再询问。
EOF
exit 1
