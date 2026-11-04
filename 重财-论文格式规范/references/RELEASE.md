# 发布流程

> 上层：`../SKILL.md`；按需读取本文件。

## Skill 自身更新与发布流程【强制】· v2.2.0

当用户直接要求「更新 / 升级这个 skill（脚本）」时，按以下标准流程执行。**目标：让 agent 自动定位 GitHub 仓库与当前版本，完成一次可追溯的发布，无需用户提供仓库地址或版本号。**

### 1. 自动定位仓库与版本

- **GitHub 仓库**：`https://github.com/wpc6666/cfec-thesis-formatter`（main 分支 + 语义化 tag `vX.Y.Z`）。本地 clone 目录：`/tmp/cfec-thesis-formatter`（不存在则 `git clone`）。
- **本地 skill 目录**：`<skill目录>/`（Codex / 中央 Skill 目录安装；其他平台按实际安装路径调整，仍与仓库 `重财-论文格式规范/` 一一对应）。
- **当前版本**：以 `manifest.json` 的 `version` 字段为**单一权威来源**，须与 SKILL.md frontmatter、CHANGELOG 最新条目一致。
- 执行 `python3 scripts/release.py --check` 一次，即可同时拿到仓库、当前版本、最近 tag 与本地状态。

### 2. 版本递增（语义化版本）

- **patch**：修复 Bug（X.Y.Z → X.Y.Z+1）
- **minor**：新增向后兼容功能（X.Y.Z → X.Y+1.0）
- **major**：破坏性变更（X.Y.Z → X+1.0.0）

运行 `python3 scripts/release.py --bump <patch|minor|major>`，自动同步 **SKILL.md frontmatter 与标题、manifest.json、CHANGELOG 新条目、README 徽章、index.html**（后两者存在时）五处版本号；先用 `--dry-run` 预览，确认无误后正式写入，并补充 CHANGELOG 新条目的实际改动内容。

### 3. 本地验证（不允许未验证声称成功）

- 语法检查：`python3 -m py_compile scripts/*.py`
- 功能验证：实际运行受影响脚本（如 DOCX 引擎生成一份样本文档、release.py `--check`/`--bump` 往返），并检查产物关键点。
- 文档一致性：SKILL.md / manifest.json / CHANGELOG / README / index.html（存在时）版本号一致。

### 4. 同步本地与仓库 clone

```bash
LOCAL="${CFEC_SKILL_DIR:-$HOME/.agents/skills/cfec-thesis-formatter}"
REPO="${CFEC_REPO_CLONE:-/tmp/cfec-thesis-formatter}/重财-论文格式规范"
rm -rf "$REPO" && mkdir -p "$REPO"
cp -R "$LOCAL/." "$REPO/"
find "$REPO" -name "__pycache__" -type d -exec rm -rf {} +
diff -r "$LOCAL" "$REPO" --exclude="__pycache__"   # 必须输出「完全一致」
```

### 5. 提交、打 tag、推送（署名固定）

- 提交身份：`RELEASE_AUTHOR_NAME <RELEASE_AUTHOR_EMAIL>`；未设置时读本机 `git config`，两者都缺则拒绝提交。
- 提交信息：首行 `vX.Y.Z: <本次改动一句话>`，正文列出改动点。
- 打 tag `vX.Y.Z`，推送 `main` + tag；推送后 `git ls-remote --tags origin | grep vX.Y.Z` 确认。

### 6. 汇报

交付说明必须包含：本次版本号、改动内容、验证方式与结果、仓库提交与 tag 信息。

---
