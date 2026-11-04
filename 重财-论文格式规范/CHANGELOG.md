# 更新日志

本项目遵循 [语义化版本](https://semver.org/lang/zh-CN/) 规范：`主版本.次版本.修订号`。

> **版本管理说明**：仓库内仅保留一个最新版 skill 目录（`重财-论文格式规范/`），历史版本不再以目录快照（如 `_v1.1.0/`）存放，改由 Git tag 与 CHANGELOG.md 记录。安装/更新请始终使用最新目录。

---

## [3.5.1] - 2026-11-04

### 修复

- 补充 MIT `LICENSE` 文件，修复 README 徽章链接 404。
- `manifest.json` 增加 `license: MIT`。

---

## [3.5.0] - 2026-11-04

### 新增

- `thesis-format` 工具完整迁入 Skill：`make-thesis.sh`、`build_word.py`、`format_md.py`、
  `ref_zh.py`、`make-full-thesis.sh`、`thesis-template.tex`、`example.md`。
  旧工具目录只保留兼容入口与软链接。
- `manifest.json` 增加 `repository` / `homepage`，新增 `.release.json` 钉死发布坐标。

### 修复

- `release.py` 去掉写死的仓库地址，改为 `.release.json` → `.git` → manifest → SKILL.md → 上级 `.git`；
  并明确它只做定位/版本递增，不负责 clone/push/tag。
- 发布前扫描补充 HOME 变量形式的私有路径写法。
- `references/RELEASE.md` 不再使用 HOME 变量形式的私有路径。

---

## [3.4.1] - 2026-11-04

### 新增

- `format-only.sh`：直接对已有 docx 刷格式，不生成封面。

### 修复

- 修正公开 README 的安装命令与 format-only 用法，移除不存在的 `install-skill.sh` 引用。

---

## [3.4.0] - 2026-11-04

### 新增

- Word 自动导出与授权：`authorize-word.sh` 主动触发 macOS 自动化授权并打开设置面板；
  `scripts/word-export.sh` 在 Word 容器内导出 PDF，避免反复弹文件访问权限。
- `05-测试` 自动生成 PDF 预览：优先 Microsoft Word，其次 soffice；
  未授权时跳过并写入 `提醒.md`，不阻塞生成。

---

## [3.3.0] - 2026-11-04

### 新增

- 课程报告生成器支持两种模式：`--template`（给模版，生成封面+正文+附件）与
  `--no-template`（不给模版，不生成封面/附件，不套用论文模板）。
- 两个模式必须二选一；都不给时脚本拒绝执行，并提示“课程模版 / 2027 论文模版 / 不要模板”。
- `SKILL.md` 增加触发条件：只有用户明确调用 `/cfec-thesis-formatter` 时才执行格式修改，
  普通文档任务不自动套模板。

### 修复

- 图片段落只调整行距，不再改对齐/缩进/环绕；修复 Word 里图片被裁切、看不完整的问题。
- 三线表逐格补首行顶线、表头底线、末行底线，并清除表头底纹。
- 「课程报告要求」不再放进 `01-模板`，只作为原文件/附件参考。

---

## [3.2.0] - 2026-11-04

### 新增

- 新增课程报告入口 `make-report-project.sh`：输入报告 docx + 课程格式模版（+ 课程要求），
  输出 `00-原文件 / 01-模板 / 02-封面 / 03-输出 / 04-附件材料 / 05-测试` 六级项目目录。
- 从模版切出封面与附件评价表；封面自动填课程名称、项目名称，成员信息可留空或通过 `--info` 补。
- 新增 `references/REPORT-PROJECT.md`、`scripts/make_report_project.py`、`scripts/format_docx.py`。

### 修复

- docx 合并时补齐嵌入对象（`.bin`）、WMF/EMF 图片的 Content-Type，避免 Word 提示修复。
- format-only 图片段落不再固定 26 磅；表格首行设置跨页重复表头。
- 三线表改为逐格补首行顶线、表头底线、末行底线，并清除表头底纹；
  修复单元格 nil 边框覆盖表格外框、看起来只剩一条线的问题。

---

## [3.1.0] - 2026-11-04

### 新增

- **Skill 自带模板与原文件**：`assets/2027/` 内含学校手册、3 个封面、11 张附件表、
  样式样例与专业班级对照表；换电脑只装 Skill + pandoc 即可完整生成，不再依赖外部论文工具目录。
- 新增 Skill 入口 `make-thesis-2027.sh`，首次运行自动建 `.venv` 并安装
  `python-docx` / `PyYAML` / `lxml`。
- 项目输出改为六级目录：`00-原文件`、`01-模板`、`02-封面`、`03-输出`、
  `04-附件材料`、`05-测试`。
- `05-测试/` 自动生成 `完整性检查.md`（逐项通过/缺失）、`格式自检.md`、`测试说明.md`；
  若系统有 `soffice`，额外生成完整论文 PDF 预览。

### 修复（format-only 格式器）

- 图片段落改用单倍行距，避免内嵌图被固定 26 磅裁成一行高。
- 表格首行设置跨页重复表头。

### 变更

- 生成器脚本与资源迁入 `scripts/2027/`、`assets/2027/`；`thesis-format/2027/` 保留软链接兼容旧路径。
- 文档统一使用 `<skill目录>` 表述，不再要求用户知道固定的外部工具目录。

---

## [3.0.0] - 2026-11-04

### 新增

- 对齐《重庆财经学院毕业论文（设计）指导手册（2027届）》：封面、摘要、目录、
  标题、图表、参考文献、附录参数全部按 2027 手册重核。
- 新增 `make-thesis-2027.sh` 一键生成项目：封面 → 摘要 → Abstract → 目录 →
  正文 → 参考文献 → 附录，并生成每表一个子文件夹的附件材料。
- 新增学校原版封面 3 个、附件表 11 张、样式样例 7 个、专业班级对照表 144 条，
  全部从手册按“节”原样提取，原文件只读。
- 新增 `references/GENERATOR-2027.md`、`references/FORMS-2027.md`。

### 修正

- 封面：题目四号加粗，学院/年级/专业/学号/姓名/指导教师四号**不加粗**。
- 图序表序：`图1-1` / `表1-1`，「图/表」与数字之间不留空格，题名与编号间空 2 个半角空格。
- 三级标题编号改为 `1. 2. 3.`（2027 手册样例为「1.页面设置」），不再用 `1.1`。
- 一级标题、参考文献标题、附录编号不再加粗；目录条目不带蓝色超链接。

### 依赖

- 生成器与模板位于 `<论文工具目录>/2027/`；Skill 只保留流程与规范。

---

## [2.2.9] - 2026-11-02

### 重构

- 将发布流程、输出规范、DOCX/LaTeX 工作流、验证和故障处理下沉到 `references/`。
- 主 `SKILL.md` 保留触发、核心原则和入口步骤，降低首屏上下文体积。

---

## [2.2.8] - 2026-11-02

### 修复

- 清理 `cfec_thesis_docx.py` 中未使用的导入（`json`、`Emu`、`nsmap`），
  通过 Ruff 零告警检查；不改变 DOCX 格式与输出行为。

---

## [2.2.7] - 2026-11-01

### 修复

- 中文摘要标题与目录标题改用两个半角空格（`摘  要` / `目  录`），
  不再使用一个全角空格，与学校格式规范和 `thesis-format` 引擎保持一致。

---

## [2.2.6] - 2026-11-01

### 修复

- 修正 Skill 自身发布流程中的本地路径：从旧的 `.user_skills/重财-论文格式规范`
  改为中央 Skill 目录 `<skill目录>`，并支持用
  `CFEC_SKILL_DIR` / `CFEC_REPO_CLONE` 覆盖。
- 发布说明改为与实际 `release.py` 一致的五处版本同步：
  SKILL.md、manifest.json、CHANGELOG.md、README 徽章、index.html（存在时）。

---

## [2.2.5] - 2026-10-20

### 修复

- `_patch_changelog()` 现在幂等：CHANGELOG 里已有该版本号的条目就跳过，
  不再重复插入「待补充」空模板。
  此前无条件插入，导致 v2.2.1 / v2.2.3 / v2.2.4 各出现两条——
  一条是真内容、一条是空模板。
- 清理 CHANGELOG 里 3 条重复的空模板条目（v2.2.1 / v2.2.3 / v2.2.4）。
- 删除从未发布的 v2.2.2 空条目，并在 v2.2.3 处注明版本号跳过的原因。
- 修正 `apply_bump()` 里 `_patch_changelog()` 返回值的解包索引
  （2 元组被当成 3 元组，会抛 `IndexError: tuple index out of range`）。

---

## [2.2.4] - 2026-10-19

### 修复

- `release.py` 同步 README 的 Version 徽章（`badge/Version-X.Y.Z-`）。
  正则只匹配 shields.io 徽章那一处，正文里的历史版本号不动。
  至此发版会同步五处：SKILL.md / manifest.json / CHANGELOG.md /
  index.html / README 徽章。

## [2.2.3] - 2026-10-19

### 修复

- `release.py` 的发版流程现在会一并更新 `index.html` 的版本号。
  此前只改 SKILL.md / manifest.json / CHANGELOG.md 三处，
  导致 index.html 长期漂在 v2.0.0（v2.2.0 现场）。
  index.html 在仓库根、不在 skill 子目录，新增 `_PATH_MAP` 记录真实路径，
  并补上缺失的 `import io`。

> **版本号说明**：`v2.2.2` 未实际发布（无对应提交，仅有一次误建的 CHANGELOG 空条目），
> 版本号从 `v2.2.1` 直接跳到 `v2.2.3`。

## [2.2.1] - 2026-10-19

### 修复

- `manifest.json` 的 `scripts` 数组漏登记 `./scripts/release.py`，
  导致 agent 读 manifest 时看不到发布脚本。
- `index.html` 页面版本号仍是 v2.0.0，与 SKILL.md / manifest.json 的 v2.2.0 不一致。
- `README.md` 的三处徽章链接指向 `重财-论文格式规范/CHANGELOG.md`，
  但 `README.md` 本身在仓库根，该相对路径在 GitHub 上 404。
- 顺带用 self-release v1.8.0 修好的坐标自愈逻辑重新校验了
  `.release.json` 的 `subdir`（应为 `重财-论文格式规范`）。

## [2.2.0] - 2026-10-19

### 新增：Skill 自身发布自动化（agent 自动定位 GitHub 仓库与版本）

- **`scripts/release.py`**（纯标准库，新增）：
  - 自动定位 GitHub 仓库 `wpc6666/cfec-thesis-formatter` 与本地 skill 目录、当前版本
  - `--check`：一键显示仓库、当前版本（manifest.json 单一权威来源）、最近 tag、本地状态
  - `--bump <patch|minor|major>`：按语义化版本递增，自动同步 **SKILL.md frontmatter/标题、manifest.json、CHANGELOG** 三处版本号
  - `--dry-run`：只预览将做的改动，不写文件
- **SKILL.md 新增「Skill 自身更新与发布流程【强制】」章节**：
  - 用户直接要求「更新这个 skill/脚本」时的标准发布流程：定位仓库 → 定位版本 → 版本递增 → 本地验证 → 同步本地↔仓库 clone（`diff -r` 一致）→ 提交（署名白澤 程）→ 打 tag → 推送 main + tag → 汇报

---

## [2.1.0] - 2026-10-19

### 新增：DOCX 域自动更新 + 类型子文件夹输出

- **`scripts/cfec_thesis_docx.py`**：
  - 新增 `_enable_auto_update_fields()`：保存前向 `settings.xml` 注入 `<w:updateFields w:val="true"/>`，Word 打开文档时**自动更新全部域（目录 TOC、页码 PAGE）**，不再要求用户手动 `Ctrl+A → F9`；`save()` 与 `build_and_save()` 均自动生效
  - 目录占位文本同步改为「打开文档时将自动更新目录/页码域」
- **`scripts/output_manager.py`（v2.1.0 输出结构）**：
  - 版本目录 `v<N>/` 内按类型分子文件夹：`word/`（DOCX）、`latex/`（TEX + PDF）、`data/`（JSON）
  - 文件名仍携带版本号：`word/<题目slug>_v<N>.docx`、`latex/<题目slug>_v<N>.tex` 等（命名规则见 SKILL.md）
  - `path_for()/output_paths()/tree()/write_manifest()` 全部适配子文件夹结构
- **SKILL.md 升级 v2.1.0**：
  - 「输出目录树规范【强制】」更新为子文件夹结构 + 完整命名规则表
  - 新增「域自动更新规则」：DOCX 打开自动更新域 ↔ LaTeX 编译两次，目录/页码两侧同步
  - 验证清单与交付说明同步（不再要求手动 F9）
- **references/latex-guide.md**：新增「目录/页码与 DOCX 同步」说明

---

## [2.0.1] - 2026-10-19

### 修复：标题/计数器编号 Bug 与 LaTeX 模板编译错误

- **`scripts/cfec_thesis_docx.py`**：`add_h1/add_h2/add_h3` 现在会重置子级标题计数器（`_h2/_h3/_h4`），每章二级标题（一）（二）…、三级标题 1. 2. … 编号按章节归零，与 LaTeX 模板行为一致。
- **`assets/chongqing-caijing-thesis.tex`（修复三处编译错误，模板此前无法通过 XeLaTeX 编译）**：
  1. 计数器/宏名含数字导致分词错误：`cnt1–cnt4` → `cntA–cntD`，`\h1–\h4` → `\hA–\hD`（TeX 控制序列名只能由字母组成）
  2. caption 包报 `heiti/zihao undefined`：新增 `\DeclareCaptionFont{cfec}{\heiti\zihao{5}}` 声明
  3. 公式编号缺失章节号：新增 `\renewcommand{\theequation}{\arabic{chapternum}.\arabic{equation}}`（与 DOCX 引擎 （1.1） 编号一致）
- **SKILL.md / references/latex-guide.md**：同步更新为 `\hA–\hD` 命令名。

---

## [2.0.0] - 2026-10-19

### 重大修复：输出版本管理重构（修复同名覆盖 Bug）

历史问题：输出文件名固定为「原文件名_重财规范版.docx / .tex」，同一论文重复格式化会互相覆盖，且无项目隔离、无版本记录。

本次变更：

- **新增 `scripts/output_manager.py`（输出版本管理器，纯标准库）**：
  - 按论文题目建项目目录（slug 化，≤40 字符）
  - 版本号自动递增：扫描 `v1…vN` 取最大 +1，第 1 次 `v1`、第 2 次 `v2`，**绝不覆盖历史输出**
  - 支持 `--plan` 预览、`--version N` 指定版本、`--list` 输出目录树
  - 项目根目录自动生成 `MANIFEST.md` 版本清单（各版本生成时间/文件/说明）
  - `tree()` 输出完整目录树

- **`scripts/cfec_thesis_docx.py` 集成版本化输出**：
  - 新增 `build_and_save(thesis_data, root_dir, project, force_version)` 推荐入口
  - 移除硬编码文件名（原 `save("重财论文示例_v1.1.1.docx")`）
  - `__main__` 演示连续两次运行生成 `v1/`、`v2/` 两个独立版本

- **SKILL.md 更新**：
  - 版本号升级为 2.0.0
  - 新增「输出目录树规范【强制】」章节：目录树、命名规则表、版本号规则
  - Step 1 增加输出版本规划（`output_manager.py --plan/--list`）
  - Step 3 改走 `build_and_save()`，Step 8 交付必须附目录树与 MANIFEST
  - 错误处理新增「目标版本目录已存在」场景
  - 允许与禁止新增「不得输出到无版本号的固定文件名」「不得覆盖非空版本目录」

- **manifest.json 更新**：版本号 2.0.0，scripts 注册新增 `output_manager.py`，tags 新增「版本管理」，release_date 2026-10-19。

- **仓库结构优化**：删除 `重财-论文格式规范_v1.1.0/`、`重财-论文格式规范_v1.1.1/` 目录快照，统一为单一 skill 目录；marketplace.json 仅注册最新插件，历史版本由 Git tag 与 CHANGELOG 记录。

---

## [1.1.1] - 2026-09-02

### 重大新增

- **python-docx 原生 Word 生成引擎**（`scripts/cfec_thesis_docx.py`，869行）：
  - 直接输出二进制 `.docx` 文件，不再依赖 pandoc 中转
  - 页面设置：A4，左2.8cm，其余2.5cm，页眉1.5cm，页脚1.75cm
  - 全文固定行距 26 磅（`WD_LINE_SPACING.EXACTLY`）
  - 分节页码：封面无码、前置部分大写罗马数字、正文阿拉伯数字从1开始
  - 四级标题体系：一、→（一）→1.→（1），自动中文数字编号
  - 图 1-1 / 表 1-1 按章节自动编号
  - 三线表：仅顶线、表头线、底线，无竖线
  - 公式（1.1）右对齐按章节编号
  - 参考文献悬挂缩进 2 字符
  - 正文引用 `[1]` 自动识别并设为上标
  - 目录域（TOC \o "1-2"），Word 中 F9 更新
  - 中英文字体分离：中文宋体/黑体，西文 Times New Roman
  - 完整 `ThesisFormatter` 类，支持 `build(thesis_data)` 一键生成
  - 内置命令行示例，可直接运行生成示例论文

### SKILL.md 更新

- 版本号升级为 1.1.1
- 新增 Step 2：将论文内容解析为结构化数据 `thesis_data` 字典
- 新增 Step 3：调用 Python 引擎生成 DOCX（含完整调用示例）
- 新增 Python 引擎自动完成的格式设置对照表（13项）
- 新增 Step 4：DOCX 后处理与验证（含 Ctrl+A→F9 更新域说明）
- 新增 v1.1.1 重点验证项（11项，含页边距、行距、分节页码、三线表等）
- 新增交付说明：需告知用户 DOCX 打开后按 F9 更新目录域
- 新增 thesis_data.json 可选交付物
- 核心原则新增"原生生成优先"
- 错误处理新增 python-docx 未安装、生成失败的回退方案
- 允许与禁止列表更新

### manifest.json 更新

- 版本号 1.1.0 → 1.1.1
- 新增 `scripts` 字段，注册 `./scripts/cfec_thesis_docx.py`
- tags 新增 "python-docx"
- description 更新

---

## [1.1.0] - 2026-09-02

### 新增

- **manifest.json**：新增 Skill 元信息配置文件，支持豆包 Agent 覆盖安装识别。
- **Step 0 前置检查**：新增输入完整性检查、环境检查、格式规范加载。
- **错误处理与回退机制**：新增常见错误场景对照表及降级原则。
- **增强版自检清单**：DOCX 和 LaTeX 各新增 8+ 项检查点。
- **交付说明规范**：明确交付时必须包含的 5 项内容。
- **CHANGELOG.md**：新增版本更新日志文件。

### 优化

- SKILL.md 头部元信息新增 `version` 字段。
- 工作流步骤编号统一，从 Step 0 开始。
- 核心原则新增"可追溯"原则。

---

## [1.0.0] - 初始版本

### 功能

- 重庆财经学院本科毕业论文格式规范完整实现。
- 双版本输出：DOCX（Word）+ LaTeX（XeLaTeX）。
- 四级标题体系：一、 →（一）→ 1. →（1）。
- 页码分节：封面无码、前置罗马数字、正文阿拉伯数字。
- 图表编号：图 1-1 / 表 1-1，按章节编号。
- 公式编号：（1.1）右对齐，按章节编号。
- 参考文献：悬挂缩进 2 字符，正文引用上标。
- 目录：只显示到二级标题，点号前导符。
- LaTeX 模板：基于 ctexart + XeLaTeX，内置自定义标题命令。
- 格式规范速查表：references/format-spec.md。
- LaTeX 使用指南：references/latex-guide.md。

---

## 版本说明

- **主版本号**：不兼容的 API 或格式规范重大变更时递增。
- **次版本号**：向下兼容的功能性新增时递增。
- **修订号**：向下兼容的问题修正时递增。

### 覆盖安装说明

本 Skill 支持覆盖安装（`update_mode: overwrite`）。安装 v1.1.1 时：

1. 豆包 Agent 读取 manifest.json，识别 `id: cfec-thesis-formatter`。
2. 检测到已安装同 ID 旧版本，自动覆盖 SKILL.md、scripts/ 及配套资源。
3. 版本号更新为 1.1.1，旧版本被替换。
4. 重启 Agent 服务后生效。

> 注意：覆盖安装不会删除用户自定义的论文文件，仅更新 Skill 本身。
> v1.1.1 新增 `scripts/` 目录，覆盖安装时会自动创建。
