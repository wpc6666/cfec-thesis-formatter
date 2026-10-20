---
name: 重财-论文格式规范
version: 2.2.5
description: 将中文学术论文（重庆财经学院本科毕业论文/设计）调整为符合该校《毕业论文（设计）指导手册》的格式规范，同时输出 DOCX（Word）和 LaTeX 两个版本。v2.2.0 新增「Skill 自身发布自动化」：agent 被要求更新本 skill 时，自动定位 GitHub 仓库（wpc6666/cfec-thesis-formatter）与当前版本，运行 scripts/release.py 一键完成版本递增（--check/--bump/--dry-run）与发布同步。v2.1.0 新增「域自动更新」与「类型子文件夹输出」：DOCX 引擎自动注入 updateFields，Word 打开即自动更新目录/页码域，无需手动 Ctrl+A→F9；版本目录 v<N>/ 内按类型分 word/、latex/、data/ 子文件夹，文件名仍携带迭代版本号（v1/v2/v3…），历史输出永不覆盖，目录树与 MANIFEST.md 版本清单可追溯。涵盖封面、摘要（中英文）、目录、正文标题、正文文本、图表、公式、参考文献、附录、页码、装订顺序及字体字号。本技能完全依据重庆财经学院指导手册，不采用 GB/T 7713.2-2022 标准。只处理格式，不改写论文内容。LaTeX 模板基于 ctexart + XeLaTeX，内置四级标题命令（一、→（一）→1.→（1））、图1-1/表1-1编号、公式（1.1）编号、罗马/阿拉伯分节页码、悬挂缩进参考文献。当用户要求以下任务时自动调用：按重庆财经学院毕业论文手册调整论文格式；同时生成Word和LaTeX版本；检查摘要、Abstract、目录、正文标题层级是否符合手册要求；检查字体字号（小三、四号、小四、五号等）是否正确；检查页码规则（前置部分罗马数字，正文阿拉伯数字）；检查行距是否为固定值26磅；检查页边距（左2.8cm，上/下/右均为2.5cm）；检查图序表序是否用"图 1-1""表 1-1"格式；检查参考文献引用是否为上标；生成或更新目录（仅显示到二级标题）；生成重财本科毕设LaTeX模板；使用python-docx生成原生Word文档。
---

# 重财-论文格式规范 v2.2.5
将重庆财经学院本科毕业论文调整为符合该校《毕业论文（设计）指导手册》的格式，**同时输出 DOCX 和 LaTeX 两个版本**。

**v2.2.0 重大更新**：Skill 自身发布自动化 —— agent 被要求「更新这个 skill/脚本」时，自动定位 GitHub 仓库与当前版本，一键完成版本递增与发布。

- 新增 `scripts/release.py`（纯标准库）：自动定位 GitHub 仓库 `wpc6666/cfec-thesis-formatter` 与当前版本（manifest.json 单一权威来源）；`--check` 一键查看仓库/版本/最近 tag/本地状态；`--bump <patch|minor|major>` 按语义化版本递增并同步 **SKILL.md、manifest.json、CHANGELOG** 三处版本号；`--dry-run` 只预览不写入。
- 新增「Skill 自身更新与发布流程【强制】」章节：用户直接要求更新本 skill 时，照该流程执行（定位 → 递增 → 验证 → 同步 → 推送，署名白澤 程）。

**v2.1.0 重大更新**：域自动更新 + 类型子文件夹输出。

- **DOCX 域自动更新**：`scripts/cfec_thesis_docx.py` 保存前自动向 `settings.xml` 注入 `<w:updateFields w:val="true"/>`，Word 打开文档时**自动更新全部域（目录 TOC、页码 PAGE 等）**，无需手动 `Ctrl+A → F9`。
- **目录/页码与 LaTeX 同步**：LaTeX 侧通过「编译两次」自动生成目录与页码；DOCX 侧通过 updateFields 打开自动刷新。两侧目录条目与页码结果一致，无需人工干预。
- **类型子文件夹输出**：版本目录 `v<N>/` 内按类型分子文件夹 `word/`、`latex/`、`data/`，文件按类型归类，命名仍携带版本号（见「输出目录树规范」）。

**v2.0.0 重大更新**：重构输出版本管理，修复「重复运行输出同名文件互相覆盖」的 Bug。

- 新增 `scripts/output_manager.py` —— 输出版本管理器（纯标准库）：按论文题目建项目目录、自动递增版本号 v1/v2/v3、生成 MANIFEST.md 版本清单、输出目录树。
- `scripts/cfec_thesis_docx.py` 新增 `build_and_save()` 推荐入口，一键完成「建目录 → 定版本 → 生成 DOCX → 保存结构化数据 → 写版本清单」。
- **输出永不覆盖**：同一篇论文每次格式化都进入新的 `v<N>/` 版本目录，文件名携带版本号，历史版本全部保留。

**v1.1.1 重大更新**：新增 `scripts/cfec_thesis_docx.py` —— 基于 python-docx 的原生 Word 生成引擎，直接输出二进制 `.docx` 文件，页边距、固定26磅行距、分节页码（封面无码/前置罗马数字/正文阿拉伯数字）、参考文献悬挂缩进、三线表、四级标题样式全部代码级精确控制，无需人工在 Word 中微调。

**格式权威依据**：所有字体字号、页面设置、页码规则、编号样式等详细参数见 [references/format-spec.md](references/format-spec.md)，执行时必须读取该文件。

**LaTeX 模板**：见 [assets/chongqing-caijing-thesis.tex](assets/chongqing-caijing-thesis.tex)，使用说明见 [references/latex-guide.md](references/latex-guide.md)。

**Python DOCX 引擎**：见 [scripts/cfec_thesis_docx.py](scripts/cfec_thesis_docx.py)。

**输出版本管理器**：见 [scripts/output_manager.py](scripts/output_manager.py)。

**Skill 发布辅助（自动定位仓库与版本）**：见 [scripts/release.py](scripts/release.py)。

**更新日志**：见 [CHANGELOG.md](CHANGELOG.md)。

---

## 核心原则

- **只改格式，不改内容**：不修改论文文字、观点、数据、参考文献条目。
- **双版本输出**：默认同时生成 DOCX（python-docx 原生生成）和 LaTeX 两个版本。
- **版本化输出，永不覆盖**：同一论文每次格式化进入新的 `v<N>/` 版本目录，文件名携带迭代版本号，历史输出一律保留。
- **按项目隔离**：输出按论文题目（或用户指定的项目名）建立专属目录，多篇论文互不干扰。
- **以重财手册为准**：不采用 GB/T 7713.2-2022 或其他高校规范。
- **可追溯**：每一步格式调整都有明确依据，输出前逐项自检；项目根目录维护 MANIFEST.md 版本清单。
- **原生生成优先**：DOCX 优先使用 python-docx 脚本原生生成，保证格式零偏差。

---

## Skill 自身更新与发布流程【强制】· v2.2.0

当用户直接要求「更新 / 升级这个 skill（脚本）」时，按以下标准流程执行。**目标：让 agent 自动定位 GitHub 仓库与当前版本，完成一次可追溯的发布，无需用户提供仓库地址或版本号。**

### 1. 自动定位仓库与版本

- **GitHub 仓库**：`https://github.com/wpc6666/cfec-thesis-formatter`（main 分支 + 语义化 tag `vX.Y.Z`）。本地 clone 目录：`/tmp/cfec-thesis-formatter`（不存在则 `git clone`）。
- **本地 skill 目录**：`.user_skills/重财-论文格式规范/`（即本 SKILL.md 所在目录，与 GitHub 仓库 `重财-论文格式规范/` 目录一一对应）。
- **当前版本**：以 `manifest.json` 的 `version` 字段为**单一权威来源**，须与 SKILL.md frontmatter、CHANGELOG 最新条目一致。
- 执行 `python3 scripts/release.py --check` 一次，即可同时拿到仓库、当前版本、最近 tag 与本地状态。

### 2. 版本递增（语义化版本）

- **patch**：修复 Bug（X.Y.Z → X.Y.Z+1）
- **minor**：新增向后兼容功能（X.Y.Z → X.Y+1.0）
- **major**：破坏性变更（X.Y.Z → X+1.0.0）

运行 `python3 scripts/release.py --bump <patch|minor|major>`，自动同步 **SKILL.md frontmatter 与标题、manifest.json、CHANGELOG 新条目**三处版本号；先用 `--dry-run` 预览，确认无误后正式写入，并补充 CHANGELOG 新条目的实际改动内容。

### 3. 本地验证（不允许未验证声称成功）

- 语法检查：`python3 -m py_compile scripts/*.py`
- 功能验证：实际运行受影响脚本（如 DOCX 引擎生成一份样本文档、release.py `--check`/`--bump` 往返），并检查产物关键点。
- 文档一致性：SKILL.md / manifest.json / CHANGELOG 版本号三处一致。

### 4. 同步本地与仓库 clone

```bash
LOCAL=".user_skills/重财-论文格式规范"
REPO="/tmp/cfec-thesis-formatter/重财-论文格式规范"
rm -rf "$REPO" && mkdir -p "$REPO"
cp -R "$LOCAL/." "$REPO/"
find "$REPO" -name "__pycache__" -type d -exec rm -rf {} +
diff -r "$LOCAL" "$REPO" --exclude="__pycache__"   # 必须输出「完全一致」
```

### 5. 提交、打 tag、推送（署名固定）

- 提交身份固定：`白澤 程 <87405429+wpc6666@users.noreply.github.com>`（仓库本地 config 已固定，勿用全局身份）。
- 提交信息：首行 `vX.Y.Z: <本次改动一句话>`，正文列出改动点。
- 打 tag `vX.Y.Z`，推送 `main` + tag；推送后 `git ls-remote --tags origin | grep vX.Y.Z` 确认。

### 6. 汇报

交付说明必须包含：本次版本号、改动内容、验证方式与结果、仓库提交与 tag 信息。

---

## 输出目录树规范【强制】

所有产物必须进入版本化目录，**版本目录内按类型分子文件夹**，结构固定如下：

```
<输出根目录>/
└── <项目名slug>/                    # 项目目录：取自论文题目（≤40字符），用户可显式指定
    ├── v1/                          # 第 1 次输出（版本号自动递增，绝不覆盖）
    │   ├── word/                    # Word 成品
    │   │   └── <题目slug>_v1.docx
    │   ├── latex/                   # LaTeX 源码 + 编译产物
    │   │   ├── <题目slug>_v1.tex
    │   │   └── <题目slug>_v1.pdf    # 可选：XeLaTeX 编译生成的 PDF
    │   └── data/                    # 结构化论文数据（可复用重新生成）
    │       └── <题目slug>_v1_data.json
    ├── v2/                          # 第 2 次输出
    │   └── ...
    └── MANIFEST.md                  # 版本清单：各版本生成时间 / 文件列表 / 说明
```

### 命名规则

| 对象 | 规则 | 示例（题目：基于深度学习的X研究） |
|------|------|------|
| 项目目录 | `<题目slug>/`，≤40字符，非法字符替换为 `_` | `基于深度学习的X研究/` |
| 版本目录 | `v<N>`，N 自动递增（可指定） | `v1/`、`v2/` |
| 类型子文件夹 | `word/`、`latex/`、`data/`（固定三名，勿改名） | `v1/word/`、`v1/latex/`、`v1/data/` |
| DOCX | `word/<题目slug>_v<N>.docx` | `word/基于深度学习的X研究_v1.docx` |
| LaTeX | `latex/<题目slug>_v<N>.tex` | `latex/基于深度学习的X研究_v1.tex` |
| PDF（可选） | `latex/<题目slug>_v<N>.pdf` | `latex/基于深度学习的X研究_v1.pdf` |
| 结构化数据 | `data/<题目slug>_v<N>_data.json` | `data/基于深度学习的X研究_v1_data.json` |
| 版本清单 | 项目根目录 `MANIFEST.md` | `基于深度学习的X研究/MANIFEST.md` |

slug 化规则：保留中文、字母、数字、连字符、下划线；Windows/macOS 非法字符（`\ / : * ? " < > |`）及空白折叠为 `_`；超长截断。

### 版本号规则

1. **默认自动递增**：扫描项目目录下已有 `v1…vN`，取最大 +1。第 1 次运行 `v1`，第 2 次 `v2`，以此类推。
2. **绝不覆盖**：目标版本目录已存在且非空时，禁止写入，改用下一版本号（或提示用户）。
3. **可指定版本**：用户明确要求继续某一版本（如"在 v2 基础上修订"）时，传入 `force_version=2`；若 v2 已存在非空，报错并要求使用新版本号，不得静默覆盖。
4. **输出根目录**：默认用户工作目录；LaTeX 与 DOCX 产物必须放在**同一个** `v<N>/` 版本目录内（分别位于 `latex/` 与 `word/` 子文件夹）。

### 域自动更新规则【v2.1.0】

- **DOCX**：引擎保存前自动注入 `<w:updateFields w:val="true"/>`（见 `scripts/cfec_thesis_docx.py::_enable_auto_update_fields`）。Word 打开文档时**自动更新全部域**（目录 TOC、页码 PAGE 等），目录页码与正文自动一致，**禁止要求用户手动 `Ctrl+A → F9`**。
- **LaTeX**：通过「XeLaTeX 编译两次」自动生成目录与页码；第二次编译后目录页码与正文一致。
- **同步校验**：DOCX 打开后目录页码应与 LaTeX 编译 PDF 的目录页码一致（同一份 thesis_data，章节顺序相同）。

---

## 工作流

### Step 0：前置检查

在开始处理前，必须完成以下检查：

1. **输入完整性检查**
   - 确认用户提供了论文内容（文件或文本）。
   - 确认论文包含：封面信息、中文摘要、英文摘要、正文、参考文献。
   - 缺少部分向用户说明，用占位符标注，不编造内容。

2. **环境检查**
   - DOCX 生成：确认可用 `python-docx >= 0.8.11`。若不可用，执行 `pip install python-docx`。
   - LaTeX 生成：确认可读取模板文件，输出 `.tex` 源码。
   - 若环境不支持某一版本，明确告知用户，只输出支持的版本。

3. **格式规范加载**
   - 读取 [references/format-spec.md](references/format-spec.md)，加载全部格式参数。
   - 读取 [assets/chongqing-caijing-thesis.tex](assets/chongqing-caijing-thesis.tex)，作为 LaTeX 输出基础模板。
   - 读取 [scripts/cfec_thesis_docx.py](scripts/cfec_thesis_docx.py) 与 [scripts/output_manager.py](scripts/output_manager.py)，确认引擎可用。

---

### Step 1：确认输入与输出规划

1. 获取用户提供的论文文件（`.docx` 或文本内容）。
2. **确定项目名与题目**：
   - 优先取论文封面中的题目（`cover.title`）。
   - 用户显式指定项目名时，以用户指定为准。
   - 目录名/文件名由题目 slug 化生成（见上节命名规则）。
3. **规划输出版本**：调用输出管理器确定本次版本号——
   ```bash
   # 预览：本次输出会落在哪个目录、文件名是什么（不创建目录）
   python3 scripts/output_manager.py --root <输出根目录> --title "<论文题目>" --plan
   # 指定版本（可选）：--version 2
   ```
4. 若用户要求"在上一版基础上改"，先列出项目目录树确认已有版本：
   ```bash
   python3 scripts/output_manager.py --root <输出根目录> --project "<项目名>" --list
   ```
   然后按规则 3 处理：历史版本保留，本次输出进入下一版本目录（除非用户明确指定且目录为空）。

---

## DOCX 版本工作流（python-docx 原生生成）

### Step 2：将论文内容解析为结构化数据

将用户论文内容解析为 `thesis_data` 字典结构，供 Python 引擎消费：

```python
thesis_data = {
    "cover": {
        "title": "论文题目",
        "college": "学院",
        "grade": "年级",
        "major": "专业",
        "student_id": "学号",
        "name": "姓名",
        "advisor": "指导教师",
        "date": "二〇二六年六月"
    },
    "abstract_cn": {
        "content": "中文摘要正文...",
        "keywords": ["关键词1", "关键词2", "关键词3"]
    },
    "abstract_en": {
        "content": "English abstract...",
        "keywords": ["keyword1", "keyword2"]
    },
    "chapters": [
        {
            "title": "绪论",
            "sections": [
                {
                    "title": "研究背景",
                    "paragraphs": ["正文段落1[1]", "正文段落2"],
                    "subsections": [
                        {
                            "title": "三级标题",
                            "paragraphs": ["正文段落"]
                        }
                    ]
                }
            ]
        }
    ],
    "references": [
        "作者. 标题[J]. 期刊, 年份, 卷(期): 页码.",
        "作者. 书名[M]. 出版地: 出版社, 年份."
    ],
    "appendices": [
        {"label": "A", "title": "附录标题", "content": "附录内容"}
    ]
}
```

**解析规则**：
- 一级标题 → `chapters[].title`
- 二级标题 → `chapters[].sections[].title`
- 三级标题 → `sections[].subsections[].title`
- 四级标题 → 在 paragraphs 中以 `（1）` 开头标注，Python 引擎自动识别
- 正文引用 `[1]` → 保留在文本中，Python 引擎自动转为上标
- 图片 → 标记为 `[图:图片路径]`，Python 引擎自动插入并编号
- 表格 → 标记为 `[表:表头|数据行]`，Python 引擎自动生成三线表
- 公式 → 标记为 `[公式:E=mc^2]`，Python 引擎自动编号右对齐

### Step 3：调用 Python 引擎生成 DOCX（版本化输出）

```bash
# 安装依赖（首次）
pip install python-docx

# 运行生成脚本（默认输出到 ./output/<题目slug>/v<N>/）
python3 scripts/cfec_thesis_docx.py
```

推荐代码调用（自动版本化，绝不覆盖历史）：

```python
from scripts.cfec_thesis_docx import ThesisFormatter

formatter = ThesisFormatter()
docx_path, data_path, mgr = formatter.build_and_save(
    thesis_data,
    root_dir="output",                # 输出根目录
    project="基于Python的论文",       # 项目名（默认取 cover.title）
    # force_version=2,                # 可选：指定版本
)
print(mgr.tree())                     # 打印输出目录树
```

**`build_and_save()` 自动完成**：建项目目录 → 计算下一版本号 v<N> → 生成 DOCX 到 `v<N>/` → 保存 `thesis_data` JSON → 刷新 MANIFEST.md。

**Python 引擎自动完成的格式设置**（全部代码级精确控制，无需人工调整）：

| 格式项 | 实现方式 |
|--------|---------|
| 页面 A4 + 左2.8cm/其余2.5cm | `section.page_width/height` + `geometry` |
| 全文固定26磅行距 | `WD_LINE_SPACING.EXACTLY` + `Pt(26)` |
| 封面无页码 | 第一节页脚清空 + 页码起始0 |
| 前置部分大写罗马数字页码 | `pgNumType fmt="upperRoman"` + 页脚 PAGE 域 |
| 正文阿拉伯数字从1开始 | 新节 `pgNumType fmt="decimal" start=1` |
| 四级标题编号（一、→（一）→1.→（1）） | 自定义计数器 + 中文数字转换 |
| 一级标题黑体小三加粗居中 | `set_run_font(heiti, 15pt, bold)` + 居中 |
| 图 1-1 编号（图题在下） | 按章节计数器 + `\caption` 位置 |
| 表 1-1 编号（表题在上）+ 三线表 | 按章节计数器 + `set_cell_border` 仅顶/中/底三线 |
| 公式（1.1）右对齐 | 按章节计数器 + 右对齐编号 |
| 参考文献悬挂缩进2字符 | `left_indent=2char` + `first_line_indent=-2char` |
| 正文引用 [1] 自动上标 | 正则分割 + `run.font.superscript=True` |
| 目录域（显示到二级） | `TOC \o "1-2"` 域代码；已注入 updateFields，打开文档自动更新 |
| 中英文字体分离 | `w:rFonts eastAsia=宋体 ascii=Times New Roman` |

### Step 4：DOCX 后处理与验证

1. **打开文档验证**：确认文件可正常打开，无损坏。
2. **域自动更新**：引擎已注入 `updateFields`，Word 打开文档时自动更新目录/页码域（无需手动 `Ctrl+A → F9`；若个别 Word 版本未弹窗刷新，可按 `F9` 手动补一次，但不作为交付要求）。
3. **对照自检清单**：逐项核对 [references/format-spec.md](references/format-spec.md) 末尾的"常见易错点清单"。
4. **抽查页面**：封面、中文摘要页、Abstract 页、目录页、正文第一页、含图表页、参考文献页。
5. **导出 PDF 验证**：确认分页正确，目录页码与正文一致。

---

## LaTeX 版本工作流

### Step 5：生成 LaTeX 源码（版本化输出）

基于 [assets/chongqing-caijing-thesis.tex](assets/chongqing-caijing-thesis.tex) 模板，将论文内容填入对应位置。**LaTeX 产物必须与 DOCX 放入同一个 `v<N>/` 版本目录**（DOCX 在 `word/`，LaTeX 在 `latex/`）：

1. **复制模板**：将模板文件复制为 `<题目slug>_v<N>.tex`（版本号与 DOCX 一致）。
2. **封面**：替换 titlepage 环境内的题目、学院、年级、专业、学号、姓名、指导教师、日期。
3. **中文摘要**：替换摘要正文和关键词（关键词间用全角分号"；"）。
4. **英文摘要**：替换 Abstract 正文和 Keywords（关键词间用半角分号";"）。
5. **正文**：使用自定义标题命令撰写，**禁止使用标准 \section/\subsection**：
   - `\hA{第一章标题}` → 一级，"一、"编号，黑体小三号加粗居中
   - `\hB{二级标题}` → 二级，"（一）"编号，黑体四号首行缩进
   - `\hC{三级标题}` → 三级，"1."编号，黑体小四号首行缩进
   - `\hD{四级标题}` → 四级，"（1）"编号，宋体小四号加粗首行缩进
6. **图表**：
   - 图片用 `figure` 环境，`\caption{}` 置于 `\includegraphics` 之后（图题在下方），自动编号"图 1-1"
   - 表格用 `table` 环境，`\caption{}` 置于 `tabular` 之前（表题在上方），推荐 `booktabs` 三线表，自动编号"表 1-1"
7. **公式**：用 `equation` 环境，自动编号（1.1）右对齐。
8. **参考文献**：在 `therefs` 环境中用 `\item` 逐条添加，自动悬挂缩进 2 字符、五号字。正文引用用 `\cite{key}` 自动上标。
9. **附录**（如有）：复制附录 A 块，改编号为"附录 B"等。
10. **页码与目录**：模板已自动处理——封面无页码、前置部分大写罗马数字、正文阿拉伯数字从1开始；目录只显示到二级，点号前导符。**编译两次后目录与页码自动生成，与 DOCX 打开自动更新后的目录/页码同步一致**。

### Step 6：LaTeX 格式校验

对照 [references/latex-guide.md](references/latex-guide.md) 检查：

1. 编译器设置为 **XeLaTeX**（非 PDFLaTeX）。
2. 所有标题使用 `\hA`–`\hD` 命令，未使用标准 `\section` 等。
3. 图题在 `\includegraphics` 之后，表题在 `tabular` 之前。
4. 中文关键词用全角分号"；"，英文用半角分号";"。
5. 参考文献在 `therefs` 环境内，每条用 `\item`。
6. 正文引用使用 `\cite{}`，未手动打上标。
7. 编译两次以生成目录和交叉引用；编译后目录页码与正文一致（与 DOCX 自动更新域的结果同步）。

---

## 验证与交付

### Step 7：双版本验证

**DOCX 版本（python-docx 原生生成）**：

1. 文件可正常打开，无损坏、无乱码、无字体缺失。
2. Word 打开文档时目录/页码域自动更新（已注入 updateFields），目录页码与正文一致；若所用 Word 版本未自动刷新，手动 `F9` 一次后一致亦可。
3. 对照 [references/format-spec.md](references/format-spec.md) 末尾的"常见易错点清单"逐项核对。
4. 抽查页面：封面、中文摘要页、Abstract 页、目录页、正文第一页、含图表页、参考文献页。
5. 可正常导出 PDF 且分页正确。
6. **v2.1.0 重点验证项（域自动更新 + 子文件夹输出）**：
   - [ ] DOCX 的 `word/settings.xml` 含 `<w:updateFields w:val="true"/>`
   - [ ] 输出位于 `<项目slug>/v<N>/word/`、`latex/`、`data/` 三个子文件夹，文件名携带版本号
   - [ ] 打开 Word 后目录/页码域自动刷新，目录页码与正文一致
   - [ ] LaTeX 编译两次后目录/页码与 DOCX 结果同步一致
7. **v2.0.0 重点验证项（输出版本管理）**：
   - [ ] 输出进入 `<题目slug>/v<N>/` 版本目录，文件名携带版本号
   - [ ] 同一论文第二次格式化生成 `v<N+1>/`，旧版本文件未被修改
   - [ ] 项目根目录 MANIFEST.md 记录了全部版本的生成时间与文件清单
   - [ ] 未出现同名覆盖（v1 与 v2 文件名不同）
8. **v1.1.1 重点验证项**：
   - [ ] 页边距：左 2.8cm，其余 2.5cm（Word→布局→页边距→自定义边距查看）
   - [ ] 全文固定行距 26 磅（Word→开始→段落→行距→固定值 26磅）
   - [ ] 封面无页码
   - [ ] 摘要/目录页码为大写罗马数字 Ⅰ、Ⅱ、Ⅲ
   - [ ] 正文页码从 1 开始阿拉伯数字
   - [ ] 参考文献悬挂缩进 2 字符
   - [ ] 正文引用 [1] 全部上标
   - [ ] 表格为三线表（仅顶线、表头线、底线，无竖线）
   - [ ] 图题在图下方，表题在表上方
   - [ ] 一级标题"一、"居中黑体小三加粗

**LaTeX 版本**：

1. `.tex` 文件语法正确，XeLaTeX 编译无致命错误。
2. 编译两次后目录、图表编号、公式编号、引用均正确。
3. 生成的 PDF 页边距、行距、字体字号符合规范。
4. 页码分节正确（封面无码、前置罗马数字、正文阿拉伯数字）。
5. 所有标题使用 `\hA`–`\hD`，无标准 `\section`。
6. 文件名与 DOCX 同版本号，位于同一 `v<N>/` 版本目录的 `latex/` 子文件夹。

### Step 8：交付

交付以下文件（全部位于同一版本目录 `v<N>/` 内，按类型分入子文件夹）：

- `word/<题目slug>_v<N>.docx` — Word 版本（python-docx 原生生成，打开时自动更新目录/页码域）
- `latex/<题目slug>_v<N>.tex` — LaTeX 源码版本（XeLaTeX 编译两次）
- `latex/<题目slug>_v<N>.pdf` — 可选，LaTeX 编译生成的 PDF
- `data/<题目slug>_v<N>_data.json` — 结构化论文数据，可复用重新生成
- `<项目目录>/MANIFEST.md` — 版本清单

**交付说明必须包含**：
1. **输出目录树**：用 `mgr.tree()` 或输出管理器 `--list` 打印完整目录树（含 word/latex/data 子文件夹），让用户看到每个版本的存放位置。
2. 本次版本号与历史版本列表（引用 MANIFEST.md）。
3. DOCX 生成方式：python-docx 引擎 v2.1.0，已注入 updateFields，**打开时自动更新目录/页码域，无需手动 F9**。
4. LaTeX 编译方式：XeLaTeX，编译两次（目录与页码自动生成，与 DOCX 同步）。
5. 已完成的格式调整项（对照自检清单）。
6. 自检结果（通过/未通过项）。
7. 用户需要手动确认的事项（如图片替换、参考文献核对）。

---

## 错误处理与回退

### 常见错误及处理

| 错误场景 | 处理方式 |
|---------|---------|
| python-docx 未安装 | 执行 `pip install python-docx`，安装后重试 |
| 目标版本目录已存在且非空 | 不覆盖；改用下一版本号，或提示用户清理/换版本号 |
| DOCX 生成失败 | 回退为输出格式说明文档，告知用户手动调整步骤；或使用 pandoc 中转 |
| LaTeX 模板读取失败 | 输出最小可用模板，标注缺失部分 |
| 论文内容不完整 | 用占位符标注缺失部分，不编造内容 |
| 字体不可用 | 说明替代字体，优先保证格式结构正确 |
| 目录生成异常 | 手动列出目录结构；DOCX 已注入 updateFields 打开自动刷新，标注为临时方案 |
| 图片路径不存在 | 插入图片占位符，标注需用户手动替换 |

### 降级原则

- 双版本输出失败时，至少保证一个版本完整输出。
- python-docx 原生生成失败时，回退为 Markdown + pandoc 中转，并明确告知格式偏差范围。
- 复杂格式（如自动目录）失败时，输出手动操作指南。
- 永远不因为格式处理失败而修改论文内容。

---

## 允许与禁止

**允许**：删除临时格式说明；调整标题层级、目录、页码、字体字号、图表题注、参考文献列表外观；统一正文引用编号为上标；将论文内容填入 LaTeX 模板对应位置；将论文内容解析为结构化数据调用 python-docx 引擎；输出格式自检报告；输出 thesis_data.json 供复用；按输出目录树规范创建项目目录与版本目录；生成 MANIFEST.md 版本清单。

**禁止**：改写观点、增删论据、编造文献、补充身份信息、修改数据结果、覆盖原文件、覆盖已存在的非空版本目录、输出到无版本号的固定文件名、在 LaTeX 中使用标准 \section 替代自定义 \hA–\hD 命令、在未告知用户的情况下省略输出版本、在 python-docx 生成失败时静默降级而不告知用户格式偏差。
