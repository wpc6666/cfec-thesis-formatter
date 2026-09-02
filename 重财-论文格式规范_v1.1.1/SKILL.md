---
name: 重财-论文格式规范
version: 1.1.1
description: 将中文学术论文（重庆财经学院本科毕业论文/设计）调整为符合该校《毕业论文（设计）指导手册》的格式规范，同时输出 DOCX（Word）和 LaTeX 两个版本。v1.1.1 新增 python-docx 原生 Word 生成引擎，可直接生成二进制 .docx 文件，页边距、行距、分节页码、悬挂缩进、三线表全部代码级精确控制。涵盖封面、摘要（中英文）、目录、正文标题、正文文本、图表、公式、参考文献、附录、页码、装订顺序及字体字号。本技能完全依据重庆财经学院指导手册，不采用 GB/T 7713.2-2022 标准。只处理格式，不改写论文内容。LaTeX 模板基于 ctexart + XeLaTeX，内置四级标题命令（一、→（一）→1.→（1））、图1-1/表1-1编号、公式（1.1）编号、罗马/阿拉伯分节页码、悬挂缩进参考文献。当用户要求以下任务时自动调用：按重庆财经学院毕业论文手册调整论文格式；同时生成Word和LaTeX版本；检查摘要、Abstract、目录、正文标题层级是否符合手册要求；检查字体字号（小三、四号、小四、五号等）是否正确；检查页码规则（前置部分罗马数字，正文阿拉伯数字）；检查行距是否为固定值26磅；检查页边距（左2.8cm，上/下/右均为2.5cm）；检查图序表序是否用"图 1-1""表 1-1"格式；检查参考文献引用是否为上标；生成或更新目录（仅显示到二级标题）；生成重财本科毕设LaTeX模板；使用python-docx生成原生Word文档。
---

# 重财-论文格式规范 v1.1.1

将重庆财经学院本科毕业论文调整为符合该校《毕业论文（设计）指导手册》的格式，**同时输出 DOCX 和 LaTeX 两个版本**。

**v1.1.1 重大更新**：新增 `scripts/cfec_thesis_docx.py` —— 基于 python-docx 的原生 Word 生成引擎，直接输出二进制 `.docx` 文件，页边距、固定26磅行距、分节页码（封面无码/前置罗马数字/正文阿拉伯数字）、参考文献悬挂缩进、三线表、四级标题样式全部代码级精确控制，无需人工在 Word 中微调。

**格式权威依据**：所有字体字号、页面设置、页码规则、编号样式等详细参数见 [references/format-spec.md](references/format-spec.md)，执行时必须读取该文件。

**LaTeX 模板**：见 [assets/chongqing-caijing-thesis.tex](assets/chongqing-caijing-thesis.tex)，使用说明见 [references/latex-guide.md](references/latex-guide.md)。

**Python DOCX 引擎**：见 [scripts/cfec_thesis_docx.py](scripts/cfec_thesis_docx.py)。

**更新日志**：见 [CHANGELOG.md](CHANGELOG.md)。

---

## 核心原则

- **只改格式，不改内容**：不修改论文文字、观点、数据、参考文献条目。
- **双版本输出**：默认同时生成 DOCX（python-docx 原生生成）和 LaTeX 两个版本。
- **不覆盖原文件**：输出新文件，命名为 `原文件名_重财规范版.docx` 和 `原文件名_重财规范版.tex`。
- **以重财手册为准**：不采用 GB/T 7713.2-2022 或其他高校规范。
- **可追溯**：每一步格式调整都有明确依据，输出前逐项自检。
- **原生生成优先**：DOCX 优先使用 python-docx 脚本原生生成，保证格式零偏差。

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
   - 读取 [scripts/cfec_thesis_docx.py](scripts/cfec_thesis_docx.py)，确认 Python 引擎可用。

---

### Step 1：确认输入与输出

1. 获取用户提供的论文文件（`.docx` 或文本内容）。
2. 确认输出：同时生成两个版本——
   - `原文件名_重财规范版.docx`（python-docx 原生生成）
   - `原文件名_重财规范版.tex`（LaTeX 源码）
3. 若用户只需要其中一个版本，按用户要求输出。

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

### Step 3：调用 Python 引擎生成 DOCX

```bash
# 安装依赖（首次）
pip install python-docx

# 运行生成脚本
python3 scripts/cfec_thesis_docx.py
```

或在代码中调用：

```python
from scripts.cfec_thesis_docx import ThesisFormatter

formatter = ThesisFormatter()
formatter.build(thesis_data)
formatter.save("原文件名_重财规范版.docx")
```

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
| 目录域（显示到二级） | `TOC \o "1-2"` 域代码，Word中F9更新 |
| 中英文字体分离 | `w:rFonts eastAsia=宋体 ascii=Times New Roman` |

### Step 4：DOCX 后处理与验证

1. **打开文档验证**：确认文件可正常打开，无损坏。
2. **更新域**：在 Word 中按 `Ctrl+A` 全选，再按 `F9` 更新所有域（目录、页码）。
3. **对照自检清单**：逐项核对 [references/format-spec.md](references/format-spec.md) 末尾的"常见易错点清单"。
4. **抽查页面**：封面、中文摘要页、Abstract 页、目录页、正文第一页、含图表页、参考文献页。
5. **导出 PDF 验证**：确认分页正确。

---

## LaTeX 版本工作流

### Step 5：生成 LaTeX 源码

基于 [assets/chongqing-caijing-thesis.tex](assets/chongqing-caijing-thesis.tex) 模板，将论文内容填入对应位置：

1. **复制模板**：将模板文件复制为 `原文件名_重财规范版.tex`。
2. **封面**：替换 titlepage 环境内的题目、学院、年级、专业、学号、姓名、指导教师、日期。
3. **中文摘要**：替换摘要正文和关键词（关键词间用全角分号"；"）。
4. **英文摘要**：替换 Abstract 正文和 Keywords（关键词间用半角分号";"）。
5. **正文**：使用自定义标题命令撰写，**禁止使用标准 \section/\subsection**：
   - `\h1{第一章标题}` → 一级，"一、"编号，黑体小三号加粗居中
   - `\h2{二级标题}` → 二级，"（一）"编号，黑体四号首行缩进
   - `\h3{三级标题}` → 三级，"1."编号，黑体小四号首行缩进
   - `\h4{四级标题}` → 四级，"（1）"编号，宋体小四号加粗首行缩进
6. **图表**：
   - 图片用 `figure` 环境，`\caption{}` 置于 `\includegraphics` 之后（图题在下方），自动编号"图 1-1"
   - 表格用 `table` 环境，`\caption{}` 置于 `tabular` 之前（表题在上方），推荐 `booktabs` 三线表，自动编号"表 1-1"
7. **公式**：用 `equation` 环境，自动编号（1.1）右对齐。
8. **参考文献**：在 `therefs` 环境中用 `\item` 逐条添加，自动悬挂缩进 2 字符、五号字。正文引用用 `\cite{key}` 自动上标。
9. **附录**（如有）：复制附录 A 块，改编号为"附录 B"等。
10. **页码与目录**：模板已自动处理——封面无页码、前置部分大写罗马数字、正文阿拉伯数字从1开始；目录只显示到二级，点号前导符。

### Step 6：LaTeX 格式校验

对照 [references/latex-guide.md](references/latex-guide.md) 检查：

1. 编译器设置为 **XeLaTeX**（非 PDFLaTeX）。
2. 所有标题使用 `\h1`–`\h4` 命令，未使用标准 `\section` 等。
3. 图题在 `\includegraphics` 之后，表题在 `tabular` 之前。
4. 中文关键词用全角分号"；"，英文用半角分号";"。
5. 参考文献在 `therefs` 环境内，每条用 `\item`。
6. 正文引用使用 `\cite{}`，未手动打上标。
7. 编译两次以生成目录和交叉引用。

---

## 验证与交付

### Step 7：双版本验证

**DOCX 版本（python-docx 原生生成）**：

1. 文件可正常打开，无损坏、无乱码、无字体缺失。
2. 在 Word 中按 `Ctrl+A` → `F9` 更新所有域后，目录页码与正文一致。
3. 对照 [references/format-spec.md](references/format-spec.md) 末尾的"常见易错点清单"逐项核对。
4. 抽查页面：封面、中文摘要页、Abstract 页、目录页、正文第一页、含图表页、参考文献页。
5. 可正常导出 PDF 且分页正确。
6. **v1.1.1 重点验证项**：
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
5. 所有标题使用 `\h1`–`\h4`，无标准 `\section`。

### Step 8：交付

交付以下文件：

- `原文件名_重财规范版.docx` — Word 版本（python-docx 原生生成，打开后按 F9 更新域）
- `原文件名_重财规范版.tex` — LaTeX 源码版本（XeLaTeX 编译两次）
- （可选）`原文件名_重财规范版.pdf` — LaTeX 编译生成的 PDF
- （可选）`thesis_data.json` — 结构化论文数据，可复用重新生成

**交付说明必须包含**：
1. 输出文件列表及路径
2. DOCX 生成方式：python-docx 引擎 v1.1.1，需在 Word 中按 F9 更新目录域
3. LaTeX 编译方式：XeLaTeX，编译两次
4. 已完成的格式调整项（对照自检清单）
5. 自检结果（通过/未通过项）
6. 用户需要手动确认的事项（如图片替换、参考文献核对）

---

## 错误处理与回退

### 常见错误及处理

| 错误场景 | 处理方式 |
|---------|---------|
| python-docx 未安装 | 执行 `pip install python-docx`，安装后重试 |
| DOCX 生成失败 | 回退为输出格式说明文档，告知用户手动调整步骤；或使用 pandoc 中转 |
| LaTeX 模板读取失败 | 输出最小可用模板，标注缺失部分 |
| 论文内容不完整 | 用占位符标注缺失部分，不编造内容 |
| 字体不可用 | 说明替代字体，优先保证格式结构正确 |
| 目录生成异常 | 手动列出目录结构，标注需用户在 Word 中 F9 更新域 |
| 图片路径不存在 | 插入图片占位符，标注需用户手动替换 |

### 降级原则

- 双版本输出失败时，至少保证一个版本完整输出。
- python-docx 原生生成失败时，回退为 Markdown + pandoc 中转，并明确告知格式偏差范围。
- 复杂格式（如自动目录）失败时，输出手动操作指南。
- 永远不因为格式处理失败而修改论文内容。

---

## 允许与禁止

**允许**：删除临时格式说明；调整标题层级、目录、页码、字体字号、图表题注、参考文献列表外观；统一正文引用编号为上标；将论文内容填入 LaTeX 模板对应位置；将论文内容解析为结构化数据调用 python-docx 引擎；输出格式自检报告；输出 thesis_data.json 供复用。

**禁止**：改写观点、增删论据、编造文献、补充身份信息、修改数据结果、覆盖原文件、在 LaTeX 中使用标准 \section 替代自定义 \h1–\h4 命令、在未告知用户的情况下省略输出版本、在 python-docx 生成失败时静默降级而不告知用户格式偏差。
