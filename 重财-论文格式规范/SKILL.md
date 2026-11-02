---
name: 重财-论文格式规范
version: 2.2.9
description: 将重庆财经学院本科毕业论文调整为学校指导手册格式，同时生成 DOCX 与 LaTeX；检查封面、摘要、目录、标题层级、字体字号、页码、行距、页边距、图表编号、参考文献上标和版本化输出。用于“重财论文格式、毕业论文排版、Word/LaTeX 双版本、格式自检、目录页码更新”等请求。只改格式，不改写内容；以重财手册为准，不采用 GB/T 7713.2-2022。
---

# 重财-论文格式规范 v2.2.9
将重庆财经学院本科毕业论文调整为符合该校《毕业论文（设计）指导手册》的格式，**同时输出 DOCX 和 LaTeX 两个版本**。

**版本变更**：当前为 v2.2.9；历史版本与迁移说明见 `CHANGELOG.md`。

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

## 按需手册

| 场景 | 指针 |
|---|---|
| 发布 / 版本递增 | `references/RELEASE.md` |
| 输出目录与版本规则 | `references/OUTPUT.md` |
| DOCX 生成 | `references/DOCX-WORKFLOW.md` |
| LaTeX 生成 | `references/LATEX-WORKFLOW.md` |
| 验证与交付 | `references/VALIDATION.md` |
| 错误处理与回退 | `references/TROUBLESHOOTING.md` |
| 格式权威参数 | `references/format-spec.md` |
| LaTeX 使用说明 | `references/latex-guide.md` |

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
   - 目录名/文件名由题目 slug 化生成，规则见
     [references/OUTPUT.md](references/OUTPUT.md#命名规则)。
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

## 允许与禁止

**允许**：删除临时格式说明；调整标题层级、目录、页码、字体字号、图表题注、参考文献列表外观；统一正文引用编号为上标；将论文内容填入 LaTeX 模板对应位置；将论文内容解析为结构化数据调用 python-docx 引擎；输出格式自检报告；输出 thesis_data.json 供复用；按输出目录树规范创建项目目录与版本目录；生成 MANIFEST.md 版本清单。

**禁止**：改写观点、增删论据、编造文献、补充身份信息、修改数据结果、覆盖原文件、覆盖已存在的非空版本目录、输出到无版本号的固定文件名、在 LaTeX 中使用标准 \section 替代自定义 \hA–\hD 命令、在未告知用户的情况下省略输出版本、在 python-docx 生成失败时静默降级而不告知用户格式偏差。
