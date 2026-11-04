---
name: 重财-论文格式规范
version: 3.5.1
description: 按重庆财经学院《毕业论文（设计）指导手册（2027届）》生成或调整本科毕业论文，输出 DOCX（正文可另出 LaTeX）；一键装配封面、中英文摘要、目录、正文、参考文献、附录，并生成 13 项附件材料（每表一个子文件夹）。用于“重财论文、2027届毕业论文、封面、附件表、开题报告、任务书、Word/LaTeX 双版本、格式自检、目录页码更新”等请求。只改格式，不改写内容；以重财 2027 手册为准，不采用 GB/T 7713.2-2022。
---

# 重财-论文格式规范 v3.5.1

依据《重庆财经学院毕业论文（设计）指导手册（2027届）》，把论文排成学校格式，并可一键生成从封面到附录的完整项目文件夹。

**版本变更**：当前为 v3.5.1（2027 届）；历史版本见 `CHANGELOG.md`。

## 触发条件（重要）

- **只有用户明确调用 `/cfec-thesis-formatter`**，或明确说“用重财论文格式规范 / 格式化这份论文”时，才执行格式修改。
- 普通文档生成、整理、转换任务中，不自动套用 2027 论文模板，不自动改字体、页码、三线表。
- 用户没给模板时，先问清楚：用课程/学校模版、用 2027 论文模版，还是不用模板。不要默认选论文模板。
- 课程报告用 `make-report-project.sh`，论文用 `make-thesis-2027.sh`，两者不要混用。

## Word 自动导出与授权（macOS）

- 生成器会自动把 `03-输出` 的 docx 导出成 `05-测试/*预览.pdf`：优先 Microsoft Word，其次 soffice。
- 首次用 Word 时，macOS 会弹“允许控制 Microsoft Word”，点“允许”即可；也可先运行
  `bash <skill目录>/authorize-word.sh` 主动触发。
- 如果之前拒绝过：系统设置 → 隐私与安全性 → 自动化 → 勾选运行脚本的 App 下的 Microsoft Word。
- 授权一次后永久生效；未授权时生成器会跳过 PDF 并在 `提醒.md` 里写原因，不会卡住生成。
- Windows / Linux 没有 AppleEvents 权限，直接走 soffice。

**自带资源**：封面、11 张附件表、样式样例、专业班级对照表和原手册都在
`assets/2027/`；换电脑只装本 Skill 即可，不需要另外的论文工具目录。

**格式权威依据**：`references/format-spec.md`（2027 参数）。

**一键生成器**：`<skill目录>/make-thesis-2027.sh`，说明见 `references/GENERATOR-2027.md`。

**附件表**：`references/FORMS-2027.md`；模板在 `<skill目录>/assets/2027/模板/`。

**LaTeX 模板**：`assets/chongqing-caijing-thesis.tex`（只用于正文，不含封面和附件表）。

**Python DOCX 引擎（format-only）**：`scripts/cfec_thesis_docx.py`。

**输出版本管理器**：`scripts/output_manager.py`。

**Skill 发布辅助**：`scripts/release.py`。

---

## 核心原则

- **只改格式，不改内容**：不修改论文文字、观点、数据、参考文献条目。
- **2027 手册优先**：页面、字体、页码、封面、附件表一律以 2027 手册为准。
- **封面与附件表用学校原件**：从手册原样提取，不重画、不仿制；查重/AIGC 报告不伪造模板。
- **双版本输出**：完整论文默认 DOCX；正文可选 LaTeX / PDF。
- **版本化输出，永不覆盖**：同一论文每次进入新的 `v<N>/` 目录；生成项目时目标目录非空则拒绝覆盖。
- **原文件只读**：学校手册与你提供的原稿只复制，不修改。
- **可追溯**：每个项目保留 `05-测试/格式自检.md`，记录生成项与未验证项。

---

## 按需手册

| 场景 | 指针 |
|---|---|
| 2027 格式参数 | `references/format-spec.md` |
| 一键生成项目、目录结构、输入字段 | `references/GENERATOR-2027.md` |
| 13 项附件材料与自动填充 | `references/FORMS-2027.md` |
| 课程报告 → 多级项目文件夹 | `references/REPORT-PROJECT.md` |
| 发布 / 版本递增 | `references/RELEASE.md` |
| 输出目录与版本规则 | `references/OUTPUT.md` |
| DOCX 生成（format-only） | `references/DOCX-WORKFLOW.md` |
| LaTeX 生成 | `references/LATEX-WORKFLOW.md` |
| 验证与交付 | `references/VALIDATION.md` |
| 错误处理与回退 | `references/TROUBLESHOOTING.md` |
| LaTeX 使用说明 | `references/latex-guide.md` |

## 工作流

### 路线 A：整篇论文（2027 届，推荐）

1. 读取 `references/format-spec.md`，确认格式参数。
2. 让用户提供一份带 YAML 头的 Markdown；字段：
   `title / type / college / grade / major / class_name / student_id / name / advisor / date / abstract_cn / keywords_cn / abstract_en / keywords_en`。
3. 执行：
   ```bash
   bash <skill目录>/make-thesis-2027.sh 论文.md --out <用户指定目录>
   ```
   未指定 `--out` 时，默认输出到当前工作目录。
4. 检查 `05-测试/完整性检查.md` 与 `05-测试/格式自检.md`：
   - 缺文件、缺字段会列为“缺失/提醒”；
   - 目录页码、公式编号、图表编号属于未自动验证项，明确告知用户。
5. 交付时说明：打开 `03-输出/完整论文.docx` 后按 `Ctrl+A`、`F9` 更新目录页码。

### 输出结构

```text
<题目>/
├── 00-原文件/      原手册、转换副本、你的原稿、论文信息.yaml
├── 01-模板/        封面/、附件材料/、样式参考/、参考/
├── 02-封面/        已填写的正文封面、附件材料封面
├── 03-输出/        完整论文.docx、分章 docx（可选 正文.tex）
├── 04-附件材料/    每张表一个子文件夹 + 目录页
└── 05-测试/        完整性检查.md、格式自检.md、测试说明.md
```

### 路线 D：课程报告（市场调查、数据分析等）

有课程格式模版和评价表的课程报告：

给模版：

```bash
bash <skill目录>/make-report-project.sh \
  --report 报告.docx --template 格式模版.docx \
  [--requirements 课程要求.docx] --out <输出目录>
```

不给模版（必须显式指定）：

```bash
bash <skill目录>/make-report-project.sh \
  --report 报告.docx --no-template --out <输出目录>
```

`--template`、`--no-template` 二选一，都不给脚本会拒绝执行并列出三种选择。
「课程报告要求」不是模板，只放 `00-原文件` 和 `04-附件材料`，不放进 `01-模板`。
详见 `references/REPORT-PROJECT.md`。

### 路线 B：只给已有论文刷格式（format-only）

1. 读取 `references/format-spec.md`。
2. 使用 `make-thesis.sh --docx` / `--md`；只改格式，不生成封面。
3. 按 `references/VALIDATION.md` 自检。

### 路线 C：只要封面或附件表

- 从 `<skill目录>/assets/2027/模板/` 复制对应 docx。
- 封面：`封面/封面-毕业论文.docx`、`封面-毕业设计.docx`、`封面-附件材料.docx`。
- 附件表：`附件材料/01-综合成绩评定表.docx` … `11-承诺书.docx`。
- 不在模板上手改字体、页边距；学校原件已带正确格式。

---

## 允许与禁止

**允许**：调整标题层级、目录、页码、字体字号、图表题注、参考文献列表外观；把论文内容填入学校封面与附件表；按项目目录树生成文件；输出格式自检报告；输出 `thesis_data.json` 供复用；复制原手册与原稿到项目 `00-原文件/`。

**禁止**：改写观点、增删论据、编造文献、补充身份信息、修改数据结果、覆盖原文件、覆盖已存在的非空项目目录、伪造查重/AIGC 报告、在 LaTeX 中使用标准 `\section` 替代自定义 `\hA–\hD`、在 python-docx 生成失败时静默降级而不告知用户。
