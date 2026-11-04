# 2027 一键生成器

## 自带资源

模板、原手册、参考样式都随 Skill 分发，换电脑只需要：

1. 安装本 Skill；
2. 安装 `pandoc`；
3. 有 Python 3（首次运行自动建隔离环境并装 `python-docx` / `PyYAML` / `lxml`）。

不需要另外的论文工具目录，也不需要手动准备模板。

## 命令

```bash
bash <skill目录>/make-thesis-2027.sh 论文.md --out <输出目录>
```

- `--out`：项目根目录；省略时输出到当前工作目录。
- `--type`：`毕业论文`（默认）或 `毕业设计`。
- `--latex`：同时生成正文 LaTeX（不含封面、附件表）。
- 目标目录已存在且非空时拒绝执行，不覆盖。
- 首次运行会在 Skill 目录下建 `.venv`；可用 `CFEC_VENV` 指定其他位置。

## 输入

一份带 YAML 头的 Markdown。最少字段：

```yaml
---
title: "题目"
type: "毕业论文"
college: "经济学院"
grade: "2023级"
major: "经济统计学"
class_name: "23经统一"
student_id: "2023xxxx"
name: "张三"
advisor: "李四、王五"
date: "2027年5月"
abstract_cn: |
  中文摘要……
keywords_cn: "关键词1；关键词2；关键词3"
abstract_en: |
  English abstract...
keywords_en: "Keyword1; Keyword2; Keyword3"
---
```

缺字段不会中断生成，但会写进 `05-测试/格式自检.md` 的“提醒”。

## 正文写法

- `# 标题` → 一、二、三、……
- `## 标题` → （一）（二）（三）……
- `### 标题` → 1. 2. 3. ……
- `#### 标题` → （1）（2）（3）……
- `# 参考文献` → 不编号，自动另起一页
- `# 附录A` + 下一段标题 → 不编号，自动另起一页；下一段自动作为附录标题（黑体四号居中）
- 图表题注：`图3-1  长方形`、`表3-1  近两年设计情况统计表`

生成器会自动去掉你手写的旧编号再重排，避免“一、一、绪论”。

## 输出

```text
<题目>/
├── 00-原文件/
│   ├── 附件4：…（2027届）.doc     只读副本
│   ├── 手册-2027（Word转换版）.docx
│   ├── 原稿-<文件名>.md
│   └── 论文信息.yaml
├── 01-模板/
│   ├── 封面/                       毕业论文、毕业设计、附件材料
│   ├── 附件材料/                   11 张空白表
│   ├── 样式参考/                   摘要、目录、正文、参考文献、附录样例
│   └── 参考/                       专业班级对照表 docx + csv
├── 02-封面/
│   ├── 正文封面.docx
│   └── 附件材料封面.docx
├── 03-输出/
│   ├── 完整论文.docx               封面 → 摘要 → Abstract → 目录 → 正文 → 参考文献 → 附录
│   ├── 01-中文摘要.docx
│   ├── 02-英文摘要.docx
│   ├── 03-正文.docx
│   ├── 04-参考文献.docx
│   ├── 05-附录.docx
│   └── 正文.tex                    仅 --latex 时生成
├── 04-附件材料/
│   ├── 00-附件材料目录/
│   ├── 01-综合成绩评定表/ … 11-承诺书/    每张表一个子文件夹
│   ├── 12-查重检测报告/说明.md
│   └── 13-AIGC检测报告/说明.md
└── 05-测试/
    ├── 完整性检查.md               逐项通过/缺失
    ├── 格式自检.md                 字段、待办、未验证项
    ├── 测试说明.md                 人工检查步骤
    ├── 完整论文预览.pdf            自动导出：优先 Word，其次 soffice
    └── 中间文件/                   前置、正文混排、拆分件
```

## 分节与页码

`03-输出/完整论文.docx` 有三节：

1. 封面：A4，无页脚、无页码。
2. 摘要 / Abstract / 目录：大写罗马数字，从 I 开始。
3. 正文 / 参考文献 / 附录：阿拉伯数字，从 1 开始。

目录是 TOC 域，打开后需 `Ctrl+A`、`F9` 更新页码。生成时设置 `updateFields=true`，多数 Word 会自动提示更新。

## 模板与重新提取

- 手册原文：`<skill目录>/assets/2027/原文件/`
- 提取模板：`<skill目录>/assets/2027/模板/`
- 重新提取：
  ```bash
  cd <skill目录>
  .venv/bin/python scripts/2027/extract_templates.py \
    --manual assets/2027/原文件/手册-2027（Word转换版）.docx
  ```
- 按节号切分；手册改版时同步改 `extract_templates.py` 的 `COVERS / FORMS / SAMPLES / MULTI` 表。

## 实现结构

| 文件 | 作用 |
|---|---|
| `scripts/2027/make_project.py` | CLI、目录结构、附件表分发、自检 |
| `scripts/2027/thesis2027.py` | 封面填充、摘要、正文排版、拆分、完整论文装配 |
| `scripts/2027/docx_merge.py` | 按节合并 docx，保留各自页脚与页码 |
| `scripts/2027/extract_templates.py` | 从手册 docx 按节提取模板 |
| `scripts/fix_tex.py` | 正文 LaTeX 的参考文献/附录目录修正 |
| `assets/thesis-reference.docx` | pandoc 参考样式 |
| `assets/body-template.tex` | 正文 LaTeX 模板 |

## Word 自动导出

首次导出会触发 macOS “允许控制 Microsoft Word”；也可运行
`bash <skill目录>/authorize-word.sh` 主动授权。拒绝过就去
系统设置 → 隐私与安全性 → 自动化 → 勾选运行脚本的 App 下的 Microsoft Word。
未授权时跳过 PDF，在 `05-测试/完整性检查.md` 的提醒里说明。

## 已知边界

- 目录页码必须用 Word 更新域；离线无法计算。
- 公式按章编号不自动生成，需在 Word 里手工右对齐。
- 图表编号不自动递增，按题注文字排版。
- 封面标题使用学校模板的「方正小标宋」；本机缺字体时 Word 会替换。
- 附件表中的签字、日期、盖章、评语内容不代写。
- 查重、AIGC 报告由学校系统导出，不提供仿制模板。
