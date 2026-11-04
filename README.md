# 重财-论文格式规范

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Version](https://img.shields.io/badge/Version-3.5.0-6366f1.svg)](CHANGELOG.md)
[![Manual](https://img.shields.io/badge/Manual-2027%E5%B1%8A-1a7f37.svg)](https://wpc6666.github.io/cfec-thesis-formatter/)
[![Pages](https://img.shields.io/badge/GitHub%20Pages-在线文档-254b9c.svg)](https://wpc6666.github.io/cfec-thesis-formatter/)

> 重庆财经学院 2027 届本科毕业论文/课程报告排版工具。
> **只改格式，不改内容。** Skill 自带学校手册、封面和附件表，换电脑不用另外找模板。

**在线文档**：<https://wpc6666.github.io/cfec-thesis-formatter/>

---

## 这个 Skill 能做什么

- **毕业论文**：封面 → 中文摘要 → Abstract → 目录 → 正文 → 参考文献 → 附录，
  一次生成完整项目文件夹，并带 11 张学校附件表。
- **课程报告**：给课程格式模版时切出封面和评价表；不给模版时只输出正文，不套论文模板。
- **已有文档**：format-only 只刷字体、字号、行距、标题、三线表、页码，不生成封面。
- **PDF 预览**：macOS 自动调用 Microsoft Word 导出；没有 Word 时回退 soffice。

## 核心特性

| 特性 | 说明 |
|---|---|
| **2027 手册** | 页面、字体、页码、标题、图表、参考文献、附录逐项对齐 2027 届指导手册 |
| **自带模板** | `assets/2027/` 内含原手册、3 个封面、11 张附件表、样式样例、班级对照表 |
| **一键项目** | 输出 `00-原文件 / 01-模板 / 02-封面 / 03-输出 / 04-附件材料 / 05-测试` |
| **附件材料** | 每张表一个子文件夹，自动填姓名、学号、专业、题目 |
| **课程报告** | 支持「给模版」和「不给模版」两种模式，二选一，不默认套论文模板 |
| **Word 导出** | 自动授权 macOS 自动化，Word 容器内导出 PDF，减少文件权限弹窗 |
| **不覆盖** | 目标目录非空拒绝写入；原文件只读复制 |
| **可追溯** | `05-测试/完整性检查.md` 记录通过/缺失项与未验证项 |

## 安装

需要 Python 3、pandoc；macOS 用 Word 导出 PDF 时还需要 Microsoft Word。

```bash
git clone https://github.com/wpc6666/cfec-thesis-formatter.git
cp -R cfec-thesis-formatter/重财-论文格式规范 ~/.agents/skills/cfec-thesis-formatter
```

首次运行会在 Skill 目录下创建 `.venv`，安装 `python-docx` / `PyYAML` / `lxml`。

## 用法

### 毕业论文（2027 届）

```bash
bash ~/.agents/skills/cfec-thesis-formatter/make-thesis-2027.sh \
  论文.md --out ~/IT/论文
```

输入是一份带 YAML 头的 Markdown：

```yaml
---
title: "数字经济对区域经济增长的影响研究"
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
keywords_cn: "数字经济；区域经济增长；面板数据"
abstract_en: |
  English abstract...
keywords_en: "Digital economy; Regional economic growth; Panel data"
---
```

正文用 `#` / `##` / `###` / `####` 写标题，生成器自动编 `一、` / `（一）` / `1.` / `（1）`。
`# 参考文献`、`# 附录A` 自动另起一页。

### 课程报告

给模版：

```bash
bash ~/.agents/skills/cfec-thesis-formatter/make-report-project.sh \
  --report 报告.docx --template 格式模版.docx \
  [--requirements 课程要求.docx] --out ~/Downloads
```

不给模版（必须显式指定）：

```bash
bash ~/.agents/skills/cfec-thesis-formatter/make-report-project.sh \
  --report 报告.docx --no-template --out ~/Downloads
```

`--template` 与 `--no-template` 二选一；都不给会拒绝执行，并提示三种选择。

### 只刷已有文档格式

```bash
bash ~/.agents/skills/cfec-thesis-formatter/format-only.sh 输入.docx
bash ~/.agents/skills/cfec-thesis-formatter/format-only.sh 输入.docx -o 输出.docx
```

## 输出结构

```text
<题目>/
├── 00-原文件/       原手册、转换副本、你的原稿、论文信息.yaml
├── 01-模板/         封面/、附件材料/、样式参考/、参考/
├── 02-封面/         已填写的正文封面、附件材料封面
├── 03-输出/         完整论文.docx、分章 docx（可选 正文.tex）
├── 04-附件材料/     每张表一个子文件夹
└── 05-测试/         完整性检查、格式自检、PDF 预览
```

## Word 自动导出与授权（macOS）

首次导出 PDF 时，macOS 会弹「允许控制 Microsoft Word」，点允许即可。也可先运行：

```bash
bash ~/.agents/skills/cfec-thesis-formatter/authorize-word.sh
```

如果之前拒绝过：

```text
系统设置 → 隐私与安全性 → 自动化
→ 勾选运行脚本的 App（终端 / Claude Code / ChatGPT / Codex）下的 Microsoft Word
```

macOS 不允许脚本静默写入 TCC 授权；授权一次后永久生效。

## 目录结构

```text
cfec-thesis-formatter/
├── 重财-论文格式规范/
│   ├── SKILL.md
│   ├── manifest.json
│   ├── CHANGELOG.md
│   ├── make-thesis-2027.sh
│   ├── make-report-project.sh
│   ├── authorize-word.sh
│   ├── format-only.sh
│   ├── make-thesis.sh
│   ├── assets/
│   │   ├── 2027/原文件/
│   │   ├── 2027/模板/
│   │   ├── body-template.tex
│   │   ├── thesis-reference.docx
│   │   └── format.yaml
│   ├── references/
│   └── scripts/
│       ├── 2027/（论文生成器）
│       ├── make_report_project.py
│       ├── build_word.py
│       ├── format_md.py
│       ├── format_docx.py
│       ├── ref_zh.py
│       ├── make-full-thesis.sh
│       ├── word-export.sh
│       └── pdf_preview.py
├── index.html
└── README.md
```

## 已知边界

- 目录页码需在 Word 中按 `Ctrl+A`、`F9` 更新域。
- 公式按章编号、图表编号不自动递增，按题注文字排版。
- 课程报告的「第×页（共×页）」页码格式以课程要求为准，当前用通用页码。
- 查重、AIGC 报告由学校系统导出，不提供仿制模板。
- 封面「方正小标宋」缺字体时 Word 会替换。

## 版本

见 [CHANGELOG.md](重财-论文格式规范/CHANGELOG.md)。当前 v3.5.0。

## License

MIT
