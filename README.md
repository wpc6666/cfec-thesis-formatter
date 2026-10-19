# 重财-论文格式规范

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Version](https://img.shields.io/badge/Version-2.2.4-6366f1.svg)](CHANGELOG.md)
[![Output](https://img.shields.io/badge/Output-DOCX%20%2B%20LaTeX-1a7f37.svg)](#%EF%B8%8F-核心特性)
[![Pages](https://img.shields.io/badge/GitHub%20Pages-在线文档-254b9c.svg)](https://wpc6666.github.io/cfec-thesis-formatter/)

> 重庆财经学院本科毕业论文自动化排版工具 —— 一键生成 DOCX + LaTeX 双格式成品
> **只调整格式，不改写论文内容。**

**在线文档**：<https://wpc6666.github.io/cfec-thesis-formatter/>

---

## 🎯 这是什么

一个 Agent Skill。把论文正文丢给 AI 助手，
它按重庆财经学院《毕业论文（设计）指导手册》的格式规范，
生成两份可直接提交的成品：

- **DOCX** — 用 Word 打开即可编辑，样式全部预设
- **LaTeX** — XeLaTeX 编译成 PDF，格式宏已内置

不用手动调页边距、行距、页码、标题层级、三线表、公式编号、参考文献缩进。

---

## ✨ 核心特性

| 特性 | 说明 |
|---|---|
| **双格式输出** | DOCX（python-docx 原生引擎）+ LaTeX（ctexart + XeLaTeX），格式一致 |
| **版本化输出** | 每次生成独立 `v<N>/` 目录，自动递增，历史输出永不覆盖 |
| **域自动更新** | DOCX 注入 `updateFields`，Word 打开即自动更新目录与页码 |
| **目录树可追溯** | 项目根目录生成 `MANIFEST.md`，记录全部版本的时间与文件清单 |
| **严格复刻规范** | 页边距、行距、页码、标题层级、参考文献格式逐项对齐手册 |
| **四级标题体系** | `一、` → `（一）` → `1.` → `(1)` 自动排版 |
| **图表公式编号** | 图 1-1 / 表 1-1 按章节编号；公式右对齐编号；支持三线表 |
| **参考文献规范** | 悬挂缩进 2 字符；正文引用自动上标 `[1]`；GB/T 7714-2015 |
| **不覆盖原件** | 默认不改原始文件，输出前自动校验关键格式项 |
| **多端兼容** | DeepSeek Harness / Codex / Claude Code 等主流 Agent 平台 |

---

## 📦 安装

### 方法一：Git 克隆

```bash
git clone https://github.com/wpc6666/cfec-thesis-formatter.git
cd cfec-thesis-formatter

# 把 Skill 目录复制到你的 Agent 技能目录
cp -R "重财-论文格式规范" ~/.harness/skills/
```

> 仓库内只保留**一个最新版** Skill 目录（`重财-论文格式规范/`）。
> 版本号记录在 `SKILL.md` frontmatter 与 `manifest.json` 里，
> 不再使用带版本号的目录快照。

### 方法二：插件市场

在 Codex / DeepSeek Harness 插件市场里添加本仓库地址，即可一键安装。

> ⚠️ 安装后需重启 Agent 服务，技能才会被索引加载。

### 依赖

| 依赖 | 用途 | 安装 |
|---|---|---|
| Python 3.9+ | DOCX 生成引擎 | 系统自带或 `brew install python` |
| `python-docx` | 生成 `.docx` | `pip install python-docx` |
| XeLaTeX | 编译 `.tex` → PDF | `brew install --cask mactex` 或 TeX Live |

---

## 🤖 快速调用

把下面这段话直接发给 AI 助手：

```
请安装 cfec-thesis-formatter 这个 Skill。
仓库地址：https://github.com/wpc6666/cfec-thesis-formatter

安装完成后，使用这个 Skill 处理我的毕业论文，
输出符合重庆财经学院本科格式规范的 DOCX 和 XeLaTeX 两份文件。
只调整格式，不要修改论文内容。
```

### 调用示例

```
使用「重财-论文格式规范」Skill 处理以下论文：

【粘贴论文正文】

要求：
1. 输出 Word 文档和 LaTeX 源文件
2. 严格按照重庆财经学院格式规范
3. 保留所有原文内容不变
4. 输出到版本化目录（项目目录/v<N>/），不要覆盖之前的版本
```

---

## 🛒 作为 Marketplace 接入

仓库已配置标准 marketplace 结构，可直接作为插件市场源接入。

### 目录结构

```
cfec-thesis-formatter/
├── .agents/
│   └── plugins/
│       └── marketplace.json    市场清单（单一插件注册）
├── 重财-论文格式规范/            唯一 Skill 目录（最新版 v2.2.0）
│   ├── SKILL.md                核心技能文件
│   ├── manifest.json           版本与资源清单
│   ├── CHANGELOG.md            更新日志
│   ├── references/             格式规范文档
│   ├── assets/                 LaTeX 模板
│   └── scripts/                python-docx 引擎 + 版本管理器
├── index.html                  在线文档
└── README.md
```

### 接入步骤

1. 确认仓库根目录存在 `.agents/plugins/marketplace.json`
2. 在 Harness Marketplace 配置中添加本仓库地址作为源
3. 刷新市场列表，一键安装

---

## 📖 使用流程

1. 准备论文**纯文本**，不需要自己调复杂格式
2. 确认 Agent 已加载本 Skill
3. 把论文内容发给 Agent，指定用本 Skill 处理
4. Agent 按规范生成版本化目录 + DOCX + TeX
5. 本地打开核对图表、参考文献，直接提交

### 输入建议

- 用纯文本，格式越干净输出越标准
- 标题层级用明确标记：`一、` `（一）` `1.` `(1)`
- 参考文献单独列出，遵循 GB/T 7714
- 图片单独提供文件，Skill 会自动插入并编号

### 输出说明

| 产物 | 说明 |
|---|---|
| `v<N>/*.docx` | Word 文档，样式已预设，打开即用 |
| `v<N>/*.tex` | XeLaTeX 源文件，编译即出 PDF |
| `v<N>/*_data.json` | 结构化数据，便于二次生成 |
| `MANIFEST.md` | 项目根目录，记录全部版本清单 |

---

## 📂 输出目录树

```
output/
└── 基于深度学习的X研究/           项目目录（题目 slug 化）
    ├── v1/
    │   ├── 基于深度学习的X研究_v1.docx
    │   ├── 基于深度学习的X研究_v1.tex
    │   └── 基于深度学习的X研究_v1_data.json
    ├── v2/                        第 2 次输出（自动递增，不覆盖 v1）
    │   ├── 基于深度学习的X研究_v2.docx
    │   ├── 基于深度学习的X研究_v2.tex
    │   └── 基于深度学习的X研究_v2_data.json
    └── MANIFEST.md                版本清单
```

> 版本号规则：默认取「已有最大版本号 + 1」。
> 可指定版本号，但目标目录非空时拒绝写入（防误覆盖）。

---

## 🔖 版本管理

- **Skill 版本**：记录于 `SKILL.md` frontmatter 与 `manifest.json`，语义化版本
- **发布标记**：Git tag（如 `v2.2.0`）
- **历史记录**：[CHANGELOG.md](CHANGELOG.md)
- **输出迭代**：`v<N>/` 目录与 Skill 版本号**相互独立**，
  Skill 升级不影响已生成的论文版本

### 版本史

| 版本 | 主要变化 |
|---|---|
| v2.2.0 | 新增 Skill 自身发布自动化（`scripts/release.py`） |
| v2.1.0 | 域自动更新 + 类型子文件夹输出（word/ latex/ data/） |
| v2.0.1 | 细节修正 |
| v2.0.0 | DOCX + LaTeX 双引擎重构 |

---

## 📐 格式规范明细

### 页面设置

| 项目 | 规范 |
|---|---|
| 纸张 | A4（210 × 297 mm） |
| 页边距 | 左 2.8 cm；上 / 下 / 右 2.5 cm |
| 行距 | 全文固定 26 磅 |
| 正文字体 | 宋体小四号 |
| 标题字体 | 黑体 |

### 页码规则

| 部分 | 页码 |
|---|---|
| 封面 | 无页码 |
| 摘要、目录 | 大写罗马数字（Ⅰ、Ⅱ、Ⅲ…），底部居中 |
| 正文 | 阿拉伯数字（1、2、3…），底部居中 |

### 标题层级

| 层级 | 格式 | 字体字号 | 对齐 |
|---|---|---|---|
| 一级 | `一、` | 黑体三号 | 居中 |
| 二级 | `（一）` | 黑体四号 | 左对齐 |
| 三级 | `1.` | 黑体小四号 | 左对齐 |
| 四级 | `(1)` | 宋体小四号加粗 | 左对齐 |

### 图表与公式

- **图号**：`图 1-1`，图题在图**下方**，宋体五号
- **表号**：`表 1-1`，表题在表**上方**，宋体五号
- **表格**：推荐三线表
- **公式**：`(1.1)` 右对齐，按章节编号

### 参考文献

- 悬挂缩进 2 字符
- 宋体五号，单倍行距
- 正文引用处自动上标 `[1]`
- 遵循 GB/T 7714-2015

---

## ❓ 常见问题

**Q：会修改我的论文内容吗？**
不会。只调整格式样式，不改动任何文字、观点、数据和参考文献内容。

**Q：重复格式化会覆盖之前的输出吗？**
不会。每次都进入新的 `v<N>/` 目录，历史版本与 `MANIFEST.md` 完整保留。

**Q：支持其他学校吗？**
当前版本专为重庆财经学院本科毕业论文设计，其他院校需定制开发。

**Q：安装后调用没反应？**
检查 Skill 是否被正确索引，重启 Agent 服务，确认 `SKILL.md` 路径正确。

**Q：可以批量处理多篇论文吗？**
可以。每篇按题目建独立项目目录，互不干扰。

**Q：LaTeX 怎么出 PDF？**
用 XeLaTeX 引擎编译。推荐 TeX Live 或 MiKTeX。**不要用 pdfLaTeX**，中文会掉字。

**Q：Word 里的目录是空的？**
打开后按 `Ctrl+A` 再按 `F9`（Mac 用 `Cmd+A` + `F9`）手动更新一次域。
  v2.1.0 起已注入 `updateFields`，多数情况会自动更新。

---

## 🔗 相关链接

| 用途 | 链接 |
|---|---|
| 在线文档 | <https://wpc6666.github.io/cfec-thesis-formatter/> |
| 问题反馈 | <https://github.com/wpc6666/cfec-thesis-formatter/issues> |
| 更新日志 | [CHANGELOG.md](CHANGELOG.md) |
| 个人主页 | <https://wpc6666.github.io/> |

---

## 📄 许可

[MIT License](LICENSE) — 可自由使用、修改、分发。

作者：白澤 程
