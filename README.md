# 重财-论文格式规范工具
> 重庆财经学院本科毕业论文自动化排版工具 | 一键生成 DOCX + LaTeX(XeLaTeX) 双格式成品

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Output: DOCX+LaTeX](https://img.shields.io/badge/Output-DOCX%20%2B%20LaTeX-brightgreen.svg)]()
[![Version: 2.0.0](https://img.shields.io/badge/Version-2.0.0-6366f1.svg)]()
[![GitHub Pages](https://img.shields.io/badge/GitHub-Pages-254b9c.svg)](https://wpc6666.github.io/cfec-thesis-formatter/)

---

## 📑 目录
- [🏫 关于项目](#-关于项目)
- [✨ 项目特性](#-项目特性)
- [📦 安装 Skill](#-安装-skill)
- [🤖 快速调用 Agent](#-快速调用-agent)
- [🛒 作为 Marketplace 接入](#-作为-marketplace-接入)
- [📖 使用说明](#-使用说明)
- [📂 输出目录树](#-输出目录树)
- [🔖 版本管理](#-版本管理)
- [📐 格式规范明细](#-格式规范明细)
- [❓ 常见问题](#-常见问题)
- [📄 许可证](#-许可证)

---

## 🏫 关于项目
重庆财经学院是重庆地区办学质量领先的民办本科院校之一，本项目专为该校本科毕业生打造，**严格遵循学校《毕业论文（设计）指导手册》格式规范**，通过 Agent Skill 一键生成符合要求的 Word 文档与 LaTeX 源文件。

> **设计理念：只处理格式，绝不修改论文正文、观点、数据与参考文献，保证学术完整性。**

项目主页：https://wpc6666.github.io/cfec-thesis-formatter/

---

## ✨ 项目特性
- **双格式输出**：一键生成 DOCX(Word) 与 TeX(LaTeX-XeLaTeX) 两份成品，格式完全一致
- **版本化输出，永不覆盖**：按论文题目建立项目目录，每次格式化自动递增 `v1/v2/v3…` 版本目录，文件名携带版本号，历史输出全部保留
- **目录树可追溯**：项目根目录自动生成 `MANIFEST.md` 版本清单，交付时附完整目录树
- **严格复刻规范**：完整复刻学校官方页边距、行距、页码、标题层级、参考文献格式
- **页码规则适配**：前置部分（摘要、目录）大写罗马数字；正文阿拉伯数字；封面无页码
- **四级标题体系**：`一、 → （一） → 1. → (1)` 完整层级自动排版
- **图表公式编号**：图 1-1 / 表 1-1 编号规范；公式右对齐按章节编号；支持三线表
- **参考文献规范**：悬挂缩进2字符；正文引用自动上标 `[1]`；遵循 GB/T 7714 格式
- **安全交付机制**：默认不覆盖原始文件，输出前自动校验关键格式项
- **多端兼容**：支持 DeepSeek Harness、Codex 等主流 Agent 平台

---

## 📦 安装 Skill

### 方法一：Git 克隆本地安装
```bash
# 克隆仓库
git clone https://github.com/wpc6666/cfec-thesis-formatter.git
cd cfec-thesis-formatter

# 将最新版 Skill 目录复制到你的 Agent 技能目录
cp -R "重财-论文格式规范" ~/.harness/skills/
```

> 仓库内只保留一个最新版 Skill 目录（`重财-论文格式规范/`），版本号在 `SKILL.md` frontmatter 与 `manifest.json` 中记录，不再使用带版本号的目录快照。

### 方法二：插件市场一键安装
在 Codex / DeepSeek-Harness 插件市场中添加本仓库地址，即可一键安装 Skill。

> ⚠️ 安装完成后重启 Agent 服务，技能才会被索引加载。

---

## 🤖 快速调用 Agent
复制下面这段话，直接发送给你的 Agent / 助手，即可自动安装并调用本技能。

```
请安装 cfec-thesis-formatter 这个 Skill。
仓库地址：https://github.com/wpc6666/cfec-thesis-formatter

安装完成后，请使用这个 Skill 处理我的毕业论文，
输出符合重庆财经学院本科格式规范的 DOCX 和 XeLaTeX 两份文件。
只调整格式，不要修改论文内容。
```

### 调用示例
```
使用重财论文格式规范 Skill，处理以下论文内容：

【粘贴你的论文正文】

要求：
1. 输出 Word 文档和 LaTeX 源文件
2. 严格按照重庆财经学院格式规范
3. 保留所有原文内容不变
4. 输出到版本化目录（项目目录/v<N>/），不要覆盖我之前的版本
```

---

## 🛒 作为 Marketplace 接入
本仓库已配置标准 marketplace 结构，可直接作为插件市场源接入 Harness / Codex。

### 目录结构
```
cfec-thesis-formatter/
├── .agents/
│   └── plugins/
│       └── marketplace.json    # 市场清单文件（单一插件注册）
├── 重财-论文格式规范/           # 唯一 Skill 目录（最新版 v2.0.0）
│   ├── SKILL.md                # 核心技能文件
│   ├── manifest.json           # 版本与资源清单
│   ├── CHANGELOG.md            # 更新日志
│   ├── references/             # 格式规范文档
│   ├── assets/                 # LaTeX 模板
│   └── scripts/                # python-docx 引擎 + 输出版本管理器
└── README.md
```

### marketplace.json 配置
```json
{
  "name": "cfec-thesis-formatter",
  "interface": { "displayName": "重财论文排版工具" },
  "plugins": [
    {
      "name": "thesis-formatter",
      "source": { "source": "local", "path": "./重财-论文格式规范" },
      "policy": { "installation": "AVAILABLE", "authentication": "ON_INSTALL" },
      "category": "Productivity"
    }
  ]
}
```

### 接入步骤
1. 确认仓库根目录存在 `.agents/plugins/marketplace.json`
2. 在 Harness Marketplace 配置中添加本仓库地址作为源
3. 刷新市场列表，即可看到本工具并一键安装 Skill

---

## 📖 使用说明

### 基础使用流程
1. 准备你的论文纯文本，建议不要自行调整复杂格式
2. 确保 Agent 已加载本项目 Skill
3. 向 Agent 发送论文内容，指定使用本 Skill 处理
4. Agent 按输出目录树规范生成版本化目录与标准 Word(DOCX)、XeLaTeX 源文件
5. 本地打开预览，核对图表、参考文献后直接用于提交

### 输入建议
- 建议输入纯文本内容，格式越干净，输出越标准
- 标题使用明确的层级标记：`一、` `（一）` `1.` `(1)`
- 参考文献单独列出，格式为标准 GB/T 7714
- 图片请单独提供文件，Skill 会自动插入并编号

### 输出说明
- **版本化目录**：每次输出进入 `<题目slug>/v<N>/`，版本号自动递增，绝不覆盖历史版本
- **DOCX**：可直接用 Word 打开编辑，样式已全部预设
- **TeX**：XeLaTeX 编译即可生成 PDF，已内置完整格式宏
- **MANIFEST.md**：项目根目录自动记录全部版本的生成时间与文件清单

---

## 📂 输出目录树

同一篇论文多次格式化后，输出结构如下（示例题目：基于深度学习的X研究）：

```
output/
└── 基于深度学习的X研究/           # 项目目录（题目 slug 化）
    ├── v1/                        # 第 1 次输出
    │   ├── 基于深度学习的X研究_v1.docx
    │   ├── 基于深度学习的X研究_v1.tex
    │   └── 基于深度学习的X研究_v1_data.json
    ├── v2/                        # 第 2 次输出（自动递增，不覆盖 v1）
    │   ├── 基于深度学习的X研究_v2.docx
    │   ├── 基于深度学习的X研究_v2.tex
    │   └── 基于深度学习的X研究_v2_data.json
    └── MANIFEST.md                # 版本清单
```

版本号规则：默认自动取「已有最大版本号 + 1」；用户可指定版本号，但目标目录非空时拒绝写入（防误覆盖）。

---

## 🔖 版本管理

- **Skill 版本**：记录于 `SKILL.md` frontmatter 与 `manifest.json`，遵循语义化版本（主.次.修订）。
- **发布版本**：使用 Git tag（如 `v2.0.0`）标记，仓库内不再存放带版本号的目录快照。
- **历史记录**：见 `重财-论文格式规范/CHANGELOG.md`。
- **输出迭代**：每次格式化生成独立的 `v<N>/` 版本目录（见上节输出目录树），与 Skill 版本号相互独立——Skill 升级不影响已生成的论文版本。

---

## 📐 格式规范明细

### 页面设置
- 纸张：A4 (210mm × 297mm)
- 页边距：左 2.8cm，上/下/右 2.5cm
- 行距：全文固定 26 磅
- 字体：正文宋体小四号，标题黑体

### 页码规则
- 封面：无页码
- 摘要、目录：大写罗马数字（Ⅰ、Ⅱ、Ⅲ...），底部居中
- 正文：阿拉伯数字（1、2、3...），底部居中

### 标题层级
| 层级 | 格式 | 字体字号 | 对齐 |
|------|------|----------|------|
| 一级 | 一、 | 黑体三号 | 居中 |
| 二级 | （一） | 黑体四号 | 左对齐 |
| 三级 | 1. | 黑体小四号 | 左对齐 |
| 四级 | (1) | 宋体小四号加粗 | 左对齐 |

### 图表与公式
- 图号：`图 1-1`，图题在图下方，宋体五号
- 表号：`表 1-1`，表题在表上方，宋体五号
- 表格：推荐使用三线表
- 公式：`(1.1)` 右对齐，按章节编号

### 参考文献
- 悬挂缩进 2 字符
- 宋体五号，单倍行距
- 正文引用处自动上标 `[1]`
- 遵循 GB/T 7714-2015 格式

---

## ❓ 常见问题

**Q1：会修改我的论文内容吗？**
绝对不会。本 Skill 只调整格式样式，不改动任何文字、观点、数据和参考文献内容。

**Q2：重复格式化会不会覆盖我之前的输出？**
不会。每次格式化自动进入新的 `v<N>/` 版本目录，文件名携带版本号，历史版本与 MANIFEST.md 清单完整保留。

**Q3：支持哪些学校格式？**
当前版本专为重庆财经学院本科毕业论文设计。其他院校可定制开发。

**Q4：安装后调用没反应？**
检查 Skill 是否被正确索引，重启 Agent 服务，确认 SKILL.md 路径配置正确。

**Q5：可以批量处理多篇论文吗？**
可以。每篇论文按题目建立独立项目目录，互不干扰。

**Q6：LaTeX 文件怎么生成 PDF？**
使用 XeLaTeX 引擎编译即可，推荐使用 TeX Live 或 MiKTeX 环境。

> 更多问题欢迎在 [GitHub Issues](https://github.com/wpc6666/cfec-thesis-formatter/issues) 提交反馈。

---

## 📄 许可证
本项目基于 **MIT License** 开源，可自由使用、修改、分发。

作者：wpc6666  
项目仓库：https://github.com/wpc6666/cfec-thesis-formatter
