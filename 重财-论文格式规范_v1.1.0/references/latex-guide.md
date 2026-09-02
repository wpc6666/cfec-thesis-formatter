# LaTeX 模板使用说明

> 模板文件：`assets/chongqing-caijing-thesis.tex`
> 适用：重庆财经学院本科毕业论文（设计）
> 编译引擎：**XeLaTeX**（必须编译两次以生成目录和交叉引用）

---

## 一、编译方式

### 命令行编译
```bash
xelatex main.tex
xelatex main.tex   # 第二次编译生成目录和引用
```

### Overleaf 在线编译
1. 上传 `.tex` 文件到 Overleaf
2. 左上角 Menu → Compiler → 选择 **XeLaTeX**
3. 点击 Recompile（需编译两次）

### TeXstudio / VS Code
- 编译器设置为 XeLaTeX
- 构建工具选择 `xelatex` 或 `latexmk -xelatex`

---

## 二、模板结构与替换指南

| 区域 | 位置 | 替换方式 |
|------|------|---------|
| 封面 | `\begin{titlepage}...\end{titlepage}` | 替换题目、学院、年级等信息 |
| 中文摘要 | `% ---------- 中文摘要 ----------` | 替换摘要正文和关键词 |
| 英文摘要 | `% ---------- 英文摘要 ----------` | 替换 Abstract 和 Keywords |
| 正文 | `% ---------- 第X章 ----------` | 使用 \h1/\h2/\h3/\h4 命令撰写 |
| 图表 | 正文内 figure/table 环境 | 替换图片路径和表格内容 |
| 公式 | equation 环境 | 替换公式内容 |
| 参考文献 | `\begin{therefs}...\end{therefs}` | 按 \item 逐条添加 |
| 附录 | `% ---------- 附录 ----------` | 如有则保留，无则删除 |

---

## 三、自定义命令速查

### 标题命令（必须使用，不要用标准 \section）

| 命令 | 级别 | 编号样式 | 字体字号 | 对齐 |
|------|------|---------|---------|------|
| `\h1{标题}` | 一级 | 一、二、三… | 黑体小三号加粗 | 居中 |
| `\h2{标题}` | 二级 | （一）（二）… | 黑体四号 | 首行缩进2字符 |
| `\h3{标题}` | 三级 | 1. 2. 3… | 黑体小四号 | 首行缩进2字符 |
| `\h4{标题}` | 四级 | （1）（2）… | 宋体小四号加粗 | 首行缩进2字符 |

**注意**：不要使用标准的 `\section`、`\subsection` 等命令，必须使用上述自定义命令，否则编号样式和目录会出错。

### 正文段落
直接在标题命令后写正文即可，模板已自动设置：
- 宋体（中文）/ Latin Modern（英文），小四号
- 首行缩进 2 字符
- 两端对齐
- 固定行距 26 磅

### 图表

**图片**（图题在下方）：
```latex
\begin{figure}[htbp]
\centering
\includegraphics[width=0.8\textwidth]{图片路径}
\caption{图题内容}
\end{figure}
```
自动编号为"图 1-1"格式（章节号-序号）。

**表格**（表题在上方，推荐三线表）：
```latex
\begin{table}[htbp]
\centering
\caption{表题内容}
\begin{tabular}{ccc}
\toprule
列1 & 列2 & 列3 \\
\midrule
数据 & 数据 & 数据 \\
\bottomrule
\end{tabular}
\end{table}
```
自动编号为"表 1-1"格式。

### 公式
```latex
\begin{equation}
E = mc^2
\end{equation}
```
自动编号为（1.1）格式，右对齐。

### 参考文献引用
正文中引用使用 `\cite{key}`，会自动生成上标 [1]。
参考文献列表使用 `therefs` 环境，每条用 `\item`：
```latex
\begin{therefs}
\item 作者. 标题[J]. 期刊, 年份, 卷(期): 页码.
\end{therefs}
```

---

## 四、字体设置

### 默认行为
模板使用 ctex 宏包自动选择中文字体：
- `\songti` → 宋体（Windows: SimSun / macOS: Songti SC / Linux: Noto Serif CJK SC）
- `\heiti` → 黑体（Windows: SimHei / macOS: Heiti SC / Linux: Noto Sans CJK SC）
- 英文默认使用 Latin Modern

### 切换为 Times New Roman（英文）
如果系统已安装 Times New Roman 字体，取消模板中这一行的注释：
```latex
% \setmainfont{Times New Roman}
```
改为：
```latex
\setmainfont{Times New Roman}
```

### Windows 下强制使用 SimSun/SimHei
在模板导言区添加：
```latex
\setCJKmainfont{SimSun}
\setCJKsansfont{SimHei}
```

---

## 五、页码规则（模板已自动处理）

| 部分 | 页码格式 | 实现方式 |
|------|---------|---------|
| 封面 | 无页码 | `\thispagestyle{empty}` |
| 摘要/Abstract/目录 | 大写罗马数字 Ⅰ Ⅱ Ⅲ | `\pagenumbering{Roman}` |
| 正文/参考文献/附录 | 阿拉伯数字 1 2 3 | `\pagenumbering{arabic}` |

页码位置：页面底端居中，小四号。

---

## 六、目录（模板已自动处理）

- 只显示到二级标题（`\setcounter{tocdepth}{2}`）
- 一级项无缩进，二级项左缩进 2 字符
- 页码右对齐，点号前导符
- 目录标题"目录"两字间空 2 个半角空格
- **必须编译两次**才能正确生成目录

---

## 七、常见问题

### Q1: 编译报错 "Font ... not found"
A: 系统缺少对应中文字体。安装 Noto CJK 字体或使用 Overleaf（已预装字体）。

### Q2: 目录不显示或页码不对
A: 再编译一次 XeLaTeX。目录和交叉引用需要两次编译。

### Q3: 图表编号不是"图 1-1"格式
A: 确认使用了 `\h1{}` 命令开始新章节，不要用 `\section{}`。图表计数器在每个 `\h1` 后自动重置。

### Q4: 行距不是 26 磅
A: 模板已通过 `\AtBeginDocument{\setlength{\baselineskip}{26pt}}` 设置。如果某些宏包覆盖了此设置，在 `\begin{document}` 后第一行添加 `\setlength{\baselineskip}{26pt}`。

### Q5: 如何添加附录 B
A: 复制附录 A 的整个块，将"附录 A"改为"附录 B"，标题和内容替换即可。

### Q6: 参考文献需要 GB/T 7714 格式
A: 本模板不强制参考文献著录格式，按学校要求手动撰写每条 `\item` 即可。如需自动化，可引入 `biblatex-gb7714-2015` 宏包。

---

## 八、与 Word 版的对应关系

| 规范项 | Word 实现 | LaTeX 实现 |
|--------|----------|-----------|
| 页边距左2.8cm | 页面设置 | `geometry` 宏包 |
| 固定行距26磅 | 段落设置 | `\setlength{\baselineskip}{26pt}` |
| 四级标题编号 | 多级列表 | `\h1`–`\h4` 自定义命令 |
| 目录只到二级 | TOC域设置 | `\setcounter{tocdepth}{2}` |
| 图 1-1 编号 | 题注章节编号 | `\renewcommand{\thefigure}` |
| 公式（1.1） | 公式编号 | `\numberwithin{equation}` |
| 罗马/阿拉伯页码 | 分节符 | `\pagenumbering{Roman/arabic}` |
| 参考文献悬挂缩进 | 段落悬挂缩进 | `therefs` 环境 + enumitem |
| 引用上标 | 字体上标 | `\cite{}` 自动上标 |
