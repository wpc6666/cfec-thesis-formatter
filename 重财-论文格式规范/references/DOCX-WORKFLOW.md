# DOCX 工作流（旧版引擎）

> 2027 主流程用 `make-thesis-2027.sh`；本文档描述 `scripts/cfec_thesis_docx.py` 的 API/版本化输出，供脚本复用。

> 上层：`../SKILL.md`；按需读取本文件。

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
- 四级标题、图、表、公式：`build()` **不会自动解析标记**；需要显式调用
  `add_h4()` / `add_figure()` / `add_three_line_table()` / `add_equation()`。
- 正文引用 `[1]` → `build()` 用 `add_body_with_citations()` 转上标。

### Step 3：调用 Python 引擎生成 DOCX（版本化输出）

```bash
# 首次运行由 Skill 自动建 venv；手动安装：
.venv/bin/pip install -r scripts/requirements.txt

# 注意：该脚本无输入参数，直接运行只会生成内置 sample_thesis 示例（v1/v2）
# 生成用户论文请调用下面的 build_and_save(thesis_data, ...)
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
| 一级标题黑体小三、不加粗、居中 | `set_run_font(heiti, 15pt, bold=False)` + 居中 |
| 图1-1 编号（图题在下） | 按章节计数器 + `\caption` 位置 |
| 表1-1 编号（表题在上）+ 三线表 | 按章节计数器 + `set_cell_border` 仅顶/中/底三线 |
| 公式（1.1）右对齐 | 按章节计数器 + 右对齐编号 |
| 参考文献悬挂缩进2字符 | `left_indent=2char` + `first_line_indent=-2char` |
| 正文引用 [1] 自动上标 | 正则分割 + `run.font.superscript=True` |
| 目录域（显示到二级） | `TOC \o "1-2"` 域代码；已注入 updateFields，打开文档自动更新 |
| 中英文字体分离 | `w:rFonts eastAsia=宋体 ascii=Times New Roman` |

### Step 4：DOCX 后处理与验证

1. **打开文档验证**：确认文件可正常打开，无损坏。
2. **域自动更新**：引擎已注入 `updateFields`，Word 打开文档时自动更新目录/页码域（无需手动 `Ctrl+A → F9`；若个别 Word 版本未弹窗刷新，可按 `F9` 手动补一次，但不作为交付要求）。
3. **对照自检清单**：逐项核对 [references/format-spec.md](format-spec.md) 末尾的"常见易错点清单"。
4. **抽查页面**：封面、中文摘要页、Abstract 页、目录页、正文第一页、含图表页、参考文献页。
5. **导出 PDF 验证**：确认分页正确，目录页码与正文一致。

---
