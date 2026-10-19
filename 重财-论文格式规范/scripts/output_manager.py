#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
重财-论文格式规范 · 输出版本管理器
==================================
解决历史 Bug：输出文件名写死（如「原文件名_重财规范版.docx」），
同一论文重复格式化会互相覆盖，且没有按项目隔离、没有版本记录。

设计原则
--------
1. 按项目隔离：每个论文（项目）一个专属目录，目录名取自论文题目，也可由用户显式指定。
2. 按版本迭代：每次运行自动递增 v1、v2、v3……，绝不复用旧文件名、绝不覆盖旧输出。
3. 一次输出 = 一个版本子目录：docx / tex / pdf / data.json 全部落在同一个 v<N>/ 内。
4. 可追溯：项目根目录自动维护 MANIFEST.md，记录每个版本的生成时间与文件清单。

输出目录树
----------
<输出根目录>/
└── <项目名slug>/                    # 项目目录（取自论文题目，或用户指定）
    ├── v1/                          # 第 1 次输出（自动递增，绝不覆盖）
    │   ├── word/                    # Word 成品
    │   │   └── <题目slug>_v1.docx
    │   ├── latex/                   # LaTeX 源码 + 编译产物
    │   │   ├── <题目slug>_v1.tex
    │   │   └── <题目slug>_v1.pdf    # 可选
    │   ├── data/                    # 结构化数据（可复用重新生成）
    │   │   └── <题目slug>_v1_data.json
    ├── v2/                          # 第 2 次输出
    │   └── ...
    └── MANIFEST.md                  # 版本清单（各版本时间 / 文件 / 说明）

依赖：标准库，无第三方依赖。

用法（代码）：
    from output_manager import OutputManager
    mgr = OutputManager(root_dir="output", project="基于Python的论文", version=None)  # None=自动递增
    mgr.create_dirs()
    docx_path = mgr.path_for("docx")
    tex_path  = mgr.path_for("tex")
    mgr.write_manifest(files=[docx_path, tex_path], notes="首次格式化")
    print(mgr.tree())

用法（命令行，常用于规划/预览）：
    # 预览下一次输出的版本号与文件名（不创建目录）
    python3 output_manager.py --title "基于Python的毕业论文格式自动化生成系统研究" --plan
    # 指定版本输出（覆盖校验：已存在的非空版本目录会报错，防止误覆盖）
    python3 output_manager.py --title "..." --version 2
    # 列出某项目的目录树
    python3 output_manager.py --root output --project "基于Python的论文" --list
"""

import os
import re
import json
import sys
import datetime
import argparse

# Windows / macOS 文件名非法字符：\ / : * ? " < > | 及控制字符
_ILLEGAL_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
_WHITESPACE = re.compile(r'\s+')
_VERSION_DIR = re.compile(r'^v(\d+)$')


def slugify(name, max_len=60, fallback="thesis"):
    """把论文题目 / 项目名转换为安全的目录名或文件名片段。

    - 保留中文、字母、数字、连字符、下划线（中文论文题目可原样入目录）
    - 非法字符与空白折叠为下划线
    - 超长截断；结果为空或非法时回退 fallback
    """
    if not name:
        return fallback
    s = str(name)
    s = _ILLEGAL_CHARS.sub('_', s)
    s = _WHITESPACE.sub('_', s)
    s = re.sub(r'_+', '_', s).strip('_')
    s = s[:max_len].rstrip('._-')
    return s or fallback


def _existing_versions(project_dir):
    """扫描项目目录下所有 v<N> 子目录，返回版本号列表（升序）。"""
    versions = []
    if os.path.isdir(project_dir):
        for entry in os.listdir(project_dir):
            m = _VERSION_DIR.fullmatch(entry)
            if m and os.path.isdir(os.path.join(project_dir, entry)):
                versions.append(int(m.group(1)))
    return sorted(versions)


def _walk_version_dir(vdir):
    """递归列出版本目录内的全部文件（相对路径，含类型子文件夹）。

    返回如 ['word/题目_v1.docx', 'latex/题目_v1.tex', 'data/题目_v1_data.json']。
    """
    files = []
    if not os.path.isdir(vdir):
        return files
    for root, dirs, names in os.walk(vdir):
        dirs.sort()
        for name in sorted(names):
            rel_dir = os.path.relpath(root, vdir)
            rel = name if rel_dir == "." else os.path.join(rel_dir, name)
            files.append(rel)
    return files


def resolve_next_version(project_dir, force=None):
    """计算下一次输出的版本号。

    - force=None：自动取「现有最大版本号 + 1」，保证绝不覆盖历史输出
    - force=N：强制使用 vN；若 vN 已存在且非空，抛 ValueError 阻止误覆盖
    """
    if force is not None:
        fv = int(force)
        if fv < 1:
            raise ValueError("版本号必须 >= 1")
        vdir = os.path.join(project_dir, "v%d" % fv)
        if os.path.isdir(vdir) and os.listdir(vdir):
            raise ValueError(
                "版本目录已存在且非空：%s（不会覆盖历史输出，请使用下一版本号或清理后重试）" % vdir)
        return fv
    versions = _existing_versions(project_dir)
    return (versions[-1] + 1) if versions else 1


class OutputManager:
    """论文输出目录与文件名管理。

    一次实例化 = 一次输出规划：确定 项目目录 + 版本号 v<N> + 全部产物文件名。
    """

    def __init__(self, root_dir=".", project=None, title="", version=None):
        self.root_dir = os.path.abspath(root_dir)
        # 项目名：优先用户显式指定；否则取论文题目；再回退通用名
        self.project = slugify(project or title or "thesis", max_len=40)
        self.title_slug = slugify(title or project or self.project, max_len=60)
        self.project_dir = os.path.join(self.root_dir, self.project)
        self.version = resolve_next_version(self.project_dir, force=version)
        self.version_dir = os.path.join(self.project_dir, "v%d" % self.version)
        # 所有产物共用的文件名主干：<题目slug>_v<N>
        self.stem = "%s_v%d" % (self.title_slug, self.version)

    # ----------------------------------------------------------
    #  目录
    # ----------------------------------------------------------
    # 版本目录内的类型子目录（分类存放产物，互不干扰）
    SUBDIRS = {"word": "word", "latex": "latex", "data": "data"}

    def create_dirs(self):
        """创建 项目目录/v<N>/{word,latex,data}/。已存在（空目录）则复用。"""
        os.makedirs(self.version_dir, exist_ok=True)
        for sub in self.SUBDIRS.values():
            os.makedirs(os.path.join(self.version_dir, sub), exist_ok=True)
        return self.version_dir

    # ----------------------------------------------------------
    #  文件名（v2.1.0：按类型分入子文件夹，文件名仍携带版本号）
    # ----------------------------------------------------------
    def path_for(self, kind):
        """返回某类型产物的绝对路径（v2.1.0 分文件夹）：

        - docx -> <版本目录>/word/<题目slug>_v<N>.docx
        - tex  -> <版本目录>/latex/<题目slug>_v<N>.tex
        - pdf  -> <版本目录>/latex/<题目slug>_v<N>.pdf
        - json -> <版本目录>/data/<题目slug>_v<N>_data.json
        """
        kind = kind.lower().lstrip('.')
        if kind == "json":
            return os.path.join(self.version_dir, "data", "%s_data.json" % self.stem)
        sub, ext = {"docx": ("word", "docx"),
                    "tex": ("latex", "tex"),
                    "pdf": ("latex", "pdf")}[kind]
        return os.path.join(self.version_dir, sub, "%s.%s" % (self.stem, ext))

    def data_path(self):
        """结构化论文数据文件：<版本目录>/data/<题目slug>_v<N>_data.json"""
        return self.path_for("json")

    def output_paths(self):
        """返回本次输出全部产物的路径字典。"""
        return {
            "docx": self.path_for("docx"),
            "tex": self.path_for("tex"),
            "pdf": self.path_for("pdf"),
            "data": self.data_path(),
        }

    # ----------------------------------------------------------
    #  MANIFEST.md 版本清单（自动由文件系统生成，可重复执行）
    # ----------------------------------------------------------
    def write_manifest(self, notes=""):
        """在项目根目录生成/刷新 MANIFEST.md，汇总所有版本的输出记录。

        采用「扫描目录 + 重写」而非增量追加，保证任意时刻清单与磁盘一致。
        notes 仅写入当前版本条目。
        """
        os.makedirs(self.project_dir, exist_ok=True)
        manifest_path = os.path.join(self.project_dir, "MANIFEST.md")
        rows = []
        for v in _existing_versions(self.project_dir):
            vdir = os.path.join(self.project_dir, "v%d" % v)
            files = _walk_version_dir(vdir)  # 相对路径列表（含子文件夹）
            mtime = datetime.datetime.fromtimestamp(
                os.path.getmtime(vdir)).strftime("%Y-%m-%d %H:%M")
            is_current = (v == self.version)
            note = notes if is_current else ""
            rows.append((v, mtime, files, note))

        lines = ["# 输出版本清单 · %s" % self.project, ""]
        lines.append("> 每次格式化输出独立版本目录 v<N>，历史版本一律保留、绝不覆盖。")
        lines.append("")
        for v, mtime, files, note in rows:
            lines.append("## v%d · %s" % (v, mtime))
            lines.append("")
            lines.append("- 输出目录：`%s/v%d/`" % (self.project, v))
            lines.append("- 文件：")
            for f in files:
                lines.append("  - `%s`" % f)
            if note:
                lines.append("- 说明：%s" % note)
            lines.append("")
        with open(manifest_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return manifest_path

    # ----------------------------------------------------------
    #  目录树
    # ----------------------------------------------------------
    def tree(self, include_planned=True):
        """返回项目目录树（v2.1.0 含类型子文件夹的嵌套结构）。

        include_planned=True：包含本次规划的版本目录（即使尚未创建，用于预览）；
        include_planned=False：只列磁盘上已存在的版本（用于 --list 查询历史）。
        """
        lines = [self.project_dir.rstrip("/") + "/"]
        versions = _existing_versions(self.project_dir)
        if include_planned and self.version not in versions:
            versions = sorted(versions + [self.version])
        for idx, v in enumerate(versions):
            is_last_v = (idx == len(versions) - 1)
            vdir = os.path.join(self.project_dir, "v%d" % v)
            entries = _walk_version_dir(vdir) if os.path.isdir(vdir) else []
            v_prefix = "└── " if is_last_v else "├── "
            lines.append("%sv%d/" % (v_prefix, v))
            for j, rel in enumerate(entries):
                is_last = (j == len(entries) - 1)
                pad = "    " if is_last_v else "│   "
                lines.append("%s%s %s" % (pad, "└──" if is_last else "├──", rel))
        if os.path.isfile(os.path.join(self.project_dir, "MANIFEST.md")):
            lines.append("└── MANIFEST.md")
        return "\n".join(lines)

    # ----------------------------------------------------------
    #  便捷：构建并保存结构化数据（配合 cfec_thesis_docx.py 使用）
    # ----------------------------------------------------------
    def save_thesis_data(self, thesis_data):
        """把本次输入的 thesis_data 保存为 JSON，便于复用重新生成。"""
        data_path = self.data_path()
        with open(data_path, "w", encoding="utf-8") as f:
            json.dump(thesis_data, f, ensure_ascii=False, indent=2)
        return data_path


# ============================================================
#  命令行入口
# ============================================================
def _cmd_plan(args):
    mgr = OutputManager(root_dir=args.root, project=args.project,
                        title=args.title, version=args.version)
    print("项目目录  : %s" % mgr.project_dir)
    print("本次版本  : v%d" % mgr.version)
    print("输出目录  : %s" % mgr.version_dir)
    print("产物文件  :")
    for key, path in mgr.output_paths().items():
        print("  %-5s -> %s" % (key, os.path.relpath(path, mgr.root_dir)))
    print("目录树    :")
    print(mgr.tree())


def _cmd_list(args):
    mgr = OutputManager(root_dir=args.root, project=args.project,
                        title=args.title, version=None)
    print(mgr.tree(include_planned=False))


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="重财-论文格式规范：输出版本管理器（规划/预览输出目录与文件名）")
    parser.add_argument("--root", default=".", help="输出根目录（默认当前目录）")
    parser.add_argument("--project", default="", help="项目名（默认取论文题目）")
    parser.add_argument("--title", default="", help="论文题目（决定项目目录与文件名）")
    parser.add_argument("--version", type=int, default=None,
                        help="指定版本号 vN（默认自动取下一版本）")
    parser.add_argument("--plan", action="store_true", help="预览下次输出的版本与文件名（不创建目录）")
    parser.add_argument("--list", action="store_true", help="列出某项目的目录树")
    args = parser.parse_args(argv)

    if args.list:
        _cmd_list(args)
    else:
        _cmd_plan(args)
    return 0


if __name__ == "__main__":
    sys.exit(main())
