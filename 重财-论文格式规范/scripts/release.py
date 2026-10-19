#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
重财-论文格式规范 · Skill 发布辅助脚本（release.py）
====================================================

用途：当用户直接要求 agent「更新 / 升级这个 skill（脚本）」时，
自动定位 GitHub 仓库与当前版本，辅助完成一次标准发布：

    改代码 → 版本递增（SKILL.md / manifest.json / CHANGELOG 三处同步）
         → 本地验证 → 同步本地 skill 与仓库 clone → 提交推送（白澤 程 署名）→ 打 tag

自动定位（本脚本不依赖额外配置文件，全部按约定解析）：
- GitHub 仓库  : https://github.com/wpc6666/cfec-thesis-formatter
                 （发布以 main 分支 + 语义化 tag vX.Y.Z 记录）
- 本地 skill  : 本脚本所在目录的上一级（<skill_root>/），
                 即与 SKILL.md、manifest.json 同仓
- 当前版本    : manifest.json 的 version 字段（单一权威来源，
                 须与 SKILL.md frontmatter、CHANGELOG 最新条目一致）

版本递增规则（语义化版本 X.Y.Z，见 https://semver.org/lang/zh-CN/）：
- patch（X.Y.Z → X.Y.Z+1）：修复 Bug，不影响外部行为
- minor（X.Y.Z → X.Y+1.0）：新增向后兼容的功能
- major（X.Y.Z → X+1.0.0）：破坏性变更

用法：
    python3 scripts/release.py --check                 # 定位：仓库、当前版本、最近 tag、本地状态
    python3 scripts/release.py --bump minor            # 递增版本并同步三处文件
    python3 scripts/release.py --bump patch --dry-run  # 只预览将做的改动，不写文件
"""

import io
import os
import re
import sys
import json
import datetime
import subprocess
import argparse

_PATH_MAP = {}  # 文件名 → 真实路径（index.html 在仓库根，不在 SKILL_DIR）

REPO_URL = "https://github.com/wpc6666/cfec-thesis-formatter.git"
AUTHOR_NAME = "白澤 程"
AUTHOR_EMAIL = "87405429+wpc6666@users.noreply.github.com"

# 本文件位于 <skill_root>/scripts/release.py → skill 根目录为上一级
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MANIFEST_JSON = os.path.join(SKILL_DIR, "manifest.json")
SKILL_MD = os.path.join(SKILL_DIR, "SKILL.md")
CHANGELOG_MD = os.path.join(SKILL_DIR, "CHANGELOG.md")

_VERSION_RE = re.compile(r"^\d+\.\d+\.\d+$")


def _err(msg):
    sys.stderr.write("错误: %s\n" % msg)
    sys.exit(1)


# ----------------------------------------------------------
#  版本读写
# ----------------------------------------------------------
def read_current_version():
    """读取 manifest.json 的 version（单一权威来源）。"""
    if not os.path.isfile(MANIFEST_JSON):
        _err("找不到 manifest.json（期望位于 %s）" % MANIFEST_JSON)
    with open(MANIFEST_JSON, encoding="utf-8") as f:
        data = json.load(f)
    version = data.get("version", "")
    if not _VERSION_RE.fullmatch(version):
        _err("manifest.json 中 version 字段异常: %r（期望 X.Y.Z）" % version)
    return version


def _bump_version(current, part):
    major, minor, patch = (int(x) for x in current.split("."))
    if part == "major":
        return "%d.0.0" % (major + 1)
    if part == "minor":
        return "%d.%d.0" % (major, minor + 1)
    if part == "patch":
        return "%d.%d.%d" % (major, minor, patch + 1)
    _err("未知的递增类型: %r（可选 major/minor/patch）" % part)


# ----------------------------------------------------------
#  三处文件同步（SKILL.md / manifest.json / CHANGELOG.md）
# ----------------------------------------------------------
def _patch_skill_md(new_version):
    """更新 SKILL.md frontmatter 的 version 与一级标题版本号。"""
    with open(SKILL_MD, encoding="utf-8") as f:
        text = f.read()
    before = text
    # 1) frontmatter: version: X.Y.Z
    text = re.sub(r"(?m)^version:\s*\d+\.\d+\.\d+\s*$",
                  "version: %s" % new_version, text, count=1)
    # 2) 一级标题: # 重财-论文格式规范 vX.Y.Z
    text = re.sub(r"^(#\s*重财-论文格式规范\s*)v\d+\.\d+\.\d+\s*$",
                  r"\1v%s" % new_version, text, count=1, flags=re.M)
    return before, text


def _patch_manifest(new_version):
    with open(MANIFEST_JSON, encoding="utf-8") as f:
        data = json.load(f)
    old = data.get("version")
    data["version"] = new_version
    return old, json.dumps(data, ensure_ascii=False, indent=2) + "\n"


def _patch_index_html(new_version):
    """把 index.html 里的版本号同步。

    index.html 在仓库根、不在 skill 子目录里，所以往上找一级。
    返回 (name, before, after)；找不到文件返回 None。真实路径记进 _PATH_MAP，
    因为写入循环按文件名查路径，而 index.html 不在 SKILL_DIR 下。
    """
    for cand in (os.path.join(os.path.dirname(SKILL_DIR), "index.html"),
                 os.path.join(SKILL_DIR, "index.html")):
        if not os.path.isfile(cand):
            continue
        text = io.open(cand, encoding="utf-8").read()
        new_text, n = re.subn(r"v\d+\.\d+\.\d+", "v" + new_version, text)
        if n == 0:
            return None
        _PATH_MAP["index.html"] = cand
        return ("index.html", text, new_text)
    return None


def _patch_changelog(new_version):
    """在 CHANGELOG 第一个版本条目（## [x.y.z]）之前插入新条目。"""
    with open(CHANGELOG_MD, encoding="utf-8") as f:
        text = f.read()
    today = datetime.date.today().strftime("%Y-%m-%d")
    entry = (
        "## [%s] - %s\n\n"
        "### 新增 / 修复\n\n"
        "- （待补充：本次改动内容）\n\n"
        "---\n\n" % (new_version, today)
    )
    m = re.search(r"(?m)^##\s*\[", text)
    if m:
        return text, text[:m.start()] + entry + text[m.start():]
    return text, text.rstrip("\n") + "\n\n" + entry


def apply_bump(part, dry_run=False):
    current = read_current_version()
    new = _bump_version(current, part)
    print("当前版本  : v%s" % current)
    print("新版本    : v%s（%s 递增）" % (new, part))
    print("仓库      : %s" % REPO_URL)
    print("本地 skill: %s" % SKILL_DIR)
    if dry_run:
        print("\n[dry-run] 将修改以下文件（未写入）：")
    else:
        print("\n写入中：")

    targets = [
        ("SKILL.md", *_patch_skill_md(new)),
        ("manifest.json", *_patch_manifest(new)),
    ]
    _, changelog_before, changelog_after = ("CHANGELOG.md", *_patch_changelog(new))
    targets.append(("CHANGELOG.md", changelog_before, changelog_after))

    # index.html 也带版本号，不一起改就会长期漂在旧版本（v2.2.0 的现场）
    idx = _patch_index_html(new)
    if idx:
        targets.append(idx)

    for name, before, after in targets:
        changed = (before != after)
        print("  %-14s %s" % (name, "更新 → 版本 %s" % new if changed else "无变化"))
        if changed and not dry_run:
            path = _PATH_MAP.get(name) or {
                "SKILL.md": SKILL_MD, "manifest.json": MANIFEST_JSON,
                "CHANGELOG.md": CHANGELOG_MD}[name]
            with open(path, "w", encoding="utf-8") as f:
                f.write(after)
    if dry_run:
        print("\n提示: 去掉 --dry-run 后真正写入；写入后请补充 CHANGELOG 新条目内容。")
    return new


# ----------------------------------------------------------
#  --check：自动定位仓库与版本
# ----------------------------------------------------------
def run_check():
    version = read_current_version()
    print("== 自动定位结果 ==")
    print("GitHub 仓库 : %s" % REPO_URL)
    print("本地 skill  : %s" % SKILL_DIR)
    print("当前版本    : v%s（manifest.json 单一权威来源）" % version)
    print("SKILL.md    : %s" % os.path.join(SKILL_DIR, "SKILL.md"))
    print("CHANGELOG   : %s" % os.path.join(SKILL_DIR, "CHANGELOG.md"))
    # 最近 tag（若本地存在 git 仓库）
    git_root = os.path.join(SKILL_DIR, ".git")
    if os.path.isdir(git_root):
        try:
            tag = subprocess.check_output(
                ["git", "tag", "--sort=-version:refname"],
                cwd=SKILL_DIR, stderr=subprocess.DEVNULL).decode().strip().splitlines()
            print("最近 tag    : %s" % (tag[0] if tag else "（无 tag）"))
            head = subprocess.check_output(
                ["git", "log", "-1", "--format=%h %s"],
                cwd=SKILL_DIR, stderr=subprocess.DEVNULL).decode().strip()
            print("HEAD        : %s" % head)
        except Exception:
            print("最近 tag    : （读取失败，忽略）")
    else:
        print("最近 tag    : （本地 skill 目录无 .git；发布时以仓库 clone 为准）")


# ----------------------------------------------------------
#  入口
# ----------------------------------------------------------
def main(argv=None):
    parser = argparse.ArgumentParser(
        description="重财-论文格式规范：Skill 发布辅助（自动定位 GitHub 仓库与版本）")
    parser.add_argument("--check", action="store_true",
                        help="定位并显示仓库、当前版本、最近 tag、本地状态")
    parser.add_argument("--bump", choices=["major", "minor", "patch"],
                        help="递增版本（major/minor/patch）并同步 SKILL.md / manifest.json / CHANGELOG")
    parser.add_argument("--dry-run", action="store_true",
                        help="与 --bump 联用：只预览改动，不写文件")
    args = parser.parse_args(argv)

    if args.check:
        run_check()
    elif args.bump:
        apply_bump(args.bump, dry_run=args.dry_run)
    else:
        parser.print_help()
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
