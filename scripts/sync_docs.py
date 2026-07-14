#!/usr/bin/env python3
"""
# sync_docs.py — 从 opencode 分支生成 claude-code 分支的 CLAUDE.md / README.md / .gitignore

从 `refs/heads/opencode` 读取 AGENTS.md / README.md / .gitignore，做机械替换后写入当前分支
（应在 `claude-code` 分支上运行）。

替换表见 AGENTS.md "文档生成替换表"小节。新增涉及工具名的文本时，需同步更新
此脚本的替换表与 AGENTS.md 的表格。

**注意**：分支名 `opencode` / `claude-code` 不做替换——它们是 git 分支名，
两个分支共用同一套。仅在散文中指代工具名时替换。
"""
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent


def _read_opencode_file(rel_path: str) -> str:
    """从 refs/heads/opencode 读取文件内容（UTF-8）。"""
    result = subprocess.run(
        ["git", "show", f"refs/heads/opencode:{rel_path}"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        cwd=REPO_ROOT,
    )
    if result.returncode != 0:
        print(f"Error: cannot read {rel_path} from refs/heads/opencode:\n{result.stderr}",
              file=sys.stderr)
        sys.exit(1)
    return result.stdout


# --- 替换表 ---------------------------------------------------------------
# AGENTS.md -> CLAUDE.md：仅工具名，分支名不动
AGENTS_REPLACEMENTS = [
    ("opencode skills", "Claude Code skills"),
]

# README.md -> README.md：标题、工具名、路径、AI 指引文件名
README_REPLACEMENTS = [
    ("# OpenCode Skills",          "# Claude Skills"),
    ("OpenCode 技能集合",           "Claude Code 技能集合"),
    ("opencode 技能目录",            "Claude Code 技能目录"),
    ("~/.config/opencode/skills",  "~/.claude/skills"),
    ("├── AGENTS.md",              "├── CLAUDE.md"),
]

# .gitignore -> .gitignore：项目 AI 配置文件扩展名
GITIGNORE_REPLACEMENTS = [
    ("*.opencode",            "*.claude"),
    ("# opencode",            "# claude"),
]


def _apply(text: str, replacements: list[tuple[str, str]]) -> str:
    for old, new in replacements:
        text = text.replace(old, new)
    return text


def _write(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8", newline="\n")


def main() -> None:
    agents_src = _read_opencode_file("AGENTS.md")
    claude_out = _apply(agents_src, AGENTS_REPLACEMENTS)
    _write(REPO_ROOT / "CLAUDE.md", claude_out)
    print("Generated CLAUDE.md from AGENTS.md")

    readme_src = _read_opencode_file("README.md")
    readme_out = _apply(readme_src, README_REPLACEMENTS)
    _write(REPO_ROOT / "README.md", readme_out)
    print("Generated README.md from README.md (opencode)")

    gitignore_src = _read_opencode_file(".gitignore")
    gitignore_out = _apply(gitignore_src, GITIGNORE_REPLACEMENTS)
    _write(REPO_ROOT / ".gitignore", gitignore_out)
    print("Generated .gitignore from .gitignore (opencode)")


if __name__ == "__main__":
    main()
