"""
# docx 节级内容填充器

**核心原则**：文本匹配定位 + 样式名定义界。不推断标题级别。
**工具只做机械操作**：定位标题 -> 确定边界 -> 清除旧内容 -> 插入新内容 -> 验证。
智能判断由 Agent 通过 --scan 输出完成。

## 标题定位：
1. 全局定位："实验过程及分析" — 全文搜索第一个匹配段落
2. 限定定位："实验七 / 实验过程及分析" — 先找父标题，再在范围内找子标题

**边界检测只有两种策略**：
1. 相同样式名的段落 -> 同级标题 -> 节边界
2. Word 内置 Heading 样式的段落 -> 节边界
"""

from dataclasses import dataclass
from pathlib import Path
from sys import stderr
from typing import Any

from . import elements, locator, scanner
from ..deps import ensure_import


@dataclass
class _FillSession:
    Inches: Any
    body: Any
    doc: Any
    locate_ctx: locator.LocateContext


# ===================================================================
#  公开 API
# ===================================================================

def scan_docx(doc_path: Path) -> None:
    """
    ## 扫描 `.docx` 模板，输出标题结构和样式信息。
    - 输出每个标题段落的样式名和文本，标注空节，Claude 依据此输出决定 --heading-style 等参数。
    
    Args:
        doc_path (Path): `.docx` 文件路径。
    """
    Document = ensure_import("python-docx", "docx", attr="Document")
    qn = ensure_import("python-docx", "docx.oxml.ns", attr="qn")
    body = Document(str(doc_path)).element.body
    all_paras, style_counts, style_texts = scanner.collect_paragraphs(body, qn)
    headings, heading_styles = scanner.identify_headings(all_paras, style_counts, style_texts)
    empty_set = scanner.find_empty_sections(all_paras, headings)
    total = scanner.print_structure(headings, empty_set)
    scanner.print_style_hints(heading_styles, style_counts, style_texts, total)


def fill_docx_sections(doc_path: Path,
                       sections: dict[str, list[dict[str, Any]]],*,
                       mode: str = "replace",
                       heading_style: str | None = None,
                       dry_run: bool = False) -> int:
    """
    ## 按节标题填充 `.docx` 文档内容。

    Args:
        doc_path (Path): 已复制的 `.docx` 文件路径（原地修改）。
        sections (dict[str, list[dict[str, Any]]]): 标题文本 -> 内容项列表。支持 "父标题 / 子标题" 限定定位。
        mode (str): 'replace'（清空旧内容再填充）或 'append'（追加）。
        heading_style (str | None): 手动指定标题样式名（如 'a4'），用于边界检测。
        dry_run (bool): 仅检测定位和边界，不修改文件。

    Returns:
        int: 成功填充的节数量。
    """
    session = _open_fill_session(doc_path, heading_style)
    filled_count = 0
    missed: list[str] = []
    filled_summary: list[tuple[str, int, int]] = []

    for heading_text, items in sections.items():
        session.locate_ctx.children = list(session.body)

        loc = locator.locate_section(session.locate_ctx, heading_text)
        if loc is None:
            missed.append(heading_text)
            continue
        heading_idx, end_idx = loc

        if dry_run:
            end_desc = str(end_idx) if end_idx is not None else "end"
            pc = sum(1 for it in items if it.get("type") != "image")
            ic = sum(1 for it in items if it.get("type") == "image")
            print(
                f"  [DRY] '{heading_text}' -> idx={heading_idx}, end={end_desc}, "
                f"{pc} paragraphs + {ic} images"
            )
            filled_count += 1
            continue

        # 清除旧内容（replace 模式）—— 清除后索引移位，需重新定位标题作锚点
        if mode == "replace":
            elements.remove_elements_between(session.body, session.locate_ctx.children, heading_idx, end_idx)
            heading_idx = locator.find_heading_index_scoped(session.locate_ctx, heading_text)
            if heading_idx is None:
                missed.append(heading_text)
                continue

        # 插入新内容
        anchor = session.locate_ctx.children[heading_idx]
        para_count, img_count = elements.insert_items(session.doc, anchor, items, Inches=session.Inches)

        filled_count += 1
        filled_summary.append((heading_text, para_count, img_count))

    if dry_run:
        if missed:
            print(
                f"\nWarning: {len(missed)} heading(s) not found: "
                + ", ".join(f"'{t}'" for t in missed), file=stderr
            )
        print(f"\nDry-run: {filled_count}/{len(sections)} sections located, no changes made.")
        return filled_count

    session.doc.save(str(doc_path))
    _verify_filled(doc_path, filled_summary, sections)

    if missed:
        print(
            f"Warning: {len(missed)} heading(s) not found: "
            + ", ".join(f"'{t}'" for t in missed), file=stderr
        )
    return filled_count


# ===================================================================
#  会话与验证
# ===================================================================

def _open_fill_session(doc_path: Path, heading_style: str | None = None) -> _FillSession:
    """打开 docx 填充会话，返回 _FillSession 对象。"""
    Document = ensure_import("python-docx", "docx", attr="Document")
    qn = ensure_import("python-docx", "docx.oxml.ns", attr="qn")
    Inches = ensure_import("python-docx", "docx.shared", attr="Inches")
    doc = Document(str(doc_path))
    body = doc.element.body
    hs_lower = heading_style.lower() if heading_style else None
    locate_ctx = locator.LocateContext(children=list(body), qn=qn, hs_lower=hs_lower)
    return _FillSession(Inches=Inches, body=body, doc=doc, locate_ctx=locate_ctx)


def _verify_filled(doc_path: Path,
                   filled_summary: list[tuple[str, int, int]],
                   sections: dict[str, list[dict[str, Any]]]) -> None:
    """填充后轻量验证——重新打开文件检查已填充节的内容存在。"""
    if not filled_summary:
        return

    try:
        Document = ensure_import("python-docx", "docx", attr="Document")
    except ImportError as e:
        print(f'  [警告] 验证跳过: {e}', file=stderr)
        return
    doc = Document(str(doc_path))

    # 收集所有节标题的子文本（去重后），用于段落匹配
    all_child_texts: list[str] = []  # 所有子标题文本（小写），按顺序
    heading_keys: dict[str, list[str]] = {}  # child_text -> [orig_key1, orig_key2, ...]
    for ht in sections:
        _, sep, child = ht.partition(locator.SCOPE_SEP)
        key = child.strip() if sep else ht.strip()
        heading_keys.setdefault(key.lower(), []).append(ht)
        all_child_texts.append(key.lower())

    # 用指针方式遍历：每个标题对应一个指针，遇到匹配段落时推进
    # 按文档顺序，将每个段落分配给当前活跃的节
    section_texts: dict[str, list[str]] = {ht: [] for ht in sections}
    active_section: str | None = None

    # 构建 ordered_keys 列表
    ordered_keys: list[str] = []
    for ht in sections:
        _, sep, child = ht.partition(locator.SCOPE_SEP)
        key = child.strip() if sep else ht.strip()
        ordered_keys.append(ht)
    ptr = 0  # 指向 ordered_keys 中下一个待匹配的标题

    for p in doc.paragraphs:
        p_text = p.text.strip()
        if not p_text:
            continue
        
        # 检查是否匹配某个待匹配的标题
        matched = False
        if ptr < len(ordered_keys):
            _, sep, child = ordered_keys[ptr].partition(locator.SCOPE_SEP)
            key = child.strip() if sep else ordered_keys[ptr].strip()
            if key.lower() in p_text.lower():
                active_section = ordered_keys[ptr]
                ptr += 1
                matched = True
        
        # 正文内容归入当前活跃节
        if not matched and active_section and active_section in section_texts:
            section_texts[active_section].append(p_text[:80])

    print("Verification:")
    for heading_text, para_count, img_count in filled_summary:
        texts = section_texts.get(heading_text, [])
        status = "OK" if texts else "EMPTY"
        detail = (
            f"{para_count} paragraphs + {img_count} images"
            if status == "OK" else "no content found"
        )
        print(f"  [{status}] '{heading_text}': {detail}")
