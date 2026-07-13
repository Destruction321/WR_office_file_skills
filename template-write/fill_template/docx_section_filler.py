"""
# docx 节级内容填充器

**核心原则**：文本匹配定位 + 样式名定义界。不推断标题级别。
**工具只做机械操作**：定位标题 -> 确定边界 -> 清除旧内容 -> 插入新内容 -> 验证。
智能判断由 Claude 通过 --scan 输出完成。

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

from .deps import ensure_import
from .docx_scanner import (
    collect_paragraphs, identify_headings, find_empty_sections,
    print_structure, print_style_hints,
    get_style_name, text_of,
)

_SCOPE_SEP = ' / '

@dataclass
class _LocateContext:
    children: list[Any]
    qn: Any
    hs_lower: str | None = None

@dataclass
class _FillSession:
    Inches: Any
    body: Any
    doc: Any
    locate_ctx: _LocateContext


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
    all_paras, style_counts, style_texts = collect_paragraphs(body, qn)
    headings, heading_styles = identify_headings(all_paras, style_counts, style_texts)
    empty_set = find_empty_sections(all_paras, headings)
    total = print_structure(headings, empty_set)
    print_style_hints(heading_styles, style_counts, style_texts, total)


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

        loc = _locate_section(session.locate_ctx, heading_text)
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
            _remove_elements_between(session.body, session.locate_ctx.children, heading_idx, end_idx)
            heading_idx = _find_heading_index_scoped(session.locate_ctx, heading_text)
            if heading_idx is None:
                missed.append(heading_text)
                continue

        # 插入新内容
        anchor = session.locate_ctx.children[heading_idx]
        para_count, img_count = _insert_items(session.doc, anchor, items, Inches=session.Inches)

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
#  定位
# ===================================================================

def _locate_section(locate_ctx: _LocateContext, heading_text: str) -> tuple[int, int | None] | None:
    """
    定位 (heading_idx, end_idx)：先查标题索引，再以该标题样式算节边界。

    全局定位："标题文本"；限定定位："父标题 / 子标题"。
    标题索引查找共用 `_find_heading_index_scoped`，与 replace 后重定位一致。
    """
    heading_idx = _find_heading_index_scoped(locate_ctx, heading_text)
    if heading_idx is None:
        return None

    end_idx = _find_style_boundary(
        locate_ctx,
        heading_idx,
        get_style_name(locate_ctx.children[heading_idx], locate_ctx.qn)
    )
    return heading_idx, end_idx


def _find_heading_index_scoped(locate_ctx: _LocateContext, heading_text: str) -> int | None:
    """按标题文本查找段落索引，支持限定定位语法。

    - 全局："标题文本" -> 全文找第一个匹配段落。
    - 限定："父标题 / 子标题" -> 先找父标题，再在其后找子标题
      （不限制 parent_end，因父子可能同样式，样式边界无法区分）。

    Returns:
        int | None: 标题段落在 children 中的索引，未找到返回 None。
    """
    parent_text, sep, child_text = heading_text.partition(_SCOPE_SEP)
    if not sep:
        return _find_heading_index(locate_ctx, parent_text)

    parent_idx = _find_heading_index(locate_ctx, parent_text)
    if parent_idx is None:
        return None
    return _find_heading_index(locate_ctx, child_text.strip(), start=parent_idx + 1)


def _find_heading_index(locate_ctx: _LocateContext, text: str, start: int = 0) -> int | None:
    """在 children[start:len(children)] 范围内找到包含 text 的段落，返回索引。大小写不敏感。"""
    text_lower = text.strip().lower()
    end = len(locate_ctx.children)
    for i in range(start, end):
        child = locate_ctx.children[i]
        if child.tag != locate_ctx.qn("w:p"):
            continue
        
        if text_lower in text_of(child, locate_ctx.qn).lower():
            return i
    
    return None


def _find_style_boundary(locate_ctx: _LocateContext,
                         heading_idx: int,
                         heading_style: str) -> int | None:
    """找到 heading_idx 之后第一个同级标题段落，返回其索引。

    边界判定：
    - 与目标标题相同样式名的段落 -> 边界
    - 用户指定 --heading-style 的段落 -> 边界
    - Word 内置 Heading 样式的段落 -> 边界（当原标题不是内置样式时）
    """
    search_end = len(locate_ctx.children)
    hs_lower_local = heading_style.lower() if heading_style else ""
    is_builtin = hs_lower_local.startswith("heading")

    for i in range(heading_idx + 1, search_end):
        child = locate_ctx.children[i]
        if child.tag != locate_ctx.qn("w:p"):
            continue

        style_lower = get_style_name(child, locate_ctx.qn).lower()

        # 相同样式名 -> 同级标题 -> 边界
        if hs_lower_local and style_lower == hs_lower_local:
            return i
        
        # 用户指定的 heading_style -> 边界
        if locate_ctx.hs_lower and style_lower == locate_ctx.hs_lower:
            return i
        
        # Word 内置 Heading -> 边界（当原样式不是内置样式时）
        if not is_builtin and style_lower.startswith("heading"):
            return i

    return None


# ===================================================================
#  元素操作
# ===================================================================

def _remove_elements_between(body, children, start: int, end: int | None):
    """删除 body 中 (start, end) 之间的所有子元素。"""
    actual_end = end if end is not None else len(children)
    for i in range(actual_end - 1, start, -1):
        body.remove(children[i])


def _insert_items(doc, anchor, items: list[dict[str, Any]], *, Inches) -> tuple[int, int]:
    """
    将内容项逐个构建为段落，用 addnext 链式插入到 anchor 之后。

    - image 项走 `_build_image_para`，其余走 `_build_text_para`。
    - 每次插入后把 anchor 前移到新段落，保证顺序与 items 一致。
    """
    para_count = 0
    img_count = 0
    for item in items:
        if item.get("type") == "image":
            p = _build_image_para(doc, item, Inches)
            img_count += 1
        
        else:
            p = _build_text_para(doc, item)
            para_count += 1
        anchor.addnext(p)
        anchor = p
    
    return para_count, img_count


def _build_image_para(doc, item: dict[str, Any], Inches):
    """构建内嵌图片段落。"""
    image_path = Path(item["path"])
    if not image_path.exists():
        p = doc.add_paragraph()
        run = p.add_run()
        run.text = f"[Image not found: {image_path.name}]"
        run.bold = True
        p._element.getparent().remove(p._element)
        return p._element

    width = item.get("width_inches", 5.5)
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_picture(str(image_path), width=Inches(width))
    p._element.getparent().remove(p._element)
    return p._element


def _build_text_para(doc, item: dict[str, Any]):
    """构建文本段落。支持 runs 格式（行内加粗/斜体）和旧 text/bold 格式。"""
    p = doc.add_paragraph()

    if "runs" in item:
        for run_spec in item["runs"]:
            run = p.add_run()
            run.text = run_spec.get("text", "")
            if run_spec.get("bold"):
                run.bold = True
            if run_spec.get("italic"):
                run.italic = True
    
    else:
        run = p.add_run()
        run.text = item.get("text", "")
        if item.get("bold"):
            run.bold = True

    p._element.getparent().remove(p._element)
    return p._element


def _open_fill_session(doc_path: Path, heading_style: str | None = None) -> _FillSession:
    """打开 docx 填充会话，返回 _FillSession 对象。"""
    Document = ensure_import("python-docx", "docx", attr="Document")
    qn = ensure_import("python-docx", "docx.oxml.ns", attr="qn")
    Inches = ensure_import("python-docx", "docx.shared", attr="Inches")
    doc = Document(str(doc_path))
    body = doc.element.body
    hs_lower = heading_style.lower() if heading_style else None
    locate_ctx = _LocateContext(children=list(body), qn=qn, hs_lower=hs_lower)
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
        _, sep, child = ht.partition(_SCOPE_SEP)
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
        _, sep, child = ht.partition(_SCOPE_SEP)
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
            _, sep, child = ordered_keys[ptr].partition(_SCOPE_SEP)
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
