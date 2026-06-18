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

from pathlib import Path
from sys import stderr
from typing import Any

from .deps import ensure_import

_SCOPE_SEP = ' / '

_BODY_STYLE_NAMES = frozenset({
    'normal', '正文', '默认段落字体', 'body text', 'bodytext',
    'caption', 'footnote', 'endnote', 'header', 'footer',
    'toc 1', 'toc 2', 'toc 3', 'tocheading',
    'table grid', 'list paragraph',
})


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
    mod = ensure_import("python-docx", "docx")
    Document = mod.Document
    from docx.oxml.ns import qn

    doc = Document(str(doc_path))
    body = doc.element.body

    # 收集段落信息 & 样式计数 & 每样式文本列表（用于长度/标题判断）
    all_paras: list[tuple[str, str]] = []  # (style, text)
    style_counts: dict[str, int] = {}
    style_texts: dict[str, list[str]] = {}
    for p_elem in body.iter(qn("w:p")):
        style = _get_style_name(p_elem)
        text = _text_of(p_elem).strip()
        all_paras.append((style, text))
        if style and text:
            style_counts[style] = style_counts.get(style, 0) + 1
            style_texts.setdefault(style, []).append(text)

    # 识别标题段落：Word 内置 Heading 或「短文本 + 高频」的自定义样式
    # 关键：正文样式（如 a8）的段落较长，标题样式（如 a4）的段落较短——
    # 用平均文本长度区分，避免把正文样式误判为标题。
    headings: list[tuple[str, str]] = []  # (style, text)
    heading_styles: set[str] = set()
    for style, text in all_paras:
        if not text:
            continue
        
        if style and style.lower().startswith("heading"):
            headings.append((style, text))
        
        elif style and _is_custom_heading_style(style, style_counts, style_texts):
            headings.append((style, text))
            heading_styles.add(style)

    # 空节检测：标题后无任何非空正文即为空节
    empty_set = _find_empty_sections(all_paras, headings)

    # 输出：按样式推断缩进层级
    print("Document structure:")
    style_to_depth: dict[str, int] = {}
    if headings:
        for style, _ in headings:
            if style in style_to_depth:
                continue

            if style.lower().startswith("heading"):
                try:
                    d = int(style.split()[-1]) - 1
                except ValueError:
                    d = len(style_to_depth)
            else:
                d = len(style_to_depth)
            
            style_to_depth[style] = d

        min_depth = min(style_to_depth.values())
        for hi, (style, text) in enumerate(headings):
            d = style_to_depth.get(style, 0) - min_depth
            indent = "  " * max(d, 0)
            empty_tag = " [EMPTY]" if hi in empty_set and d > 0 else ""
            print(f"  {indent}[{style}] {text}{empty_tag}")

    # 统计
    total = len(headings)
    if headings and style_to_depth:
        min_d = min(style_to_depth.values())
        empty_count = sum(
            1 for hi in empty_set
            if style_to_depth.get(headings[hi][0], 0) > min_d
        )
    
    else:
        empty_count = 0
    print(f"\nHeadings: {total} total, {empty_count} empty sub-sections")

    # 样式提示——排除图片标题样式（非节边界），按出现次数排序取最频繁的
    # 图片标题样式（如 a3，段落多以"图"/"Figure"开头）不是节边界，不应作为 --heading-style
    boundary_styles = [s for s in heading_styles if not _is_caption_style(s, style_texts)]
    if boundary_styles:
        ranked = sorted(boundary_styles, key=lambda s: style_counts.get(s, 0), reverse=True)
        styles_str = ", ".join(sorted(heading_styles))
        primary = ranked[0]
        print(f"Custom heading styles: {styles_str}")
        print(f"Hint: use --heading-style {primary} for section filling.")
    
    elif heading_styles:
        # 只有图片标题样式时，回退到全部标题样式
        ranked = sorted(heading_styles, key=lambda s: style_counts.get(s, 0), reverse=True)
        print(f"Custom heading styles: {', '.join(ranked)}")
        print(f"Hint: use --heading-style {ranked[0]} for section filling.")
    
    else:
        print("No custom heading styles detected. Document uses standard Heading styles.")
    
    if total > 0:
        print('Hint: use scoped syntax "Parent / Child" to disambiguate duplicate headings.')


def fill_docx_sections(doc_path: Path,
                       sections: dict[str, list[dict[str, Any]]], *,
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
    mod = ensure_import("python-docx", "docx")
    Document = mod.Document

    doc = Document(str(doc_path))
    body = doc.element.body
    hs_lower = heading_style.lower() if heading_style else None

    filled_count = 0
    missed: list[str] = []
    filled_summary: list[tuple[str, int, int]] = []

    for heading_text, items in sections.items():
        children = list(body)

        loc = _locate_section(children, heading_text, hs_lower=hs_lower)
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
            _remove_elements_between(body, children, heading_idx, end_idx)
            children = list(body)
            heading_idx = _find_heading_index_scoped(children, heading_text)
            if heading_idx is None:
                missed.append(heading_text)
                continue

        # 插入新内容
        anchor = children[heading_idx]
        para_count, img_count = _insert_items(doc, anchor, items)

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

    doc.save(str(doc_path))
    _verify_filled(doc_path, filled_summary, sections)

    if missed:
        print(
            f"Warning: {len(missed)} heading(s) not found: "
            + ", ".join(f"'{t}'" for t in missed), file=stderr
        )
    return filled_count


def _verify_filled(doc_path: Path,
                   filled_summary: list[tuple[str, int, int]],
                   sections: dict[str, list[dict[str, Any]]]) -> None:
    """填充后轻量验证——重新打开文件检查已填充节的内容存在。"""
    if not filled_summary:
        return

    mod = ensure_import("python-docx", "docx")
    Document = mod.Document
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


# ===================================================================
#  定位
# ===================================================================

def _locate_section(children: list,
                    heading_text: str, *,
                    hs_lower: str | None = None) -> tuple[int, int | None] | None:
    """定位 (heading_idx, end_idx)：先查标题索引，再以该标题样式算节边界。

    全局定位："标题文本"；限定定位："父标题 / 子标题"。
    标题索引查找共用 `_find_heading_index_scoped`，与 replace 后重定位一致。
    """
    heading_idx = _find_heading_index_scoped(children, heading_text)
    if heading_idx is None:
        return None

    end_idx = _find_style_boundary(
        children, heading_idx, _get_style_name(children[heading_idx]),
        hs_lower=hs_lower,
    )
    return heading_idx, end_idx


def _find_heading_index(children: list, text: str, *, start: int = 0, end: int | None = None) -> int | None:
    """在 children[start:end] 范围内找到包含 text 的段落，返回索引。大小写不敏感。"""
    from docx.oxml.ns import qn

    text_lower = text.strip().lower()
    search_end = end if end is not None else len(children)
    for i in range(start, search_end):
        child = children[i]
        if child.tag != qn("w:p"):
            continue
        
        if text_lower in _text_of(child).lower():
            return i
    
    return None


def _find_heading_index_scoped(children: list, heading_text: str) -> int | None:
    """按标题文本查找段落索引，支持限定定位语法。

    - 全局："标题文本" -> 全文找第一个匹配段落。
    - 限定："父标题 / 子标题" -> 先找父标题，再在其后找子标题
      （不限制 parent_end，因父子可能同样式，样式边界无法区分）。

    Returns:
        int | None: 标题段落在 children 中的索引，未找到返回 None。
    """
    parent_text, sep, child_text = heading_text.partition(_SCOPE_SEP)
    if not sep:
        return _find_heading_index(children, parent_text)

    parent_idx = _find_heading_index(children, parent_text)
    if parent_idx is None:
        return None
    return _find_heading_index(children, child_text.strip(), start=parent_idx + 1)


def _find_style_boundary(children: list,
                         heading_idx: int,
                         heading_style: str, *,
                         hs_lower: str | None = None,
                         hard_end: int | None = None) -> int | None:
    """找到 heading_idx 之后第一个同级标题段落，返回其索引。

    边界判定：
    - 与目标标题相同样式名的段落 -> 边界
    - 用户指定 --heading-style 的段落 -> 边界
    - Word 内置 Heading 样式的段落 -> 边界（当原标题不是内置样式时）
    - hard_end 限制搜索范围（限定定位时不超过父节边界）
    """
    from docx.oxml.ns import qn

    search_end = hard_end if hard_end is not None else len(children)
    hs_lower_local = heading_style.lower() if heading_style else ""
    is_builtin = hs_lower_local.startswith("heading")

    for i in range(heading_idx + 1, search_end):
        child = children[i]
        if child.tag != qn("w:p"):
            continue

        style_lower = _get_style_name(child).lower()

        # 相同样式名 -> 同级标题 -> 边界
        if hs_lower_local and style_lower == hs_lower_local:
            return i
        
        # 用户指定的 heading_style -> 边界
        if hs_lower and style_lower == hs_lower:
            return i
        
        # Word 内置 Heading -> 边界（当原样式不是内置样式时）
        if not is_builtin and style_lower.startswith("heading"):
            return i

    return None


# ===================================================================
#  工具函数
# ===================================================================

def _text_of(p_elem) -> str:
    """收集 w:p 元素内所有 w:t 文本。"""
    from docx.oxml.ns import qn
    return "".join(t.text or "" for t in p_elem.iter(qn("w:t")))


def _find_empty_sections(all_paras: list[tuple[str, str]],
                         headings: list[tuple[str, str]]) -> set[int]:
    """
    检测哪些标题节为空（标题后无任何非空正文）。

    **用指针顺序匹配**：按文档顺序遍历段落,
    仅当段落文本等于下一个待匹配标题的文本时才推进指针,
    避免重复标题或正文恰好与标题同名时误判;
    标题之间的非空正文标记当前标题非空。
    """
    empty_set = set(range(len(headings)))
    ptr = 0       # 指向下一个待匹配的标题
    current = -1  # 当前所在标题索引

    for _, text in all_paras:
        if not text:
            continue
        
        # 段落文本等于下一个待匹配标题 -> 进入该标题
        if ptr < len(headings) and text == headings[ptr][1]:
            current = ptr
            ptr += 1
            continue
        
        # 非标题正文 -> 标记当前标题非空
        if current >= 0:
            empty_set.discard(current)

    return empty_set


def _is_custom_heading_style(style: str,
                             style_counts: dict[str, int],
                             style_texts: dict[str, list[str]]) -> bool:
    """判断自定义样式是否为标题样式（而非正文样式）。

    判据：出现 ≥2 次、非已知正文样式名、且段落平均文本长度较短
    （标题短、正文长）。阈值 60 字符可区分 a4 标题与 a8 正文。
    """
    if style.lower() in _BODY_STYLE_NAMES:
        return False
    if style_counts.get(style, 0) < 2:
        return False
    
    texts = style_texts.get(style, [])
    if not texts:
        return False
    
    avg_len = sum(len(t) for t in texts) / len(texts)
    return avg_len <= 60


def _is_caption_style(style: str, style_texts: dict[str, list[str]]) -> bool:
    """判断样式是否为图片标题样式（段落多以"图"/"Figure"/"Fig"开头）。

    图片标题不是节边界，不应作为 --heading-style 建议。
    """
    texts = style_texts.get(style, [])
    if not texts:
        return False
    
    caption_count = sum(1 for t in texts if t.startswith(("图", "Figure", "Fig")))
    return caption_count / len(texts) > 0.5


def _get_style_name(p_elem) -> str:
    """提取段落的 w:pStyle 值，无样式时返回空字符串。"""
    from docx.oxml.ns import qn
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return ""
    
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        return ""
    
    return pStyle.get(qn("w:val"), "")


# ===================================================================
#  元素操作
# ===================================================================

def _remove_elements_between(body, children, start: int, end: int | None):
    """删除 body 中 (start, end) 之间的所有子元素。"""
    actual_end = end if end is not None else len(children)
    for i in range(actual_end - 1, start, -1):
        body.remove(children[i])


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


def _build_image_para(doc, item: dict[str, Any]):
    """构建内嵌图片段落。"""
    from docx.shared import Inches

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


def _insert_items(doc, anchor, items: list[dict[str, Any]]) -> tuple[int, int]:
    """
    将内容项逐个构建为段落，用 addnext 链式插入到 anchor 之后。

    - image 项走 `_build_image_para`，其余走 `_build_text_para`。
    - 每次插入后把 anchor 前移到新段落，保证顺序与 items 一致。
    """
    para_count = 0
    img_count = 0
    for item in items:
        if item.get("type") == "image":
            p = _build_image_para(doc, item)
            img_count += 1
        
        else:
            p = _build_text_para(doc, item)
            para_count += 1
        anchor.addnext(p)
        anchor = p
    
    return para_count, img_count
