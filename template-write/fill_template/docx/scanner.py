"""
# docx 模板扫描器 — 结构分析与样式提示。
- `scan_docx` 调用的内部工具：收集段落、识别标题、空节检测、打印结构与样式提示。
"""

from . import xmlutils

_BODY_STYLE_NAMES = frozenset({
    'normal', '正文', '默认段落字体', 'body text', 'bodytext',
    'caption', 'footnote', 'endnote', 'header', 'footer',
    'toc 1', 'toc 2', 'toc 3', 'tocheading',
    'table grid', 'list paragraph',
})


def collect_paragraphs(body, qn) -> tuple[list[tuple[str, str]], dict[str, int], dict[str, list[str]]]:
    """
    ## 收集段落信息、样式计数、每样式文本列表（用于长度/标题判断）。
    
    Args:
        body: docx.Document.body
        qn: docx.oxml.ns.qn 函数
        
    Returns:
        (all_paras, style_counts, style_texts)\
        (tuple[list[tuple[str, str]], dict[str, int], dict[str, list[str]]]):
        1. **all_paras** *(list[tuple[str, str]])*: **(style, text)** 段落列表，按文档顺序。
        2. **style_counts** *(dict[str, int])*: 样式名 -> 出现次数。
        3. **style_texts** *(dict[str, list[str]])*: 样式名 -> 段落文本列表。
    """
    all_paras: list[tuple[str, str]] = []  # (style, text)
    style_counts: dict[str, int] = {}
    style_texts: dict[str, list[str]] = {}

    for p_elem in body.iter(qn("w:p")):
        style = xmlutils.get_style_name(p_elem, qn)
        text = xmlutils.text_of(p_elem, qn).strip()
        all_paras.append((style, text))
        if not style or not text:
            continue

        style_counts[style] = style_counts.get(style, 0) + 1
        style_texts.setdefault(style, []).append(text)

    return all_paras, style_counts, style_texts


def identify_headings(all_paras: list[tuple[str, str]],
                      style_counts: dict[str, int],
                      style_texts: dict[str, list[str]]) -> tuple[list[tuple[str, str]], set[str]]:
    """
    ## 识别标题段落：Word 内置 Heading 或「短文本 + 高频」的自定义样式
    - 关键：正文样式（如 a8）的段落较长，标题样式（如 a4）的段落较短 ——
    用平均文本长度区分，避免把正文样式误判为标题。
    
    Args:
        all_paras (list[tuple[str, str]]): (style, text) 段落列表，按文档顺序。
        style_counts (dict[str, int]): 样式名 -> 出现次数。
        style_texts (dict[str, list[str]]): 样式名 -> 段落文本列表。
        
    Returns:
        (headings, heading_styles) (tuple[list[tuple[str, str]], set[str]]):
        1. **headings** *(list[tuple[str, str]])*: **(style, text)** 的标题段落列表，按文档顺序。
        2. **heading_styles** *(set[str])*: 被识别为标题的自定义样式名集合（不含内置 Heading 样式）。
    """
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

    return headings, heading_styles


def find_empty_sections(all_paras: list[tuple[str, str]],
                        headings: list[tuple[str, str]]) -> set[int]:
    """
    ## 检测哪些标题节为空（标题后无任何非空正文）。

    **用指针顺序匹配**：
    - 按文档顺序遍历段落;
    - 仅当段落文本等于下一个待匹配标题的文本时才推进指针;
    - 避免重复标题或正文恰好与标题同名时误判；
    - 标题之间的非空正文标记当前标题非空。
    
    Args:
        all_paras (list[tuple[str, str]]): (style, text) 段落列表，按文档顺序。
        headings (list[tuple[str, str]]): (style, text) 的标题段落列表，按文档顺序。
        
    Returns:
        empty_set (set[int]): 空节标题的索引集合（在 headings 中的索引）。
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


def print_structure(headings: list[tuple[str, str]], empty_set: set[int]) -> int:
    """
    ## 打印标题结构，标注空节，返回总标题数。
    
    Args:
        headings (list[tuple[str, str]]): (style, text) 的标题段落列表，按文档顺序。
        empty_set (set[int]): 空节标题的索引集合（在 headings 中的索引）。
    
    Returns:
        total (int): 标题总数。
    """
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

    return _count_total_headings(headings, style_to_depth, empty_set)


def print_style_hints(heading_styles: set[str],
                      style_counts: dict[str, int],
                      style_texts: dict[str, list[str]],
                      total: int) -> None:
    """
    ## 样式提示
    - 排除图片标题样式（非节边界），按出现次数排序取最频繁的图片标题样式
    （如 a3，段落多以"图"/"Figure"开头）不是节边界，不应作为 `--heading-style`
    
    Args:
        heading_styles (set[str]): 被识别为标题的自定义样式名集合（不含内置 Heading 样式）。
        style_counts (dict[str, int]): 样式名 -> 出现次数。
        style_texts (dict[str, list[str]]): 样式名 -> 段落文本列表。
        total (int): 标题总数。
    """
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


def _is_custom_heading_style(style: str,
                            style_counts: dict[str, int],
                            style_texts: dict[str, list[str]]) -> bool:
    """
    ## 判断自定义样式是否为标题样式（而非正文样式）。

    - **判据**：出现 >= 2 次、非已知正文样式名、且段落平均文本长度较短
    （标题短、正文长）。阈值 60 字符可区分 a4 标题与 a8 正文。
    
    Args:
        style (str): 样式名。
        style_counts (dict[str, int]): 样式名 -> 出现次数。
        style_texts (dict[str, list[str]]): 样式名 -> 段落文本列表。
    
    Returns:
        is_custom (bool): 如果是自定义标题样式则返回 True，否则返回 False。
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


def _count_total_headings(headings: list[tuple[str, str]], 
                         style_to_depth: dict[str, int],
                         empty_set: set[int]) -> int:
    """
    ## 统计标题数量
    
    Args:
        headings (list[tuple[str, str]]): (style, text) 的标题段落列表，按文档顺序。
        style_to_depth (dict[str, int]): 样式名 -> 层级深度（0 为最高级）。
        empty_set (set[int]): 空节标题的索引集合（在 headings 中的索引）。
        
    Returns:
        total (int): 标题总数。
    """
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
    return total


def _is_caption_style(style: str, style_texts: dict[str, list[str]]) -> bool:
    """
    ## 判断样式是否为图片标题样式（段落多以"图"/"Figure"/"Fig"开头）。
    - 图片标题不是节边界，不应作为 `--heading-style` 建议。
    
    Args:
        style (str): 样式名。
        style_texts (dict[str, list[str]]): 样式名 -> 段落文本列表。
        
    Returns:
        is_caption (bool): 如果是图片标题样式则返回 True，否则返回 False。
    """
    texts = style_texts.get(style, [])
    if not texts:
        return False
    
    caption_count = sum(1 for t in texts if t.startswith(("图", "Figure", "Fig")))
    return caption_count / len(texts) > 0.5
