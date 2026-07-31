"""
# docx 模板扫描器 — 结构分析与样式提示。
- `scan_docx` 调用的内部工具：收集段落、识别标题、空节检测、打印结构与样式提示。
"""

from dataclasses import dataclass

from . import xmlutils

_BODY_STYLE_NAMES = frozenset({
    'normal', '正文', '默认段落字体', 'body text', 'bodytext',
    'caption', 'footnote', 'endnote', 'header', 'footer',
    'toc 1', 'toc 2', 'toc 3', 'tocheading',
    'table grid', 'list paragraph',
})


@dataclass(frozen=True)
class _Para:
    """文档段落快照：(样式名, 文本)。"""
    style: str
    text: str


@dataclass(frozen=True)
class _StyleStats:
    """样式统计：出现次数与每样式文本列表。"""
    counts: dict[str, int]        # 样式名 -> 出现次数
    texts: dict[str, list[str]]   # 样式名 -> 段落文本列表


@dataclass(frozen=True)
class _ScanResult:
    """一次扫描的产物：段落列表 + 样式统计。"""
    all_paras: list[_Para]
    stats: _StyleStats


@dataclass(frozen=True)
class _TableInfo:
    """表格快照：body 索引 + 是否有内容 + 文本预览 + 单元格内空段落数。"""
    idx: int
    has_content: bool
    preview: str
    empty_paras: int = 0


type _Paras = list[_Para]  # 段落列表，按文档顺序


def collect_paragraphs(body, qn) -> _ScanResult:
    """
    ## 收集段落信息、样式计数、每样式文本列表（用于长度/标题判断）。

    Args:
        body: docx.Document.body
        qn: docx.oxml.ns.qn 函数

    Returns:
        scan (ScanResult): 一次扫描的产物。
        - **all_paras** *(list[Para])*: 段落列表，按文档顺序。
        - **stats.counts** *(dict[str, int])*: 样式名 -> 出现次数。
        - **stats.texts** *(dict[str, list[str]])*: 样式名 -> 段落文本列表。
    """
    all_paras: list[_Para] = []
    counts: dict[str, int] = {}
    texts: dict[str, list[str]] = {}

    for p_elem in body.iter(qn("w:p")):
        style = xmlutils.get_style_name(p_elem, qn)
        text = xmlutils.text_of(p_elem, qn).strip()
        all_paras.append(_Para(style, text))
        if style and text:
            counts[style] = counts.get(style, 0) + 1
            texts.setdefault(style, []).append(text)

    return _ScanResult(all_paras, _StyleStats(counts, texts))


def collect_tables(body, qn) -> list[_TableInfo]:
    """
    ## 收集 body 直接子元素中的表格（w:tbl）信息。

    段落与表格按文档顺序混排，idx 即该表格在 body 中的索引，
    供 AI 判断「标题后是否有空表格作为答案区」。

    Args:
        body: docx.Document.body
        qn: docx.oxml.ns.qn 函数

    Returns:
        tables (list[_TableInfo]): 表格快照列表，按文档顺序。
        - empty_paras 统计**单元格内空段落数**——表内占位答案区的判据
          （如「标题 + 签名表(内含 41 个空段)」的书写区）。
    """
    tables: list[_TableInfo] = []
    for i, child in enumerate(body):
        if child.tag != qn('w:tbl'):
            continue
        paras = child.findall('.//' + qn('w:p'))
        texts = [''.join(t.text or '' for t in p.iter(qn('w:t'))).strip() for p in paras]
        empties = sum(1 for t in texts if not t)
        joined = ''.join(texts).strip()
        tables.append(_TableInfo(
            idx=i, has_content=bool(joined), preview=joined[:40], empty_paras=empties,
        ))

    return tables


def print_tables(tables: list[_TableInfo]) -> None:
    """打印表格结构，标注空表格与单元格内空段数。"""
    if not tables:
        return

    empty = sum(1 for t in tables if not t.has_content)
    print(f"\nTables: {len(tables)} total, {empty} empty")
    for t in tables:
        hint = f"  单元格内空段: {t.empty_paras}" if t.empty_paras else ""
        if t.has_content:
            print(f"  idx={t.idx} [有内容] {t.preview}{hint}")
        else:
            print(f"  idx={t.idx} [空表格]{hint}")


def identify_headings(scan: _ScanResult) -> tuple[_Paras, set[str]]:
    """
    ## 识别标题段落：Word 内置 Heading 或「短文本 + 高频」的自定义样式
    - 关键：正文样式（如 a8）的段落较长，标题样式（如 a4）的段落较短 ——
    用平均文本长度区分，避免把正文样式误判为标题。

    Args:
        scan (_ScanResult): 包含 all_paras 与 stats。

    Returns:
        (headings, heading_styles) (tuple[list[_Para], set[str]]):
        1. **headings** *(list[_Para])*: 标题段落列表，按文档顺序。
        2. **heading_styles** *(set[str])*: 被识别为标题的自定义样式名集合（不含内置 Heading 样式）。
    """
    headings: list[_Para] = []
    heading_styles: set[str] = set()
    for p in scan.all_paras:
        if not p.text:
            continue

        if p.style and p.style.lower().startswith("heading"):
            headings.append(p)

        elif p.style and _is_custom_heading_style(p.style, scan.stats):
            headings.append(p)
            heading_styles.add(p.style)

    return headings, heading_styles


def find_empty_sections(all_paras: _Paras, headings: _Paras) -> set[int]:
    """
    ## 检测哪些标题节为空（标题后无任何非空正文）。

    **用指针顺序匹配**：
    - 按文档顺序遍历段落;
    - 仅当段落文本等于下一个待匹配标题的文本时才推进指针;
    - 避免重复标题或正文恰好与标题同名时误判；
    - 标题之间的非空正文标记当前标题非空。

    Args:
        all_paras (list[Para]): 段落列表，按文档顺序。
        headings (list[Para]): 标题段落列表，按文档顺序。

    Returns:
        empty_set (set[int]): 空节标题的索引集合（在 headings 中的索引）。
    """
    empty_set = set(range(len(headings)))
    ptr = 0       # 指向下一个待匹配的标题
    current = -1  # 当前所在标题索引

    for p in all_paras:
        if not p.text:
            continue

        # 段落文本等于下一个待匹配标题 -> 进入该标题
        if ptr < len(headings) and p.text == headings[ptr].text:
            current = ptr
            ptr += 1
            continue

        # 非标题正文 -> 标记当前标题非空
        if current >= 0:
            empty_set.discard(current)

    return empty_set


def print_structure(headings: _Paras, empty_set: set[int]) -> int:
    """
    ## 打印标题结构，标注空节，返回总标题数。

    Args:
        headings (list[Para]): 标题段落列表，按文档顺序。
        empty_set (set[int]): 空节标题的索引集合（在 headings 中的索引）。

    Returns:
        total (int): 标题总数。
    """
    print("Document structure:")
    style_to_depth: dict[str, int] = {}  # 样式名 -> 层级深度
    if headings:
        for p in headings:
            if p.style in style_to_depth:
                continue

            if p.style.lower().startswith("heading"):
                try:
                    d = int(p.style.split()[-1]) - 1
                except ValueError:
                    d = len(style_to_depth)
            else:
                d = len(style_to_depth)

            style_to_depth[p.style] = d

        min_depth = min(style_to_depth.values())
        for hi, p in enumerate(headings):
            d = style_to_depth.get(p.style, 0) - min_depth
            indent = "  " * max(d, 0)
            empty_tag = " [EMPTY]" if hi in empty_set and d > 0 else ""
            print(f"  {indent}[{p.style}] {p.text}{empty_tag}")

    return _count_total_headings(headings, style_to_depth, empty_set)


def print_style_hints(heading_styles: set[str], stats: _StyleStats, total: int) -> None:
    """
    ## 样式提示
    - 排除图片标题样式（非节边界），按出现次数排序取最频繁的图片标题样式
    （如 a3，段落多以"图"/"Figure"开头）不是节边界，不应作为 `--heading-style`

    Args:
        heading_styles (set[str]): 被识别为标题的自定义样式名集合（不含内置 Heading 样式）。
        stats (StyleStats): 样式统计（counts + texts）。
        total (int): 标题总数。
    """
    boundary_styles = [s for s in heading_styles if not _is_caption_style(s, stats)]
    if boundary_styles:
        ranked = sorted(boundary_styles, key=lambda s: stats.counts.get(s, 0), reverse=True)
        styles_str = ", ".join(sorted(heading_styles))
        primary = ranked[0]
        print(f"Custom heading styles: {styles_str}")
        print(f"Hint: use --heading-style {primary} for section filling.")

    elif heading_styles:
        # 只有图片标题样式时，回退到全部标题样式
        ranked = sorted(heading_styles, key=lambda s: stats.counts.get(s, 0), reverse=True)
        print(f"Custom heading styles: {', '.join(ranked)}")
        print(f"Hint: use --heading-style {ranked[0]} for section filling.")

    else:
        print("No custom heading styles detected. Document uses standard Heading styles.")

    if total > 0:
        print('Hint: use scoped syntax "Parent / Child" to disambiguate duplicate headings.')


def _is_custom_heading_style(style: str, stats: _StyleStats) -> bool:
    """
    判断自定义样式是否为标题样式（而非正文样式）。
    - **判据**：出现 >= 2 次、非已知正文样式名、且段落平均文本长度较短
    （标题短、正文长）。阈值 60 字符可区分 a4 标题与 a8 正文。
    """
    if style.lower() in _BODY_STYLE_NAMES:
        return False
    if stats.counts.get(style, 0) < 2:
        return False

    texts = stats.texts.get(style, [])
    if not texts:
        return False

    avg_len = sum(len(t) for t in texts) / len(texts)
    return avg_len <= 60


def _count_total_headings(headings: _Paras, style_to_depth: dict[str, int], empty_set: set[int]) -> int:
    """统计标题数量"""
    total = len(headings)
    if headings and style_to_depth:
        min_d = min(style_to_depth.values())
        empty_count = sum(
            1 for hi in empty_set
            if style_to_depth.get(headings[hi].style, 0) > min_d
        )
    else:
        empty_count = 0

    print(f"\nHeadings: {total} total, {empty_count} empty sub-sections")
    return total


def _is_caption_style(style: str, stats: _StyleStats) -> bool:
    """
    判断样式是否为图片标题样式（段落多以"图"/"Figure"/"Fig"开头）。
    - 图片标题不是节边界，不应作为 `--heading-style` 建议。
    """
    texts = stats.texts.get(style, [])
    if not texts:
        return False

    caption_count = sum(1 for t in texts if t.startswith(("图", "Figure", "Fig")))
    return caption_count / len(texts) > 0.5
