"""小节过滤 — 只提取文档中匹配关键字的部分，节省 tokens。"""
from re import compile, match, Pattern


# ===================================================================
#  中文标题模式 — 用于检测实验报告、论文中的章节标题
# ===================================================================

_CHINESE_HEADING_PATTERNS: list[tuple[Pattern, int]] = [
    (compile(r'^实验[一二三四五六七八九十百零\d]+[：:，,\s]'),          1),
    (compile(r'^第[一二三四五六七八九十百零\d]+章'),                    1),
    (compile(r'^第[一二三四五六七八九十百零\d]+节'),                    2),
    (compile(r'^[一二三四五六七八九十]+[、．.]\s*\S'),                  2),
    (compile(r'^（[一二三四五六七八九十]+）\s*\S'),                     3),
    (compile(r'^\([一二三四五六七八九十\d]+\)\s*\S'),                   3),
    (compile(r'^\d+\.\d+\s*\S'),                                        3),
    (compile(r'^\d+[、．.]\s*\S'),                                      2),
]


def detect_chinese_heading(text: str) -> int | None:
    """
    从中文字段模式检测标题级别，非标题则返回 None。

    Args:
        text: 要检测的文本行。

    Returns:
        标题级别（1-3）或 None（非标题）。
    """
    if len(text) > 100:
        return None

    for pat, level in _CHINESE_HEADING_PATTERNS:
        if pat.match(text):
            return level

    return None


# ===================================================================
#  主过滤入口
# ===================================================================

def filter_section(lines: list[str], section: str) -> list[str]:
    """
    过滤提取出的文本，仅保留匹配 section 关键字的小节。

    自动检测输出结构类型（Markdown 标题 / 幻灯片 / 页面 / 工作表）
    并选择对应的过滤策略；无匹配结构时回退到关键字上下文提取。

    Args:
        lines: 提取出的全部文本行。
        section: 要匹配的小节关键字。

    Returns:
        仅包含匹配小节的行。无匹配时返回 [未找到...] 单行列表。
    """
    if not section:
        return lines

    section_low = section.lower()

    # 从标记检测输出结构
    has_md_headings = any(match(r'^#{1,6}\s+\S', line) for line in lines[:50])

    if has_md_headings:
        return _filter_by_heading_markers(lines, section_low)

    if any(line.startswith('--- Slide') for line in lines):
        return _filter_by_unit_markers(lines, section_low, '--- Slide', context=1)

    if any(line.startswith('--- Page') for line in lines):
        return _filter_by_unit_markers(lines, section_low, '--- Page', context=1)

    if any(line.startswith('--- Sheet:') for line in lines):
        return _filter_by_unit_markers(lines, section_low, '--- Sheet:', context=0)

    # 回退：关键字上下文提取
    return _filter_by_keyword_context(lines, section_low)


def _filter_by_heading_markers(lines: list[str], section_low: str) -> list[str]:
    """
    按 Markdown 标题标记（# ## ###）过滤。

    从匹配关键字的标题开始，到同级或更高级别标题结束。
    """
    result: list[str] = []
    in_section = False
    section_level: int = 0

    for line in lines:
        m = match(r'^(#{1,6})\s+(.+)', line)
        if m:
            level = len(m.group(1))
            heading_text = m.group(2).lower()

            if not in_section and section_low in heading_text:
                in_section = True
                section_level = level
                result.append(line)
                continue
            elif in_section and level <= section_level:
                break

        if in_section:
            result.append(line)

    if not result:
        result = _filter_by_keyword_context(lines, section_low)

    return result


def _filter_by_unit_markers(lines: list[str],
                            section_low: str,
                            marker_prefix: str,
                            context: int = 1) -> list[str]:
    """
    过滤包含关键字的幻灯片/页面/工作表，带 context 个相邻单元。

    例如 context=1 会同时包含匹配幻灯片的前后各一页。
    """
    units: list[list[str]] = []
    current: list[str] = []
    for line in lines:
        if line.startswith(marker_prefix) and current:
            units.append(current)
            current = []
        current.append(line)

    if current:
        units.append(current)

    matched: set[int] = set()
    for idx, unit in enumerate(units):
        for line in unit:
            if section_low in line.lower():
                matched.add(idx)
                break

    with_context: set[int] = set()
    for idx in matched:
        for delta in range(-context, context + 1):
            neighbour = idx + delta
            if 0 <= neighbour < len(units):
                with_context.add(neighbour)

    result = [line for idx in sorted(with_context) for line in units[idx]]

    if not result:
        result = _filter_by_keyword_context(lines, section_low)

    return result


def _filter_by_keyword_context(lines: list[str],
                               section_low: str,
                               before: int = 3,
                               after: int = 50) -> list[str]:
    """回退方案：提取关键字附近的行（前 3 行 + 后 50 行）。"""
    match_indices = [i for i, line in enumerate(lines) if section_low in line.lower()]
    if not match_indices:
        return [f'[未找到匹配 "{section_low}" 的内容]']

    included: set[int] = set()
    for idx in match_indices:
        for i in range(max(0, idx - before), min(len(lines), idx + after)):
            included.add(i)

    return [lines[i] for i in sorted(included)]
