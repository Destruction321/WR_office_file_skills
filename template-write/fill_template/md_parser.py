"""
# Markdown 节级规格解析器。

- 将 Markdown 文本按标题分割为「标题 -> 内容项列表」的映射，
  供 `fill_docx_sections` 使用。

## 支持的 Markdown 语法

- `# ~ ###### 标题` -> 节分隔符（标题文本用于匹配文档标题）
- `## 父标题 / 子标题` -> 限定定位：先在文档中找到父标题，再在其节范围内找子标题，消除同名标题歧义
- `![alt](path)` -> 图片；可选 `{width=5.0}` 指定宽度（英寸）
- `**bold**` 和 `*italic*` -> 行内格式
- `1. ` / `- ` -> 列表项（各成一段落，保留前缀）
- 空行 -> 段落分隔
- 连续非空行 -> 合并为同一段落
"""

from re import compile

from . import items

# 标题行：1-6 个 # 开头
_HEADING_RE = compile(r'^(#{1,6})\s+(.+)$')

# 图片：![alt](path) 可选 {width=5.0}
_IMAGE_RE = compile(r'!\[.*?\]\((.+?)\)(?:\{width=([\d.]+)\})?')

# 行内格式：**bold** 和 *italic*
_INLINE_RE = compile(r'(\*\*.+?\*\*|\*.+?\*)')

# 列表项：1. 或 - 开头
_LIST_ITEM_RE = compile(r'^(\d+\.\s|[-*]\s)')


def parse_sections_md(md_text: str) -> dict[str, list[items.Item]]:
    """
    ## 将 Markdown 文本解析为节标题 -> 内容项列表的映射。

    每遇到一个 `# 标题` 行，就开始一个新节。
    标题文本用于在 docx 中匹配对应的标题段落。

    Args:
        md_text (str): Markdown 文本。

    Returns:
        sections (dict[str, list[items.Item]]): 标题文本 -> 内容项列表。
    """
    sections: dict[str, list[items.Item]] = {}
    current_heading: str | None = None
    current_items: list[items.Item] = []
    pending_lines: list[str] = []

    for line in md_text.splitlines():
        stripped = line.strip()

        # 标题行 -> 保存当前节，开始新节
        heading_match = _HEADING_RE.match(stripped)
        if heading_match:
            _flush_pending(pending_lines, current_items)
            pending_lines = []
            if current_heading is not None:
                sections[current_heading] = current_items

            current_heading = heading_match.group(2).strip()
            current_items = []
            continue

        # 空行 -> 结束当前段落
        if not stripped:
            _flush_pending(pending_lines, current_items)
            pending_lines = []
            continue

        # 图片行（独立成段）
        img_match = _IMAGE_RE.match(stripped)
        if img_match:
            _flush_pending(pending_lines, current_items)
            pending_lines = []
            current_items.append(items.ImageItem(
                path=img_match.group(1),
                width_inches=float(img_match.group(2)) if img_match.group(2) else None,
            ))
            continue

        # 列表项 -> 独立成段（每个列表项一个段落）
        if _LIST_ITEM_RE.match(stripped):
            _flush_pending(pending_lines, current_items)
            pending_lines = [stripped]
            continue

        # 普通文本 -> 收集到当前段落
        pending_lines.append(stripped)

    # 保存最后的内容
    _flush_pending(pending_lines, current_items)
    if current_heading is not None:
        sections[current_heading] = current_items

    return sections


def _flush_pending(lines: list[str], current_items: list[items.Item]) -> None:
    """将收集到的文本行合并为一个段落项，追加到 items。"""
    if not lines:
        return
    # 同一段内的连续行用空格连接
    full_text = ' '.join(lines)
    runs = _parse_inline(full_text)
    current_items.append(items.ParagraphItem(runs=runs))


def _parse_inline(text: str) -> list[items.Run]:
    """
    解析 **bold** 和 *italic* 行内格式为 Run 列表。

    - `**bold**` -> `Run(text="bold", bold=True)`
    - `*italic*` -> `Run(text="italic", italic=True)`
    - 普通文本 -> `Run(text="...")`
    """
    parts = _INLINE_RE.split(text)
    runs: list[items.Run] = []
    for part in parts:
        if not part:
            continue

        if part.startswith('**') and part.endswith('**'):
            runs.append(items.Run(text=part[2:-2], bold=True))
        elif part.startswith('*') and part.endswith('*'):
            runs.append(items.Run(text=part[1:-1], italic=True))
        else:
            runs.append(items.Run(text=part))

    return runs
