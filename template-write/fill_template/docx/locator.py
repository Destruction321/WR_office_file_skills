"""
# docx 节定位器 — 标题索引查找与节边界检测。
- `locate_section` / `find_heading_index_scoped` 由 `section_filler` 调用。
"""

from dataclasses import dataclass
from typing import Any

from . import xmlutils

SCOPE_SEP = ' / '


@dataclass
class LocateContext:
    """
    ## 段落定位上下文管理模块
    
    Attributes:
        children (list[Any]): `docx.Document.body` 列表
        qn (Any): `docx.oxml.ns.qn` 函数
        hs_lower (str | None): 指定的 `--heading-style`，统一小写。
    """
    children: list[Any]
    qn: Any
    hs_lower: str | None = None


def locate_section(locate_ctx: LocateContext, heading_text: str) -> tuple[int, int | None] | None:
    """
    ## 定位 (heading_idx, end_idx)：先查标题索引，再以该标题样式算节边界。
    - **全局定位：**"标题文本"；限定定位："父标题 / 子标题"。
    - 标题索引查找共用 `find_heading_index_scoped`，与 replace 后重定位一致。
    
    Args:
        locate_ctx (LocateContext): 包含 children、qn、hs_lower。
        heading_text (str): 标题文本，支持限定定位语法。
    
    Returns:
        (heading_idx, end_idx) (tuple[int, int | None]): 
            标题段落索引与节边界索引，end_idx 为 None 表示到文档末尾。

        **None**: 未找到标题段落。
    """
    heading_idx = find_heading_index_scoped(locate_ctx, heading_text)
    if heading_idx is None:
        return None

    end_idx = _find_style_boundary(
        locate_ctx,
        heading_idx,
        xmlutils.get_style_name(locate_ctx.children[heading_idx], locate_ctx.qn)
    )
    return heading_idx, end_idx


def find_heading_index_scoped(locate_ctx: LocateContext, heading_text: str) -> int | None:
    """
    ## 按标题文本查找段落索引，支持限定定位语法。
    - 全局："标题文本" -> 全文找第一个匹配段落。
    - 限定："父标题 / 子标题" -> 先找父标题，再在其后找子标题
    （不限制 parent_end，因父子可能同样式，样式边界无法区分）。
    
    Args:
        locate_ctx (LocateContext): 包含 children、qn、hs_lower。
        heading_text (str): 标题文本，支持限定定位语法。

    Returns:
        int | None: 标题段落在 children 中的索引，未找到返回 None。
    """
    parent_text, sep, child_text = heading_text.partition(SCOPE_SEP)
    if not sep:
        return _find_heading_index(locate_ctx, parent_text)

    parent_idx = _find_heading_index(locate_ctx, parent_text)
    if parent_idx is None:
        return None
    return _find_heading_index(locate_ctx, child_text.strip(), start=parent_idx + 1)


def _find_heading_index(locate_ctx: LocateContext, text: str, start: int = 0) -> int | None:
    """在 children[start:len(children)] 范围内找到包含 text 的段落，返回索引。大小写不敏感。"""
    text_lower = text.strip().lower()
    end = len(locate_ctx.children)
    for i in range(start, end):
        child = locate_ctx.children[i]
        if child.tag != locate_ctx.qn("w:p"):
            continue
        
        if text_lower in xmlutils.text_of(child, locate_ctx.qn).lower():
            return i
    
    return None


def _find_style_boundary(locate_ctx: LocateContext,
                         heading_idx: int,
                         heading_style: str) -> int | None:
    """
    找到 heading_idx 之后第一个同级标题段落，返回其索引。

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

        style_lower = xmlutils.get_style_name(child, locate_ctx.qn).lower()

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
