"""
# 内容项类型 — 节级填充的统一数据模型。

- `Run` / `ImageItem` / `ParagraphItem` 取代原先跨模块传递的 `dict[str, Any]`，
  用 `isinstance` 分发取代字符串 `"type"` 判别。
- `item_from_dict` 在 JSON 边界把字典加载为对应类型；Markdown 解析（`md_parser`）直接构造。
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class Run:
    """
    行内文本片段，可带 bold/italic。
    
    Attributes:
        text (str): 文本内容。
        bold (bool): 是否加粗。
        italic (bool): 是否斜体。
    """
    text: str
    bold: bool = False
    italic: bool = False


@dataclass(frozen=True)
class ImageItem:
    """
    内嵌图片项。
    
    Attributes:
        path (str): 图片文件路径。
        width_inches (float | None): 图片宽度（英寸）。
    """
    path: str
    width_inches: float | None = None


@dataclass(frozen=True)
class ParagraphItem:
    """
    文本段落项，由若干 Run 组成。
    
    Attributes:
        runs (list[Run]): Run 列表。
    """
    runs: list[Run]


type Item = ImageItem | ParagraphItem


def item_from_dict(d: dict) -> Item:
    """
    ## 从 dict（JSON 加载）构造内容项。

    - `{"type": "image", "path": ..., "width_inches": ...}` -> `ImageItem`
    - `{"type": "paragraph", "runs": [...]}` -> `ParagraphItem`
    - 旧格式 `{"type": "paragraph", "text": ..., "bold": ...}` -> 单 Run 的 `ParagraphItem`

    Args:
        d (dict): JSON 解析得到的内容项字典。

    Returns:
        body_items (Item): 对应的 ImageItem 或 ParagraphItem。
    """
    if d.get("type") == "image":
        return ImageItem(path=d["path"], width_inches=d.get("width_inches"))

    if "runs" in d:
        runs = [
            Run(
                text=r.get("text", ""),
                bold=bool(r.get("bold", False)),
                italic=bool(r.get("italic", False))
            )
            for r in d["runs"]
        ]
        return ParagraphItem(runs=runs)

    # 旧格式：单 Run 的段落
    runs = [Run(text=d.get("text", ""), bold=bool(d.get("bold", False)))]
    return ParagraphItem(runs=runs)
