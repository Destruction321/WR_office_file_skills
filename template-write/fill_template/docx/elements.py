"""
# docx 元素操作 — 段落构建与插入/删除。
- `remove_elements_between` / `insert_items` 由 `section_filler` 调用。
"""

from pathlib import Path
from typing import Any


def remove_elements_between(body, children, start: int, end: int | None) -> None:
    """
    ## 删除 body 中 (start, end) 之间的所有子元素。
    
    Args:
        body: docx.Document.body
        children: list of body children
        start (int): 起始索引（不含）
        end (int | None): 结束索引（不含），None 表示到文档末尾。
    """
    actual_end = end if end is not None else len(children)
    for i in range(actual_end - 1, start, -1):
        body.remove(children[i])


def insert_items(doc, anchor, items: list[dict[str, Any]], Inches) -> tuple[int, int]:
    """
    ## 将内容项逐个构建为段落，用 addnext 链式插入到 anchor 之后。
    - image 项走 `_build_image_para`，其余走 `_build_text_para`。
    - 每次插入后把 anchor 前移到新段落，保证顺序与 items 一致。
    
    Args:
        doc: docx.Document
        anchor: 插入锚点段落（docx.oxml.CT_P）
        items (list[dict[str, Any]]): 内容项列表
        Inches: docx.shared.Inches 函数
        
    Returns:
        (para_count, img_count) (tuple[int, int]): 插入的段落数与图片数。
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
