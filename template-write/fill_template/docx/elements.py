"""
# docx 元素操作 — 段落构建与插入/删除。
- `remove_elements_between` / `insert_items` 由 `section_filler` 调用。
- `mark_section_start` / `clear_section_marks` / `count_paras_in_section` 支撑内存验证。
- `qn` / `OxmlElement` / `Inches` 通过 `ensure_import` 延迟加载，本模块自管 docx 依赖。
"""

from pathlib import Path
from typing import Any

from ..deps import ensure_import
from ..items import ImageItem, Item, ParagraphItem
    

_qn = ensure_import("python-docx", "docx.oxml.ns", attr="qn")
_OxmlElement = ensure_import("python-docx", "docx.oxml", attr="OxmlElement")
_Inches = ensure_import("python-docx", "docx.shared", attr="Inches")

_SEC_MARK_PREFIX = "sec:"


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


def insert_items(doc, anchor, items: list[Item]) -> tuple[int, int, Any, Any]:
    """
    ## 将内容项逐个构建为段落，用 addnext 链式插入到 anchor 之后。
    - image 项走 `_build_image_para`，其余走 `_build_text_para`。
    - 每次插入后把 anchor 前移到新段落，保证顺序与 items 一致。

    Args:
        doc: docx.Document
        anchor: 插入锚点段落（docx.oxml.CT_P）
        items (list[dict[str, Any]]): 内容项列表

    Returns:
        (para_count, img_count, first_p, last_p) (tuple[int, int, Any, Any]): 
            插入的段落数与图片数，以及首末段落（供调用方打 bookmark 标记）。
    """
    para_count = 0
    img_count = 0
    first_p = None
    for item in items:
        if isinstance(item, ImageItem):
            p = _build_image_para(doc, item)
            img_count += 1
        else:
            p = _build_text_para(doc, item)
            para_count += 1
        anchor.addnext(p)
        anchor = p
        if first_p is None:
            first_p = p

    return para_count, img_count, first_p, anchor


def insert_items_into_cell(cell, doc, items: list[Item]) -> tuple[int, int, Any, Any]:
    """
    ## 将内容项逐个构建为段落，插入到单元格终止段（最后一个 w:p）之前。

    每个新段落都插在终止段前，保持 items 顺序；终止段保留以满足
    w:tc 必须以 w:p 结尾的架构约束（通常是签名行等既有内容）。

    Args:
        cell: w:tc 元素（目标单元格）。
        doc: docx.Document
        items (list[Item]): 内容项列表。

    Returns:
        (para_count, img_count, first_p, last_p) (tuple[int, int, Any, Any]):
            插入的段落数与图片数，以及首末段落（供调用方打 bookmark 标记）。
    """
    p_tag = _qn("w:p")
    paras = cell.findall(p_tag)
    if not paras:
        raise ValueError('表格单元格内没有段落，无法插入')
    terminal = paras[-1]

    para_count = 0
    img_count = 0
    first_p = None
    last_p = None
    for item in items:
        if isinstance(item, ImageItem):
            p = _build_image_para(doc, item)
            img_count += 1
        else:
            p = _build_text_para(doc, item)
            para_count += 1
        terminal.addprevious(p)
        if first_p is None:
            first_p = p
        last_p = p

    return para_count, img_count, first_p, last_p


def remove_empty_paras_in_cell(cell) -> int:
    """
    ## 删除单元格内的空段落（保留终止段，满足 w:tc 架构约束），返回删除数。

    表内占位答案区的典型形态：单元格内 N 个空段落 + 签名行。
    填充前删除这些占位空段，再在终止段前插入内容。

    Args:
        cell: w:tc 元素。

    Returns:
        removed (int): 删除的空段落数。
    """
    p_tag = _qn("w:p")
    t_tag = _qn("w:t")
    paras = cell.findall(p_tag)
    if not paras:
        return 0
    terminal = paras[-1]

    removed = 0
    for p in paras[:-1]:
        text = ''.join(t.text or '' for t in p.iter(t_tag)).strip()
        if text:
            continue
        cell.remove(p)
        removed += 1
    return removed


def mark_section_start(first_para, last_para, sec_id: str) -> None:
    """
    在 first_para 前插入 bookmarkStart、last_para 后插入 bookmarkEnd
    以包裹本节插入的所有段落，bookmark name = sec_id，id 用高基数避免与文档已有 bookmark 冲突。
    
    Args:
        first_para: docx.oxml.CT_P，首段落
        last_para: docx.oxml.CT_P，末段落
        sec_id (str): bookmark name，通常为 "sec:数字"。
    """
    num = sec_id.split(":", 1)[1] if ":" in sec_id else "0"
    bm_id = f"9{int(num):04d}" if num.isdigit() else "90000"
    bm_start = _OxmlElement("w:bookmarkStart")
    bm_start.set(_qn("w:id"), bm_id)
    bm_start.set(_qn("w:name"), sec_id)
    first_para.addprevious(bm_start)

    bm_end = _OxmlElement("w:bookmarkEnd")
    bm_end.set(_qn("w:id"), bm_id)
    last_para.addnext(bm_end)


def clear_section_marks(body) -> None:
    """
    删除 body 中所有 name 以 `sec:` 开头的 bookmarkStart，及其配对的
    bookmarkEnd，配对靠 w:id 相等。
    
    Args:
        body: docx.Document.body
    """
    start_tag = _qn("w:bookmarkStart")
    end_tag = _qn("w:bookmarkEnd")
    name_attr = _qn("w:name")
    id_attr = _qn("w:id")

    sec_ids: set[str] = set()
    for el in body.iter(start_tag):
        if el.get(name_attr, "").startswith(_SEC_MARK_PREFIX):
            sec_ids.add(el.get(id_attr))
            el.getparent().remove(el)
    
    if not sec_ids:
        return
    
    for el in body.iter(end_tag):
        if el.get(id_attr) in sec_ids:
            el.getparent().remove(el)


def count_paras_in_section(body, sec_id: str) -> tuple[int, int]:
    """
    找到 name==sec_id 的 bookmarkStart，数到对应 bookmarkEnd 之间的
    `<w:p>` 元素（递归遍历，body 级与表格单元格内的标记均支持），
    返回 (total_paras, non_empty_paras)。

    Args:
        body: docx.Document.body
        sec_id (str): bookmark name，通常为 "sec:数字"。
        
    Returns:
        (total_paras, non_empty_paras) (tuple[int, int]): 总段落数与非空段落数。
    """
    p_tag = _qn("w:p")
    start_tag = _qn("w:bookmarkStart")
    end_tag = _qn("w:bookmarkEnd")
    name_attr = _qn("w:name")
    id_attr = _qn("w:id")
    t_tag = _qn("w:t")

    target_id = None
    for el in body.iter(start_tag):
        if el.get(name_attr) == sec_id:
            target_id = el.get(id_attr)
            break

    if target_id is None:
        return 0, 0

    total = 0
    non_empty = 0
    counting = False
    for el in body.iter():
        if el.tag == start_tag:
            if el.get(id_attr) == target_id:
                counting = True
            continue

        if el.tag == end_tag:
            if el.get(id_attr) == target_id:
                break
            continue

        if not counting or el.tag != p_tag:
            continue

        total += 1
        if ("".join(t.text or "" for t in el.iter(t_tag)).strip()):
            non_empty += 1

    return total, non_empty


def _build_image_para(doc, item: ImageItem):
    """构建内嵌图片段落。"""
    image_path = Path(item.path)
    if image_path.exists():
        width = item.width_inches if item.width_inches is not None else 5.5
        p = doc.add_paragraph()
        run = p.add_run()
        run.add_picture(str(image_path), width=_Inches(width))
    else:
        p = doc.add_paragraph()
        run = p.add_run()
        run.text = f"[Image not found: {image_path.name}]"
        run.bold = True

    p._element.getparent().remove(p._element)
    return p._element


def _build_text_para(doc, item: ParagraphItem):
    """构建文本段落，支持行内加粗/斜体（runs）。"""
    p = doc.add_paragraph()
    for r in item.runs:
        run = p.add_run()
        run.text = r.text
        if r.bold:
            run.bold = True
        if r.italic:
            run.italic = True
    p._element.getparent().remove(p._element)
    return p._element
