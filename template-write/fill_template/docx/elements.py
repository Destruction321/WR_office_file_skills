"""
# docx 元素操作 — 段落构建与插入/删除。
- `remove_elements_between` / `insert_items` 由 `section_filler` 调用。
- `mark_section_start` / `clear_section_marks` / `count_paras_in_section` 支撑内存验证。
- `qn` / `OxmlElement` / `Inches` 通过 `ensure_import` 延迟加载，本模块自管 docx 依赖。
"""

from pathlib import Path
from typing import Any

from ..deps import ensure_import

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


def insert_items(doc, anchor, items: list[dict[str, Any]]) -> tuple[int, int, Any, Any]:
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
        if item.get("type") == "image":
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


def mark_section_start(first_para, last_para, sec_id: str) -> None:
    """
    在 first_para 前插入 bookmarkStart、last_para 后插入 bookmarkEnd
    以包裹本节插入的所有段落。bookmark name = sec_id，id 用高基数避免与文档已有 bookmark 冲突。
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
    `<w:p>` 元素（仅 body 直接子级，不含表格内嵌段落），返回 (total_paras, non_empty_paras)。
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
    for el in body:
        if el.tag == start_tag:
            if el.get(id_attr) == target_id:
                counting = True
            continue
        if el.tag == end_tag:
            if el.get(id_attr) == target_id:
                break
            continue
        if counting and el.tag == p_tag:
            total += 1
            text = "".join(t.text or "" for t in el.iter(t_tag)).strip()
            if text:
                non_empty += 1
    return total, non_empty


def _build_image_para(doc, item: dict[str, Any]):
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
    run.add_picture(str(image_path), width=_Inches(width))
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
