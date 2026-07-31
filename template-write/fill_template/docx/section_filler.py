"""
# docx 节级内容填充器

## 核心原则：
1. 文本匹配定位 + 样式名定义界。不推断标题级别。
2. **工具只做机械操作**：定位标题 -> 确定边界 -> 清除旧内容 -> 插入新内容 -> 验证。
智能判断由 Agent 通过 --scan 输出完成。

## 标题定位：
1. 全局定位："实验过程及分析" — 全文搜索第一个匹配段落
2. 限定定位："实验七 / 实验过程及分析" — 先找父标题，再在范围内找子标题

## 边界检测策略：
1. 相同样式名的段落 -> 同级标题 -> 节边界
2. Word 内置 Heading 样式的段落 -> 节边界
"""

from dataclasses import dataclass, replace
from pathlib import Path
from sys import stderr
from typing import Any

from . import elements, locator, scanner
from ..deps import ensure_import
from ..items import ImageItem, Item


@dataclass(frozen=True)
class _FillSession:
    """文档填充会话"""
    body: Any
    doc: Any
    locate_ctx: locator.LocateContext
    
    def refresh_locate_ctx(self, new_body) -> _FillSession:
        """
        ## 返回一个新的 _FillSession，body 刷新为 new_body。
        
        Args:
            new_body: docx.Document.body
        
        Returns:
            new_session (_FillSession): 新的 _FillSession，body 刷新为 new_body。
        """
        return replace(self, locate_ctx=self.locate_ctx.refresh_body(new_body))


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
    Document = ensure_import("python-docx", "docx", attr="Document")
    qn = ensure_import("python-docx", "docx.oxml.ns", attr="qn")
    body = Document(str(doc_path)).element.body
    scan = scanner.collect_paragraphs(body, qn)
    headings, heading_styles = scanner.identify_headings(scan)
    empty_set = scanner.find_empty_sections(scan.all_paras, headings)
    total = scanner.print_structure(headings, empty_set)
    scanner.print_tables(scanner.collect_tables(body, qn))
    scanner.print_style_hints(heading_styles, scan.stats, total)


def fill_docx_sections(doc_path: Path,
                       sections: dict[str, list[Item]],*,
                       mode: str = "replace",
                       heading_style: str | None = None,
                       dry_run: bool = False) -> int:
    """
    ## 按节标题填充 `.docx` 文档内容。

    Args:
        doc_path (Path): 已复制的 `.docx` 文件路径（原地修改）。
        sections (dict[str, list[Item]]): 标题文本 -> 内容项列表。支持 "父标题 / 子标题" 限定定位。
        mode (str): 'replace'（清空旧内容再填充）或 'append'（追加）。
        heading_style (str | None): 手动指定标题样式名（如 'a4'），用于边界检测。
        dry_run (bool): 仅检测定位和边界，不修改文件。

    Returns:
        int: 成功填充的节数量。
    """
    session = _open_fill_session(doc_path, heading_style)
    filled_count = 0
    missed: list[str] = []
    filled_summary: list[tuple[str, int, int, str]] = []

    for sec_idx, (heading_text, items) in enumerate(sections.items()):
        session = session.refresh_locate_ctx(session.body)

        loc = locator.locate_section(session.locate_ctx, heading_text)
        if loc is None:
            missed.append(heading_text)
            continue
        heading_idx, end_idx = loc

        if dry_run:
            pc = sum(1 for it in items if not isinstance(it, ImageItem))
            ic = sum(1 for it in items if isinstance(it, ImageItem))
            if mode == "cell":
                cell = _find_placeholder_cell(session.locate_ctx.qn, session.locate_ctx.children, heading_idx)
                if cell is None:
                    raise ValueError(
                        f"未找到标题 '{heading_text}' 后的表内占位单元格"
                        f"（标题后应为含空段落的表格）。请改用 --section-mode append。"
                    )
                print(
                    f"  [DRY] '{heading_text}' -> idx={heading_idx}, cell=占位单元格, "
                    f"{pc} paragraphs + {ic} images"
                )
            else:
                end_desc = str(end_idx) if end_idx is not None else "end"
                print(
                    f"  [DRY] '{heading_text}' -> idx={heading_idx}, end={end_desc}, "
                    f"{pc} paragraphs + {ic} images"
                )
            filled_count += 1
            continue

        # 清除旧内容（replace 模式）—— 清除后索引移位，需重新定位标题作锚点
        if mode == "replace":
            if end_idx is None:
                # 安全护栏：无样式边界时 replace 会清到文末，阻止误删后续节
                remaining = len(session.locate_ctx.children) - heading_idx - 1
                if remaining > 0:
                    raise ValueError(
                        f"节边界检测失败（标题 '{heading_text}' 后无同级标题或样式边界），"
                        f"replace 模式将清除其后的 {remaining} 个元素。"
                        f"请改用 --section-mode append，或用 --heading-style 指定边界样式。"
                    )
            elements.remove_elements_between(session.body, session.locate_ctx.children, heading_idx, end_idx)
            heading_idx = locator.find_heading_index_scoped(session.locate_ctx, heading_text)
            if heading_idx is None:
                missed.append(heading_text)
                continue

        # 表内占位替换（cell 模式）—— 删除标题后表格单元格内的占位空段，内容写入单元格内
        if mode == "cell":
            cell = _find_placeholder_cell(session.locate_ctx.qn, session.locate_ctx.children, heading_idx)
            if cell is None:
                raise ValueError(
                    f"未找到标题 '{heading_text}' 后的表内占位单元格"
                    f"（标题后应为含空段落的表格）。请改用 --section-mode append。"
                )
            if dry_run:
                pc = sum(1 for it in items if not isinstance(it, ImageItem))
                ic = sum(1 for it in items if isinstance(it, ImageItem))
                print(
                    f"  [DRY] '{heading_text}' -> idx={heading_idx}, cell=占位单元格, "
                    f"{pc} paragraphs + {ic} images"
                )
                filled_count += 1
                continue

            elements.remove_empty_paras_in_cell(cell)
            sec_id = f"sec:{sec_idx}"
            para_count, img_count, first_p, last_p = elements.insert_items_into_cell(
                cell, session.doc, items,
            )
            if first_p is not None:
                elements.mark_section_start(first_p, last_p, sec_id)

            filled_count += 1
            filled_summary.append((heading_text, para_count, img_count, sec_id))
            continue

        # 插入新内容
        anchor = session.locate_ctx.children[heading_idx]
        sec_id = f"sec:{sec_idx}"
        para_count, img_count, first_p, last_p = elements.insert_items(
            session.doc, anchor, items,
        )
        if first_p is not None:
            elements.mark_section_start(first_p, last_p, sec_id)

        filled_count += 1
        filled_summary.append((heading_text, para_count, img_count, sec_id))

    if dry_run:
        if missed:
            print(
                f"\nWarning: {len(missed)} heading(s) not found: "
                + ", ".join(f"'{t}'" for t in missed), file=stderr
            )
        print(f"\nDry-run: {filled_count}/{len(sections)} sections located, no changes made.")
        return filled_count

    # 内存验证（用 bookmark 标记定位）-> 清标记 -> save（干净落盘）
    _verify_filled(session.doc, filled_summary)
    elements.clear_section_marks(session.body)
    session.doc.save(str(doc_path))

    if missed:
        print(
            f"Warning: {len(missed)} heading(s) not found: "
            + ", ".join(f"'{t}'" for t in missed), file=stderr
        )
    return filled_count


# ===================================================================
#  会话与验证
# ===================================================================

def _find_placeholder_cell(qn, children, heading_idx: int):
    """
    ## 在标题之后寻找「表内占位答案区」的单元格。

    从标题后一个元素开始扫描：跳过空段落；遇到表格时返回其第一个
    含空段落的单元格（占位答案区的判据）；遇到非空段落或其他元素则放弃。

    Args:
        qn: docx.oxml.ns.qn 函数
        children: body 子元素元组
        heading_idx (int): 标题段落索引。

    Returns:
        cell (Any | None): 含空段落的 w:tc 元素；未找到返回 None。
    """
    p_tag = qn('w:p')
    t_tag = qn('w:t')
    tbl_tag = qn('w:tbl')
    for i in range(heading_idx + 1, len(children)):
        child = children[i]
        if child.tag == p_tag:
            text = ''.join(t.text or '' for t in child.iter(t_tag)).strip()
            if text:
                return None  # 非空段落：没有表内答案区
            continue  # 空段落：继续找表格
        if child.tag == tbl_tag:
            for tr in child.findall(qn('w:tr')):
                for tc in tr.findall(qn('w:tc')):
                    paras = tc.findall(p_tag)
                    if not paras:
                        continue
                    texts = [''.join(t.text or '' for t in p.iter(t_tag)).strip() for p in paras]
                    if any(not t for t in texts):
                        return tc
            return None
        return None  # 其他元素（如图片段落等）
    return None


def _open_fill_session(doc_path: Path, heading_style: str | None = None) -> _FillSession:
    """打开 docx 填充会话，返回 _FillSession 对象。"""
    Document = ensure_import("python-docx", "docx", attr="Document")
    qn = ensure_import("python-docx", "docx.oxml.ns", attr="qn")
    doc = Document(str(doc_path))
    body = doc.element.body
    hs_lower = heading_style.lower() if heading_style else None
    locate_ctx = locator.LocateContext(children=tuple(body), qn=qn, hs_lower=hs_lower)
    return _FillSession(body=body, doc=doc, locate_ctx=locate_ctx)


def _verify_filled(doc, filled_summary: list[tuple[str, int, int, str]]) -> None:
    """内存验证——按 bookmark 标记定位每个节插入的段落范围，数其中非空段落。

    不重新打开文件、不依赖文本匹配标题。bookmark 标记由 `insert_items` 打、
    `clear_section_marks` 清（save 前），故此处 doc 必须是填充后的内存对象。
    """
    if not filled_summary:
        return

    body = doc.element.body
    print("Verification:")
    for heading_text, para_count, img_count, sec_id in filled_summary:
        total, non_empty = elements.count_paras_in_section(body, sec_id)
        expected = para_count + img_count
        if non_empty > 0 and total == expected:
            status = "OK"
            detail = f"{para_count} paragraphs + {img_count} images"
        elif non_empty > 0:
            status = "OK"
            detail = f"{non_empty}/{expected} non-empty (expected {para_count}p+{img_count}i)"
        else:
            status = "EMPTY"
            detail = "no content found"
        print(f"  [{status}] '{heading_text}': {detail}")
