"""
# `.pdf` 提取器。

- 使用 `PyMuPDF`（fitz）单库完成文字 + 表格 + 图片提取。
- 文字块按 y 坐标排序，表格按其 y 位置插入到正确阅读位置。
- 落在表格区域内的文字块被剔除，避免内容重复。
- 复杂排版页（流程图、图示等）自动渲染为图片；每页附元数据供 AI 判断。
- 支持 --render-page 按需渲染指定页为图片（第二层 AI 判断）。
"""

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .common import ExtractJob
from .. import assets
from ..deps import ensure_import


@dataclass(frozen=True)
class _TextBlock:
    """
    PDF 文字块快照：(x0, y0, y1, text, font_h)。
    
    Attributes:
        x0 (float): 块左上角 x 坐标。
        y0 (float): 块左上角 y 坐标。
        y1 (float): 块右下角 y 坐标。
        text (str): 块内文字。
        font_h (float): 块内最大字号。
    """
    x0: float
    y0: float
    y1: float
    text: str
    font_h: float


def extract_pdf(job: ExtractJob) -> list[str]:
    """
    ## 通过 `PyMuPDF` 提取 PDF 文本和表格，保留阅读顺序。

    - 每页以 "--- Page N ---" 分隔。
    - 文字块按 y 坐标排序，表格按其 y 位置插入到正确位置。
    - 落在表格 bbox 内的文字块被剔除，避免重复。
    - 复杂排版页自动渲染为图片（第一层兜底）。
    - 每页末尾附元数据 `[meta] blocks=N avg_len=M tables=K`，供 AI 判断是否需要补渲染。
    - `render_pages` 指定的页码（1-based）强制渲染为图片（第二层 AI 按需）。
    - 如有资产提取要求，一并提取 PDF 中的图片。

    Args:
        job (ExtractJob): 提取作业参数。

    Returns:
        lines (list[str]): 提取出的文本行，失败时返回错误信息。
    """
    filepath, assets_dir, render_pages = job.filepath, job.assets_dir, job.render_pages
    lines: list[str] = []
    assets_result: dict[str, list[str]] = {}
    images_dir: Path | None = None
    if assets_dir:
        assets_result = assets.extract_pdf_assets(filepath, assets_dir / filepath.stem)
        images_dir = assets_dir / filepath.stem / 'pages'
        images_dir.mkdir(parents=True, exist_ok=True)

    render_set = set(render_pages or [])

    try:
        fitz_open = ensure_import('PyMuPDF', 'fitz', attr='open')
    except ImportError:
        return ['[Error: PyMuPDF 未安装。执行: pip install PyMuPDF]']

    try:
        doc = fitz_open(filepath)
        for i, page in enumerate(doc):
            page_num = i + 1
            force_render = page_num in render_set
            page_lines, meta = _extract_page(
                page, images_dir=images_dir, page_num=page_num, force_render=force_render,
            )
            if not page_lines and not meta:
                continue

            lines.append(f'--- Page {page_num} ---')
            lines.append('')
            lines.extend(page_lines)
            if meta:
                lines.append(f'[meta] {meta}')
            lines.append('')

        doc.close()

        if assets_result:
            assets.append_assets_summary(lines, assets_result)
        return lines

    except Exception as e:
        return [f'[Error: 读取 PDF 失败: {e}]']


def _extract_page(page,
                  images_dir: Path | None = None,
                  page_num: int = 0,
                  force_render: bool = False) -> tuple[list[str], str]:
    """
    提取单页内容：文字块 + 表格，按 y 坐标排序，剔除表格区域内的文字。

    复杂排版页（碎片多、无表格）自动渲染为图片；`force_render` 可强制渲染。
    """
    # 1. 检测表格及其 bbox
    try:
        tables = page.find_tables()
        table_list = list(tables.tables) if tables.tables else []
    except Exception:
        table_list = []
    table_bboxes = [t.bbox for t in table_list]

    # 2-3. 收集文字块（剔除表格区域内）并合并为段落
    text_blocks = _collect_text_blocks(page, table_bboxes)
    merged = _merge_text_blocks(text_blocks)

    # 4. 计算元数据
    n_blocks = len(merged)
    avg_len = sum(len(b.text) for b in merged) / n_blocks if n_blocks else 0
    short_blocks = sum(1 for b in merged if len(b.text) < 5)
    short_ratio = short_blocks / n_blocks if n_blocks else 0
    n_tables = len(table_list)
    page_width = page.rect.width
    x_spread = (max(b.x0 for b in merged) - min(b.x0 for b in merged)) if n_blocks >= 2 else 0
    x_ratio = x_spread / page_width if page_width > 0 else 0
    meta = (
        f'blocks={n_blocks} avg_len={avg_len:.0f} short={short_blocks} '
        f'tables={n_tables} x_spread={x_ratio:.0%}'
    )

    # 5. 第一层兜底：复杂排版自动渲染为图片
    if force_render or _is_complex_layout(n_blocks, short_ratio, n_tables, x_ratio):
        if images_dir:
            img_name = f'page_{page_num}.png'
            pix = page.get_pixmap(dpi=150)
            pix.save(str(images_dir / img_name))
            rel_hint = f'（图片: {images_dir.name}/{img_name}）'
            return [f'[此页为复杂排版，已渲染为图片{rel_hint}]'], meta
        if force_render:
            return [f'[Error: 需要 --assets-dir 才能渲染页面为图片]'], meta

    # 6-7. 收集内容项并按 y 坐标排序
    items: list[tuple[float, str, Any]] = []
    for b in merged:
        items.append((b.y0, 'text', b.text))
    for t_idx, table in enumerate(table_list):
        items.append((table.bbox[1], 'table', (t_idx, table)))
    items.sort(key=lambda item: item[0])

    # 8. 渲染输出
    page_lines = _render_items(items)
    return page_lines, meta


def _collect_text_blocks(page, table_bboxes: list) -> list[_TextBlock]:
    """
    从页面 dict 中收集文字块，剔除图片块和表格区域内文字。
    """
    text_blocks: list[_TextBlock] = []
    for b in page.get_text("dict", sort=True)["blocks"]:
        if b.get("type", 0) == 1:  # image block
            continue
        
        x0, y0, x1, y1 = b["bbox"]
        text = "".join(
            span["text"]
            for line in b.get("lines", [])
            for span in line.get("spans", [])
        ).strip()
        
        if not text:
            continue
        if _in_table_bbox((x0, y0, x1, y1), table_bboxes):
            continue
        
        font_h = max(
            (span["size"] for line in b.get("lines", []) for span in line.get("spans", [])),
            default=(y1 - y0),
        )
        text_blocks.append(_TextBlock(x0, y0, y1, text, font_h))
    
    return text_blocks


def _merge_text_blocks(text_blocks: list[_TextBlock]) -> list[_TextBlock]:
    """
    合并纵向相邻的文字块为段落（同列、y 间距小于字高倍数）。
    - 动态阈值：缩进容差 2.5 倍字号，行间距 1.5 倍字号。
    """
    text_blocks.sort(key=lambda b: (b.y0, b.x0))
    merged: list[_TextBlock] = []
    for b in text_blocks:
        if merged:
            prev = merged[-1]
            ref_h = max(b.font_h, prev.font_h)
            if abs(b.x0 - prev.x0) < ref_h * 2.5 and 0 < b.y0 - prev.y1 < ref_h * 1.5:
                merged[-1] = _TextBlock(prev.x0, prev.y0, b.y1, prev.text + b.text, prev.font_h)
                continue
        merged.append(b)
    return merged


def _render_items(items: list[tuple[float, str, Any]]) -> list[str]:
    """将内容项（文字 + 表格）渲染为 markdown 文本行。"""
    page_lines: list[str] = []
    for _, kind, payload in items:
        if kind == 'text':
            page_lines.append(payload)
        elif kind == 'table':
            t_idx, table = payload
            rows = table.extract()
            page_lines.append(f'**Table {t_idx + 1}:**')
            page_lines.append('')
            if rows:
                n_cols = len(rows[0])
                header_cells = [
                    str(c).replace('\n', ' ').replace('\r', '').replace('|', '\\|')
                    if c else ''
                    for c in rows[0]
                ]
                page_lines.append('| ' + ' | '.join(header_cells) + ' |')
                page_lines.append('| ' + ' | '.join(['---'] * n_cols) + ' |')
                for row in rows[1:]:
                    cells = [
                        str(c).replace('\n', ' ').replace('\r', '').replace('|', '\\|')
                        if c else ''
                        for c in row
                    ]
                    while len(cells) < n_cols:
                        cells.append('')
                    page_lines.append('| ' + ' | '.join(cells[:n_cols]) + ' |')
        
        page_lines.append('')
    
    return page_lines


def _is_complex_layout(n_blocks: int, short_ratio: float, n_tables: int, x_ratio: float) -> bool:
    """
    保守检测页面是否为复杂排版（流程图、图示等）。

    - 有表格的页面不做图片渲染（保住结构化数据）。
    - 块数多（>10）且短块占比高（>=40%）且 x 坐标离散度大（>50% 页宽）
      -> 文字横向散布 + 碎片化严重，渲染为图片。

    x 坐标离散度用于排除目录/索引页：目录文字集中在窄列内（x_ratio 低），
    流程图文字散布整个页面宽度（x_ratio 高）。
    """
    if n_tables > 0:
        return False
    return n_blocks > 10 and short_ratio >= 0.4 and x_ratio > 0.5


def _in_table_bbox(bbox: tuple, table_bboxes: list) -> bool:
    """检查文字块 bbox 是否落在任一表格 bbox 内（中心点或大面积重叠）。"""
    x0, y0, x1, y1 = bbox[:4]
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    for tx0, ty0, tx1, ty1 in table_bboxes:
        if tx0 <= cx <= tx1 and ty0 <= cy <= ty1:
            return True
        
        ox = max(0.0, min(x1, tx1) - max(x0, tx0))
        oy = max(0.0, min(y1, ty1) - max(y0, ty0))
        block_area = (x1 - x0) * (y1 - y0)
        if block_area > 0 and (ox * oy) / block_area > 0.5:
            return True
    
    return False
