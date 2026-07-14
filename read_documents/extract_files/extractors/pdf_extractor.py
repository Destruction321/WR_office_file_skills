"""
# `.pdf` 提取器。

- 使用 `PyMuPDF`（fitz）单库完成文字 + 表格 + 图片提取。
- 文字块按 y 坐标排序，表格按其 y 位置插入到正确阅读位置。
- 落在表格区域内的文字块被剔除，避免内容重复。
"""

from pathlib import Path
from typing import Any

from .. import assets
from ..deps import ensure_import


def extract_pdf(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """
    ## 通过 `PyMuPDF` 提取 PDF 文本和表格，保留阅读顺序。

    - 每页以 "--- Page N ---" 分隔。
    - 文字块按 y 坐标排序，表格按其 y 位置插入到正确位置。
    - 落在表格 bbox 内的文字块被剔除，避免重复。
    - 如有资产提取要求，一并提取 PDF 中的图片。

    Args:
        filepath (Path): PDF 文件路径。
        assets_dir (Path | None): 资源提取目标目录（可选）。

    Returns:
        lines (list[str]): 提取出的文本行，失败时返回错误信息。
    """
    lines: list[str] = []
    assets_result: dict[str, list[str]] = {}
    if assets_dir:
        assets_result = assets.extract_pdf_assets(filepath, assets_dir / filepath.stem)

    try:
        fitz_open = ensure_import('PyMuPDF', 'fitz', attr='open')
    except ImportError:
        return ['[Error: PyMuPDF 未安装。执行: pip install PyMuPDF]']

    try:
        doc = fitz_open(filepath)
        for i, page in enumerate(doc):
            page_lines = _extract_page(page)
            if not page_lines:
                continue

            lines.append(f'--- Page {i+1} ---')
            lines.append('')
            lines.extend(page_lines)
            lines.append('')

        doc.close()

        if assets_result:
            assets.append_assets_summary(lines, assets_result)
        return lines

    except Exception as e:
        return [f'[Error: 读取 PDF 失败: {e}]']


def _extract_page(page) -> list[str]:
    """提取单页内容：文字块 + 表格，按 y 坐标排序，剔除表格区域内的文字。"""
    # 1. 检测表格及其 bbox
    try:
        tables = page.find_tables()
        table_list = list(tables.tables) if tables.tables else []
    except Exception:
        table_list = []

    table_bboxes = [t.bbox for t in table_list]

    # 2. 获取文字块（带坐标），剔除表格区域内的文字块
    text_blocks: list[tuple[float, float, float, str]] = []  # (x0, y0, y1, text)
    for b in page.get_text("blocks"):
        x0, y0, x1, y1, text, _, block_type = b
        if block_type == 1:  # image block
            continue
        text = text.strip()
        if not text:
            continue
        if _in_table_bbox((x0, y0, x1, y1), table_bboxes):
            continue
        text_blocks.append((x0, y0, y1, text))

    # 3. 合并纵向相邻的文字块为段落（同列、y 间距小）
    text_blocks.sort(key=lambda b: (b[1], b[0]))
    merged: list[tuple[float, float, float, str]] = []  # (y0, x0, y1, text)
    for x0, y0, y1, text in text_blocks:
        if merged:
            px0, _, py1, _ = merged[-1]
            gap = y0 - py1
            same_col = abs(x0 - px0) < 30  # 容忍首行缩进
            if same_col and 0 < gap < 12:
                _merge_into_last(merged, text, y1)
                continue
        merged.append((y0, x0, y1, text))

    # 4. 收集所有内容项 (y0, kind, payload)
    items: list[tuple[float, str, Any]] = []
    for y0, _, _, text in merged:
        items.append((y0, 'text', text))
    for t_idx, table in enumerate(table_list):
        items.append((table.bbox[1], 'table', (t_idx, table)))

    # 5. 按 y 坐标排序
    items.sort(key=lambda item: item[0])

    # 6. 渲染输出
    page_lines: list[str] = []
    for _, kind, payload in items:
        if kind == 'text':
            page_lines.append(payload)
            page_lines.append('')
        elif kind == 'table':
            t_idx, table = payload
            rows = table.extract()
            page_lines.append(f'**Table {t_idx + 1}:**')
            page_lines.append('')
            for row in rows:
                cells = [
                    str(c).replace('\n', ' ').replace('\r', '').replace('|', '\\|')
                    if c else ''
                    for c in row
                ]
                page_lines.append('| ' + ' | '.join(cells) + ' |')
            page_lines.append('')

    return page_lines


def _merge_into_last(merged: list, text: str, y1: float) -> None:
    """将 text 合并到 merged 最后一项，更新其 y1 和 text。"""
    y0, x0, _, old_text = merged[-1]
    merged[-1] = (y0, x0, y1, old_text + text)


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
