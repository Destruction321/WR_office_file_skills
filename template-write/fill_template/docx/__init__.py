"""
# docx 子包 — Word 文档操作。

公开 API:
  - `scan_docx` — 扫描模板结构
  - `fill_docx_sections` — 节级内容填充
  - `fill_docx` — 占位符替换
"""
from .section_filler import scan_docx, fill_docx_sections
from .placeholder import fill_docx

__all__ = ['scan_docx', 'fill_docx_sections', 'fill_docx']
