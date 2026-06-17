"""
# 模板填写工具 — 基于模板文件生成新文档。

- 公开 API:
  - `fill_template(...)` — 占位符替换
  - `fill_docx_sections(...)` — 节级内容填充
"""

from .filler import fill_template
from .docx_section_filler import fill_docx_sections

__all__ = ['fill_template', 'fill_docx_sections']
