"""
模板填写工具 — 基于模板文件生成新文档。

公开 API:
    fill_template(template_path, output_path, content_map, placeholder_pattern=None) -> Path
"""
from .filler import fill_template

__all__ = ['fill_template']
