"""
# 文档文本 + 资源提取器 — 中文路径安全。

- 公开 API:
  1. extract_file(filepath, assets_dir=None, section=None) -> list[str]
  2. normalize_path(path) -> str
  3. EXTRACTORS  — 格式扩展名 -> 处理函数映射
"""

from .discovery import normalize_path
from .extractors import extract_file, EXTRACTORS

__all__ = ['extract_file', 'normalize_path', 'EXTRACTORS']
