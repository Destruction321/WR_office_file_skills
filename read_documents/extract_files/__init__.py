"""
文档文本 + 资源提取器 — 中文路径安全。

公开 API:
    extract_file(filepath, assets_dir=None, section=None) -> list[str]
    normalize_path(path) -> str
    EXTRACTORS  — 格式扩展名 -> 处理函数映射
"""
from .extractors import extract_file, EXTRACTORS
from .discovery import normalize_path

__all__ = ['extract_file', 'normalize_path', 'EXTRACTORS']
