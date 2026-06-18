"""
# 格式分发器 — 按扩展名派发到各格式模块。

- 公开 API:
  1. extract_file(filepath, assets_dir=None) -> list[str]
  2. EXTRACTORS — 格式扩展名 -> 处理函数映射
"""

from pathlib import Path

from .docx_extractor import extract_doc, extract_docx
from .pdf_extractor import extract_pdf
from .pptx_extractor import extract_ppt, extract_pptx
from .xlsx_extractor import extract_xls, extract_xlsx

# 7 种格式的提取器映射表
EXTRACTORS = {
    '.docx': extract_docx,
    '.doc': extract_doc,
    '.pptx': extract_pptx,
    '.ppt': extract_ppt,
    '.xlsx': extract_xlsx,
    '.xls': extract_xls,
    '.pdf': extract_pdf,
}


def _extract_plain_text(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """
    不认识的后缀当纯文本读，仅用于 DIRECT 模式。

    依次尝试 UTF-8、GBK、GB2312、UTF-16 解码。
    """
    for enc in ('utf-8', 'gbk', 'gb2312', 'utf-16'):
        try:
            with open(filepath, 'r', encoding=enc) as fh:
                return fh.read().splitlines()
        
        except (UnicodeDecodeError, LookupError):
            continue
        except Exception as e:
            return [f'[Error: 读取纯文本失败: {e}]']
    
    return ['[Error: 无法解码此文件（纯文本回退失败）]']


def extract_file(filepath: str, assets_dir: str | None = None) -> list[str]:
    """
    ## 从单个文档文件中提取文本，按扩展名分发。

    Args:
        filepath (str): 文档文件路径。
        assets_dir (str | None): 资源提取目标目录（可选）。

    Returns:
        list[str]: 提取出的文本行。出错时返回含 [Error ...] 的单行列表。
    """
    fp = Path(filepath)
    ad = Path(assets_dir) if assets_dir else None
    handler = EXTRACTORS.get(fp.suffix.lower())
    if handler is None:
        return _extract_plain_text(fp, ad)
    return handler(fp, ad)
