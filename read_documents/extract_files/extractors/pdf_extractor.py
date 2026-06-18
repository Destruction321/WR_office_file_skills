"""
# `.pdf` 提取器。

- 首选 `pdfplumber`（文字提取质量更高），
- 回退 `PyPDF2`（轻量但提取质量一般）。
"""

from pathlib import Path
from sys import stderr

from .. import assets
from ..deps import ensure_import


def extract_pdf(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """
    ## 提取 PDF 文本，尝试 `pdfplumber` -> `PyPDF2` 逐级回退。

    - 每页以 "--- Page N ---" 分隔。
    - 如有资产提取要求，一并提取 PDF 中的图片。

    Args:
        filepath (Path): 文档文件路径。
        assets_dir (Path | None): 资源提取目标目录（可选）。

    Returns:
        lines (list[str]): 提取出的文本行，失败时返回错误信息。
    """
    lines: list[str] = []
    assets_result: dict[str, list[str]] = {}
    if assets_dir:
        assets_result = assets.extract_pdf_assets(filepath, assets_dir / filepath.stem)

    # 先试 pdfplumber（文字提取效果好）
    try:
        pdfplumber_open = ensure_import('pdfplumber', attr='open')  # type: ignore[assignment]
    except ImportError:
        pdfplumber_open = None

    if pdfplumber_open:
        try:
            with pdfplumber_open(filepath) as pdf:  # type: ignore[operator]
                for i, page in enumerate(pdf.pages, 1):
                    text = page.extract_text()
                    if not text:
                        continue
                    lines.append(f'--- Page {i} ---')
                    lines.append(text)

            if assets_result:
                assets.append_assets_summary(lines, assets_result)
            if lines:
                return lines
        
        except Exception as e:
            print(f'  [警告] pdfplumber 失败: {e}，尝试 PyPDF2 ...', file=stderr)

    # 回退 PyPDF2
    try:
        PdfReader = ensure_import('PyPDF2', attr='PdfReader')  # type: ignore[assignment]
    except ImportError:
        return ['[Error: pdfplumber 和 PyPDF2 均未安装。执行: pip install pdfplumber]']

    try:
        reader = PdfReader(filepath)  # type: ignore[operator]
        for i, page in enumerate(reader.pages, 1):
            text = page.extract_text()
            if not text:
                continue
            lines.append(f'--- Page {i} ---')
            lines.append(text)

        if assets_result:
            assets.append_assets_summary(lines, assets_result)
        return lines
    
    except Exception as e:
        return [f'[Error: 读取 PDF 失败: {e}]']
