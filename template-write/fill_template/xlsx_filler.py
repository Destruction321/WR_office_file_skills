"""
# `xlsx` 模板填写器。

- 按单元格替换占位符。处理合并单元格时只替换左上角的值。
- 依赖：openpyxl（缺失时自动安装）
"""

from pathlib import Path
from re import Pattern, Match

from .deps import ensure_import


def fill_xlsx(output_path: Path, content_map: dict[str, str], pattern: Pattern[str]) -> None:
    """
    ## 在 `.xlsx` 文件中替换占位符（原地修改）。

    Args:
        output_path (Path): 已复制的 .xlsx 文件路径。
        content_map (dict[str, str]): 占位符名称到替换文本的映射。
        pattern (Pattern[str]): 占位符正则。
    """
    load_workbook = ensure_import('openpyxl', attr='load_workbook')  # type: ignore[assignment]

    wb = load_workbook(str(output_path))  # type: ignore[operator]

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is None or not isinstance(cell.value, str):
                    continue

                new_val = _replace_all(cell.value, content_map, pattern)
                if new_val == cell.value:
                    continue

                cell.value = new_val

    wb.save(str(output_path))


def _replace_all(text: str, content_map: dict[str, str], pattern: Pattern[str]) -> str:
    """将 text 中所有占位符替换为 content_map 中的值。"""
    def _replacer(match: Match[str]) -> str:
        name = match.group(1)
        return content_map.get(name, match.group(0))
    
    return pattern.sub(_replacer, text)
