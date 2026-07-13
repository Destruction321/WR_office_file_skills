"""
# `.xls` / `.xlsx` 提取器。
- 先尝试 `openpyxl` / `xlrd`，再 COM 回退处理旧 `.xls` 格式。
"""

from pathlib import Path
from subprocess import run, DEVNULL, TimeoutExpired
from sys import platform, stderr

from .common import kill_orphan_com, XLS_SCRIPT
from .. import assets
from ..deps import ensure_import
from ..util import mktemp_in_dir


def extract_xlsx(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """
    ## 通过 `openpyxl` 提取 `.xlsx` 文件，逐工作表按行提取。

    Args:
        filepath (Path): 文档文件路径。
        assets_dir (Path | None): 资源提取目标目录（可选）。

    Returns:
        lines (list[str]): 提取出的文本行，失败时返回错误信息。
    """
    lines: list[str] = []
    assets_result: dict[str, list[str]] = {}
    if assets_dir:
        assets_result = assets.extract_ooxml_assets(
            filepath, assets_dir / filepath.stem, filepath.suffix.lower()
        )

    try:
        load_workbook = ensure_import('openpyxl', attr='load_workbook')
    except ImportError:
        return ['[Error: openpyxl 未安装。执行: pip install openpyxl]']

    try:
        wb = load_workbook(filepath, data_only=True)
        for name in wb.sheetnames:
            ws = wb[name]
            max_row = ws.max_row or 0
            max_col = ws.max_column or 0
            lines.append(f'--- Sheet: {name} ({max_row} 行 x {max_col} 列) ---')
            for row in ws.iter_rows(min_row=1, max_row=max_row or None, values_only=True):
                cells = [
                    str(c).replace('\n', ' ').replace('\r', '')
                    if c is not None else '' for c in row
                ]
                lines.append(' | '.join(cells))

        if assets_result:
            assets.append_assets_summary(lines, assets_result)
        return lines
    
    except MemoryError:
        return ['[Error: 文件过大无法全部加载到内存。请尝试分段处理或使用只读模式。]']
    except Exception as e:
        return [f'[Error: 读取 XLSX 失败: {e}]']


def extract_xls(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """
    ## 通过 `xlrd` 提取旧 `.xls`（BIFF）文件，再 COM 回退。
    - `xlrd` 能处理大部分 `.xls`，失败时自动走 COM 路径。

    Args:
        filepath (Path): 文档文件路径。
        assets_dir (Path | None): 资源提取目标目录（可选）。

    Returns:
        lines (list[str]): 提取出的文本行，失败时返回错误信息。
    """
    lines: list[str] = []
    assets_result: dict[str, list[str]] = {}
    if assets_dir:
        assets_result = assets.extract_ooxml_assets(filepath, assets_dir / filepath.stem, '.xls')

    try:
        open_workbook = ensure_import('xlrd', attr='open_workbook')
    except ImportError:
        print('  [警告] xlrd 未安装，尝试 COM 回退 ...', file=stderr)
        return _extract_xls_com(filepath, assets_dir)

    try:
        wb = open_workbook(str(filepath))
        for s in range(wb.nsheets):
            sheet = wb.sheet_by_index(s)
            lines.append(f'--- Sheet: {sheet.name} ({sheet.nrows} 行 x {sheet.ncols} 列) ---')
            for row_idx in range(sheet.nrows):
                cells = [
                    sheet.cell_value(row_idx, col_idx).replace('\n', ' ').replace('\r', '')
                    for col_idx in range(sheet.ncols)
                ]
                lines.append(' | '.join(cells))

        if assets_result:
            assets.append_assets_summary(lines, assets_result)
        return lines
    
    except Exception as e:
        print(f'  [警告] xlrd 失败: {e}，尝试 COM 回退 ...', file=stderr)

    return _extract_xls_com(filepath, assets_dir)


def _extract_xls_com(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """通过 Windows COM 提取旧 .xls 文件。"""
    if platform != 'win32':
        return ['[Error: 旧格式 .xls 提取需要 Windows + Microsoft Office]']
    if not XLS_SCRIPT.exists():
        return ['[Error: 找不到 Excel 提取脚本。请先安装 xlrd: pip install xlrd]']

    tmp_out = mktemp_in_dir(filepath, prefix='tmp_xls_') / 'output.txt'
    try:
        run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-File', str(XLS_SCRIPT),
             '-XlsPath', str(filepath), '-OutFile', str(tmp_out)],
            stdout=DEVNULL, stderr=DEVNULL, timeout=120
        )
        if tmp_out.exists():
            lines = tmp_out.read_text(encoding='utf-8-sig').splitlines()
            if assets_dir:
                lines.append('[提示: 旧格式 .xls 暂不支持嵌入文件提取]')
            return lines
        return ['[Error: Excel 提取未产生输出]']

    except TimeoutExpired:
        kill_orphan_com('EXCEL.EXE')
        return ['[Error: Excel 提取超时]']

    except Exception as e:
        return [f'[Error: 通过 COM 提取 XLS 失败: {e}]']
