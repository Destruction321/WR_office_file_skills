"""
# `.xls` / `.xlsx` 提取器。
- 先尝试 `openpyxl` / `xlrd`（xlrd 可直接读真二进制 `.xls`），失败时通过
  pywin32 COM SaveAs 将 `.xls` 转换为 `.xlsx` 后复用 openpyxl 提取路径。
- 转换在系统临时目录进行，原始 `.xls` 只读打开、永不被修改。
"""

from sys import stderr

from .common import ExtractJob
from .. import assets
from ..deps import ensure_import


def extract_xlsx(job: ExtractJob) -> list[str]:
    """
    ## 通过 `openpyxl` 提取 `.xlsx` 文件，逐工作表按行提取。

    Args:
        job (ExtractJob): 待提取作业，包含文件路径和资源目录。

    Returns:
        lines (list[str]): 提取出的文本行，失败时返回错误信息。
    """
    filepath, assets_dir = job.filepath, job.assets_dir
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


def extract_xls(job: ExtractJob) -> list[str]:
    """
    ## 通过 `xlrd` 提取旧 `.xls`（BIFF）文件，失败时 COM 转换回退。
    - `xlrd` 能处理大部分 `.xls`；失败时通过 COM SaveAs 转换为 `.xlsx`
      后复用 openpyxl 提取路径。

    Args:
        job (ExtractJob): 待提取作业，包含文件路径和资源目录。

    Returns:
        lines (list[str]): 提取出的文本行，失败时返回错误信息。
    """
    filepath, assets_dir = job.filepath, job.assets_dir
    lines: list[str] = []
    assets_result: dict[str, list[str]] = {}
    if assets_dir:
        assets_result = assets.extract_ooxml_assets(filepath, assets_dir / filepath.stem, '.xls')

    try:
        open_workbook = ensure_import('xlrd', attr='open_workbook')
    except ImportError:
        print('  [警告] xlrd 未安装，尝试 COM 转换回退 ...', file=stderr)
        return _extract_xls_converted(job)

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
        print(f'  [警告] xlrd 失败: {e}，尝试 COM 转换回退 ...', file=stderr)

    return _extract_xls_converted(job)


def _extract_xls_converted(job: ExtractJob) -> list[str]:
    """通过 COM 转换 .xls -> .xlsx 后，复用 openpyxl 提取（含资源提取）。"""
    from ..convert import convert_to_modern
    try:
        modern, cleanup = convert_to_modern(job.filepath)
    except RuntimeError as e:
        return [f'[Error: {e}]']
    try:
        lines = extract_xlsx(ExtractJob(filepath=modern, assets_dir=job.assets_dir))
        if not lines or all(not line or line.startswith('[') for line in lines):
            return lines
        lines.append('[提示: 已通过 Excel COM 转换为 .xlsx 后提取]')
        return lines
    finally:
        cleanup()
