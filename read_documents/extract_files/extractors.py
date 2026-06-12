"""
格式特定的文本提取器和 COM 回退方案，覆盖所有支持的文档类型。

每个 ``_extract_*`` 函数返回 ``list[str]`` —— 提取出的文本行。
"""
from pathlib import Path
from subprocess import run, DEVNULL, TimeoutExpired
from sys import platform, stderr
from typing import Callable

from . import assets
from .util import safe_open_path, mktemp_in_dir
from .section import detect_chinese_heading


def _kill_orphan_com(process_name: str) -> None:
    """
    清理超时后残留的 Office COM 僵尸进程。

    当 ``subprocess.run`` 抛出 ``TimeoutExpired`` 时子进程已被杀掉，
    但 PowerShell 启动的 COM 服务器可能仍残留在系统中。

    此函数使用``taskkill /FI "STATUS eq NOT RESPONDING"`` ，
    只杀死无响应的进程，避免影响用户正在使用的 Office 窗口。
    """
    if platform != 'win32':
        return

    try:
        from subprocess import run as subprocess_run
        # 只杀 "NOT RESPONDING" 的实例——这些是无界面的 COM 僵尸，
        # 不是用户正在使用的 Office 窗口。
        subprocess_run(
            ['taskkill', '/F',
             '/FI', f'IMAGENAME eq {process_name}',
             '/FI', 'STATUS eq NOT RESPONDING'],
            capture_output=True, timeout=5,
        )
    except Exception:
        pass  # 尽力而为，不因清理失败而崩溃


# ===================================================================
#  COM 辅助脚本路径 — 相对于本脚本的父级目录（skills/read_documents/ps1_scripts/）
# ===================================================================

_SCRIPT_DIR = Path(__file__).resolve().parent.parent / 'ps1_scripts'
_PPT_SCRIPT: str = str(_SCRIPT_DIR / 'extract_ppt.ps1')
_DOC_SCRIPT: str = str(_SCRIPT_DIR / 'extract_doc.ps1')
_XLS_SCRIPT: str = str(_SCRIPT_DIR / 'extract_xls.ps1')


# ===================================================================
#  .doc
# ===================================================================

def _extract_doc(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """先尝试 python-docx（可处理伪装成 .doc 的 .docx），再 COM 回退。"""
    try:
        result = _extract_docx(filepath, assets_dir)
        if all(not line or line.startswith('[') for line in result):
            raise ValueError('python-docx 仅返回了错误信息')
        
        return result
    
    except Exception:
        return _extract_doc_com(filepath, assets_dir)

def _extract_doc_com(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """通过 Windows COM 提取旧 .doc 文件，失败时返回错误信息列表。"""
    if platform != 'win32':
        return ['[Error: 旧格式 .doc 提取需要 Windows + Microsoft Office]']
    
    if not Path(_DOC_SCRIPT).exists():
        return ['[Error: 找不到 Word 提取脚本]']

    tmp_out = mktemp_in_dir(filepath, prefix='tmp_doc_') / 'output.txt'
    try:
        run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-File', _DOC_SCRIPT,
             '-DocPath', str(filepath), '-OutFile', str(tmp_out)],
            stdout=DEVNULL, stderr=DEVNULL, timeout=120
        )
        if tmp_out.exists():
            lines = tmp_out.read_text(encoding='utf-8-sig').splitlines()
            if assets_dir:
                lines.append('[提示: 旧格式 .doc 暂不支持嵌入文件提取]')
            return lines
        return ['[Error: Word 提取未产生输出]']
    
    except TimeoutExpired:
        _kill_orphan_com('WINWORD.EXE')
        return ['[Error: Word 提取超时]']
    
    except Exception as e:
        return [f'[Error: 通过 COM 提取 DOC 失败: {e}]']


# ===================================================================
#  .docx
# ===================================================================

def _extract_docx(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """通过 python-docx 提取 .docx 文件，失败时返回错误信息列表。"""
    lines: list[str] = []
    assets_result: dict[str, list[str]] = {}

    if assets_dir:
        assets_result = assets.extract_ooxml_assets(filepath, assets_dir / filepath.stem, '.docx')

    try:
        from docx import Document
    
    except ImportError:
        return ['[Error: python-docx 未安装。执行: pip install python-docx]']

    with safe_open_path(filepath) as safe_path:
        try:
            doc = Document(str(safe_path))
        except Exception as e:
            return [f'[Error: 用 python-docx 打开 .docx 失败: {e}]']

        # 始终使用有序提取——保留段落/表格的真实交错顺序，
        # 并添加标题标记以便小节过滤。
        _extract_docx_body_ordered(doc, lines)

    if assets_result:
        assets.append_assets_summary(lines, assets_result)

    return lines

def _extract_docx_body_ordered(doc, lines: list[str]) -> None:
    """
    按文档顺序提取 docx 内容，添加 Markdown 标题标记。

    通过迭代 XML body 使段落和表格以真实文档顺序出现
    （有别于``doc.paragraphs`` + ``doc.tables`` 这种分离序列的方式）。
    """
    from docx.oxml.ns import qn

    def _text(elem) -> str:
        """收集 elem 内所有 w:t 节点文本，拼接后返回。"""
        return ''.join(t.text or '' for t in elem.iter(qn('w:t')))

    body = doc.element.body
    for child in body:
        if child.tag == qn('w:p'):
            text = _text(child).strip()
            if not text:
                continue

            heading_level = _get_heading_style_level(child)
            if heading_level is None:
                heading_level = detect_chinese_heading(text)

            if heading_level is not None:
                lines.append(f'{"#" * min(heading_level, 6)} {text}')
            else:
                lines.append(text)

        elif child.tag == qn('w:tbl'):
            lines.append('')
            for row_elem in child.iter(qn('w:tr')):
                cells = [
                    _text(tc).strip().replace('\n', ' ').replace('\r', '')
                    for tc in row_elem.iter(qn('w:tc'))
                ]
                if any(c for c in cells):
                    lines.append(' | '.join(cells))
            
            lines.append('')

def _get_heading_style_level(child) -> int | None:
    """检查段落是否应用了 Word 标题样式，是则返回标题级别，否则返回 None。"""
    from docx.oxml.ns import qn
    pPr = child.find(qn('w:pPr'))
    if pPr is None:
        return None

    pStyle = pPr.find(qn('w:pStyle'))
    if pStyle is None:
        return None

    style_val = pStyle.get(qn('w:val'), '')
    if not style_val.lower().startswith('heading'):
        return None

    try:
        return int(style_val.split()[-1])
    
    except ValueError:
        return 1  # 默认标题级别


# ===================================================================
#  .pptx
# ===================================================================

def _extract_pptx(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """通过 python-pptx 提取 .pptx 文件，失败时返回错误信息列表。"""
    assets_result: dict[str, list[str]] = {}

    if assets_dir:
        assets_result = assets.extract_ooxml_assets(filepath, assets_dir / filepath.stem, '.pptx')

    try:
        from pptx import Presentation

    except ImportError:
        return ['[Error: python-pptx 未安装。执行: pip install python-pptx]']

    with safe_open_path(filepath) as safe_path:
        try:
            prs = Presentation(str(safe_path))

        except Exception:
            return ['[Error: 用 python-pptx 打开 .pptx 失败]']

        lines: list[str] = []
        for i, slide in enumerate(prs.slides, 1):
            lines.append(f'--- Slide {i} ---')
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for p in shape.text_frame.paragraphs:  # type: ignore[attr-defined]
                        t = p.text.strip()
                        if not t:
                            continue

                        lines.append(t)

                if shape.has_table:
                    for row in shape.table.rows:  # type: ignore[attr-defined]
                        cells = [
                            cell.text.strip().replace('\n', ' ').replace('\r', '') for cell in row.cells
                        ]
                        lines.append(' | '.join(cells))

    if assets_result:
        assets.append_assets_summary(lines, assets_result)

    return lines

# ===================================================================
#  .ppt
# ===================================================================

def _extract_ppt(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """先尝试 python-pptx，失败时回退 COM（旧格式兼容）。"""
    try:
        result = _extract_pptx(filepath, assets_dir)
        if all(not line or line.startswith('[') for line in result):
            raise ValueError('python-pptx 仅返回了错误信息')
        
        return result
    
    except Exception:
        return _extract_ppt_com(filepath, assets_dir)

def _extract_ppt_com(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    if platform != 'win32':
        return ['[Error: 旧格式 .ppt 提取需要 Windows + Microsoft Office]']
    
    if not Path(_PPT_SCRIPT).exists():
        return ['[Error: 找不到 PowerPoint 提取脚本]']

    tmp_out = mktemp_in_dir(filepath, prefix='tmp_ppt_') / 'output.txt'
    try:
        run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-File', _PPT_SCRIPT,
             '-PptPath', str(filepath), '-OutFile', str(tmp_out)],
            stdout=DEVNULL, stderr=DEVNULL, timeout=120
        )
        if tmp_out.exists():
            lines = tmp_out.read_text(encoding='utf-8-sig').splitlines()
            if assets_dir:
                lines.append('[提示: 旧格式 .ppt 暂不支持嵌入文件提取]')
            
            return lines
        return ['[Error: PowerPoint 提取未产生输出]']
    
    except TimeoutExpired:
        _kill_orphan_com('POWERPNT.EXE')
        return ['[Error: PowerPoint 提取超时]']
    
    except Exception as e:
        return [f'[Error: 通过 COM 提取 PPT 失败: {e}]']


# ===================================================================
#  .xlsx
# ===================================================================

def _extract_xlsx(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """通过 python-xlsx 提取 .xlsx 文件，失败时返回错误信息列表。"""
    lines: list[str] = []
    assets_result: dict[str, list[str]] = {}

    if assets_dir:
        assets_result = assets.extract_ooxml_assets(
            filepath, assets_dir / filepath.stem, filepath.suffix.lower()
        )

    try:
        from openpyxl import load_workbook
        with safe_open_path(filepath) as safe_path:
            wb = load_workbook(safe_path, data_only=True)
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


# ===================================================================
#  .xls
# ===================================================================

def _extract_xls(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """通过 xlrd 提取旧 .xls（BIFF）文件，再 COM 回退。"""
    lines: list[str] = []
    assets_result: dict[str, list[str]] = {}

    if assets_dir:
        assets_result = assets.extract_ooxml_assets(filepath, assets_dir / filepath.stem, '.xls')

    try:
        from xlrd import open_workbook
        with safe_open_path(filepath) as safe_path:
            wb = open_workbook(str(safe_path))
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
    
    except ImportError:
        print('  [警告] xlrd 未安装，尝试 COM 回退 ...', file=stderr)
    
    except Exception as e:
        print(f'  [警告] xlrd 失败: {e}，尝试 COM 回退 ...', file=stderr)

    return _extract_xls_com(filepath, assets_dir)

def _extract_xls_com(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    if platform != 'win32':
        return ['[Error: 旧格式 .xls 提取需要 Windows + Microsoft Office]']
    if not Path(_XLS_SCRIPT).exists():
        return ['[Error: 找不到 Excel 提取脚本。请先安装 xlrd: pip install xlrd]']

    tmp_out = mktemp_in_dir(filepath, prefix='tmp_xls_') / 'output.txt'
    try:
        run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-File', _XLS_SCRIPT,
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
        _kill_orphan_com('EXCEL.EXE')
        return ['[Error: Excel 提取超时]']
    
    except Exception as e:
        return [f'[Error: 通过 COM 提取 XLS 失败: {e}]']


# ===================================================================
#  .pdf
# ===================================================================

def _extract_pdf(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    lines: list[str] = []
    assets_result: dict[str, list[str]] = {}

    if assets_dir:
        assets_result = assets.extract_pdf_assets(filepath, assets_dir / filepath.stem)

    try:
        from pdfplumber import open as pdfplumber_open
        with safe_open_path(filepath) as safe_path:
            with pdfplumber_open(safe_path) as pdf:
                for i, page in enumerate(pdf.pages, 1):
                    text = page.extract_text()
                    if not text:
                        continue

                    lines.append(f'--- Page {i} ---')
                    lines.append(text)

        if assets_result:
            assets.append_assets_summary(lines, assets_result)

        return lines

    except ImportError:
        print('  [警告] pdfplumber 未安装，尝试 PyPDF2 ...', file=stderr)
    
    except Exception as e:
        print(f'  [警告] pdfplumber 失败: {e}，尝试 PyPDF2 ...', file=stderr)

    try:
        from PyPDF2 import PdfReader
        with safe_open_path(filepath) as safe_path:
            reader = PdfReader(safe_path)
            for i, page in enumerate(reader.pages, 1):
                text = page.extract_text()
                if not text:
                    continue

                lines.append(f'--- Page {i} ---')
                lines.append(text)

        if assets_result:
            assets.append_assets_summary(lines, assets_result)

        return lines

    except ImportError:
        return ['[Error: pdfplumber 和 PyPDF2 均未安装。执行: pip install pdfplumber]']
    
    except Exception as e:
        return [f'[Error: 读取 PDF 失败: {e}]']


EXTRACTORS: dict[str, Callable[[Path, Path | None], list[str]]] = {
    '.docx': _extract_docx,
    '.doc':  _extract_doc,
    '.pptx': _extract_pptx,
    '.ppt':  _extract_ppt,
    '.xlsx': _extract_xlsx,
    '.xls':  _extract_xls,
    '.pdf':  _extract_pdf,
}


def _extract_plain_text(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """不认识的后缀当纯文本读，仅用于 DIRECT 模式。"""
    for enc in ('utf-8', 'gbk', 'gb2312', 'utf-16'):
        try:
            with safe_open_path(filepath) as safe_path:
                with open(safe_path, 'r', encoding=enc) as fh:
                    return fh.read().splitlines()
        
        except (UnicodeDecodeError, LookupError):
            continue
        
        except Exception as e:
            return [f'[Error: 读取纯文本失败: {e}]']
    
    return ['[Error: 无法解码此文件（纯文本回退失败）]']


def extract_file(filepath: str, assets_dir: str | None = None) -> list[str]:
    """
    从单个文档文件中提取文本，按扩展名分发。

    Args:
        filepath (str): 文档文件路径。
        assets_dir (str | None): 资源提取目标目录（可选）。

    Returns:
        list[str]: 提取出的文本行。出错时返回含 ``[Error ...]`` 的单行列表。
    """
    fp = Path(filepath)
    ad = Path(assets_dir) if assets_dir else None
    handler = EXTRACTORS.get(fp.suffix.lower())
    if handler is None:
        # DIRECT 模式可处理的任意文件：当纯文本读
        return _extract_plain_text(fp, ad)

    return handler(fp, ad)
