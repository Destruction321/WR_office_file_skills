"""
# 老格式 -> 新格式转换（pywin32 COM SaveAs，仅 Windows）。

在边界处将旧二进制格式（`.doc`/`.ppt`/`.xls`）转换为现代格式
（`.docx`/`.pptx`/`.xlsx`），下游统一走 python-docx / python-pptx /
openpyxl 提取路径，不再手写 COM 文档遍历。

- 原始文件以只读方式打开，永不被修改。
- 转换产物放在系统临时目录，调用方通过 cleanup 回调清理。
- **COM 调用隔离在子进程（conv_worker.py）中执行**，父进程带超时等待；
  Office 卡死（弹窗/无响应）时子进程被终止、僵尸进程被清理，主流程不会被阻塞。
"""

from pathlib import Path
from shutil import rmtree
from subprocess import DEVNULL, TimeoutExpired, run
from sys import executable, platform
from tempfile import mkdtemp
from typing import Callable

from .extractors.common import kill_orphan_com

# 源后缀 -> (COM ProgID, SaveAs 方法, 格式常量)
_APPS: dict[str, tuple[str, str, int]] = {
    '.doc': ('Word.Application', 'SaveAs2', 12),        # wdFormatXMLDocument
    '.ppt': ('PowerPoint.Application', 'SaveAs', 24),   # ppSaveAsOpenXMLPresentation
    '.xls': ('Excel.Application', 'SaveAs', 51),        # xlOpenXMLWorkbook
}

_MODERN: dict[str, str] = {'.doc': '.docx', '.ppt': '.pptx', '.xls': '.xlsx'}

_PROCESS: dict[str, str] = {
    'Word.Application': 'WINWORD.EXE',
    'PowerPoint.Application': 'POWERPNT.EXE',
    'Excel.Application': 'EXCEL.EXE',
}

# 转换超时（秒）—— Office 弹窗/无响应时防止主流程永久阻塞
_TIMEOUT = 120


def _convert_via_com(src: Path, dst: Path, progid: str, save_method: str, fmt: int) -> None:
    """调用 Office COM 将 src 转换为 dst（源文件只读打开）。仅在子进程中执行。"""
    if platform != 'win32':
        raise RuntimeError(f'{src.suffix} 转换需要 Windows + Microsoft Office')

    from .deps import ensure_import
    Dispatch = ensure_import('pywin32', 'win32com.client', attr='Dispatch')
    app = None
    try:
        app = Dispatch(progid)
        app.Visible = False
        try:
            app.DisplayAlerts = 0
        except Exception:
            pass

        if progid == 'Word.Application':
            doc = app.Documents.Open(
                str(src), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False,
            )
            getattr(doc, save_method)(str(dst), fmt)
            doc.Close(False)

        elif progid == 'PowerPoint.Application':
            # PowerPoint 自动化需要可见窗口才能 Open 演示文稿
            try:
                app.Visible = True
            except Exception:
                pass
            pres = app.Presentations.Open(str(src), True, False, False)
            getattr(pres, save_method)(str(dst), fmt)
            pres.Close()

        else:  # Excel.Application
            wb = app.Workbooks.Open(str(src), 0, True)
            getattr(wb, save_method)(str(dst), fmt)
            wb.Close(False)

    except Exception as e:
        raise RuntimeError(f'{src.suffix} 转换失败: {e}') from e

    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                kill_orphan_com(_PROCESS[progid])


def _run_worker(src: Path, dst: Path) -> None:
    """
    ## 在子进程中执行 COM 转换，带超时与僵尸清理。

    Args:
        src (Path): 源文件路径。
        dst (Path): 目标文件路径。

    Raises:
        RuntimeError: 非 Windows、超时或转换未产生输出。
    """
    if platform != 'win32':
        raise RuntimeError(f'{src.suffix} 转换需要 Windows + Microsoft Office')
    progid, _, _ = _APPS[src.suffix.lower()]
    skill_root = Path(__file__).resolve().parent.parent
    try:
        run(
            [executable, '-m', 'extract_files.conv_worker', str(src), str(dst)],
            stdout=DEVNULL, stderr=DEVNULL, timeout=_TIMEOUT, cwd=str(skill_root),
        )
    except TimeoutExpired:
        kill_orphan_com(_PROCESS[progid])
        raise RuntimeError(
            f'{src.suffix} 转换超时（{_PROCESS[progid]} 可能弹出对话框或无响应），已终止'
        ) from None
    if not dst.exists():
        raise RuntimeError(f'{src.suffix} 转换未产生输出（Office 不可用或文件损坏）')


def convert_to_modern(src: Path) -> tuple[Path, Callable[[], None]]:
    """
    ## 将老格式文件转换为系统临时目录中的现代格式副本。

    转换产物文件名保留原 stem（如 `报告.doc` -> `报告.docx`），
    便于下游提取器按原文件名组织资源目录。

    Args:
        src (Path): 老格式文件路径（`.doc`/`.ppt`/`.xls`）。

    Returns:
        (modern_path, cleanup) (tuple[Path, Callable[[], None]]):
            转换后的新格式路径 + 无参清理回调。

    Raises:
        RuntimeError: 不支持的格式 / 非 Windows / 转换失败（临时目录已清理）。
    """
    ext = src.suffix.lower()
    if ext not in _APPS:
        raise RuntimeError(f'不支持的格式: {ext}')

    workdir = Path(mkdtemp(prefix='conv_'))
    modern = workdir / (src.stem + _MODERN[ext])
    try:
        _run_worker(src, modern)
    except Exception:
        rmtree(workdir, ignore_errors=True)
        raise

    return modern, lambda: rmtree(workdir, ignore_errors=True)
