"""
# `.doc` 转换适配层 —— 通过 Word COM（pywin32，仅 Windows）。

- 将旧格式 `.doc` 转换为 `.docx`，使 fill_template 现有的 `.docx` 流程可直接复用。
- 遵循「先复制再填写」原则：原始 `.doc` 永不被修改。
- Windows 专用：需要 Microsoft Word。非 Windows 环境调用时直接抛 RuntimeError。
- **COM 调用隔离在子进程（conv_worker.py）中执行**，父进程带超时等待；
  Word 卡死（弹窗/无响应）时子进程被终止、僵尸进程被清理，填写流程不会被阻塞。
"""

from pathlib import Path
from shutil import rmtree
from subprocess import DEVNULL, TimeoutExpired, run
from sys import executable, platform
from tempfile import mkdtemp
from time import sleep
from typing import Callable

# 转换超时（秒）—— Word 弹窗/无响应时防止主流程永久阻塞
_TIMEOUT = 120


def rmtree_retry(path: Path, attempts: int = 10, delay: float = 0.3) -> None:
    """
    ## 递归删除目录，带重试（Windows 上索引器/杀软会瞬时占用文件句柄）。

    Args:
        path (Path): 待删除目录。
        attempts (int): 重试次数，默认 10。
        delay (float): 每次重试间隔秒数，默认 0.3。
    """
    for _ in range(attempts):
        try:
            rmtree(path)
            return
        except OSError:
            sleep(delay)
    rmtree(path, ignore_errors=True)


def _kill_word_orphans() -> None:
    """清理转换超时/失败后残留的无响应 Word 进程（不影响用户正常使用的 Word 窗口）。"""
    if platform != 'win32':
        return
    try:
        run(
            [
                'taskkill', '/F',
                '/FI', 'IMAGENAME eq WINWORD.EXE',
                '/FI', 'STATUS eq NOT RESPONDING'
            ],
            capture_output=True, timeout=5,
        )
    except Exception:
        pass  # 尽力而为，不因清理失败而崩溃


def _convert_via_com(src: Path, dst: Path, to_format: str) -> None:
    """通过 Word COM 执行格式转换（.doc <-> .docx）。仅在子进程中执行。"""
    if platform != 'win32':
        raise RuntimeError('.doc 转换需要 Windows + Microsoft Word')

    from .deps import ensure_import
    fmt = 12 if to_format == 'docx' else 0
    word = None
    try:
        Dispatch = ensure_import('pywin32', 'win32com.client', attr='Dispatch')
        word = Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        # 只读打开，禁止转换确认 / 最近文件列表，避免弹窗阻塞
        doc = word.Documents.Open(
            str(src), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False,
        )
        doc.SaveAs2(str(dst), fmt)
        doc.Close(False)
    except Exception as e:
        raise RuntimeError(f'Word 转换失败: {e}') from e
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                _kill_word_orphans()


def _convert(src: Path, dst: Path, to_format: str) -> None:
    """
    ## 在子进程中执行 Word 转换，带超时与僵尸清理。

    Args:
        src (Path): 源文件路径（以只读方式打开，永不被修改）。
        dst (Path): 目标文件路径（父目录需已存在）。
        to_format (str): 目标格式，'docx'（wdFormatXMLDocument=12）或 'doc'（wdFormatDocument=0）。

    Raises:
        RuntimeError: 非 Windows、超时或转换未产生输出。
    """
    if platform != 'win32':
        raise RuntimeError('.doc 转换需要 Windows + Microsoft Word')
    skill_root = Path(__file__).resolve().parent.parent
    try:
        run(
            [executable, '-m', 'fill_template.conv_worker',
             str(src), str(dst), to_format],
            stdout=DEVNULL, stderr=DEVNULL, timeout=_TIMEOUT, cwd=str(skill_root),
        )
    except TimeoutExpired:
        _kill_word_orphans()
        raise RuntimeError('.doc 转换超时（Word 可能弹出对话框或无响应），已终止') from None
    if not dst.exists():
        raise RuntimeError('.doc 转换未产生输出（Word 不可用或文件损坏）')


def prepare_doc_template(doc_path: Path) -> tuple[Path, Callable[[], None]]:
    """
    ## 将 `.doc` 模板转换为系统临时目录中的 `.docx` 副本。

    原始 `.doc` 文件永不被修改（符合 template-write 的 Key Rule 2：
    Never modify the original template）。转换产物放在系统临时目录，
    调用方通过返回的 cleanup 回调在结束时清理。

    Args:
        doc_path (Path): 原始 `.doc` 模板路径。

    Returns:
        (docx_path, cleanup) (tuple[Path, Callable[[], None]]):
            临时 `.docx` 路径 + 无参清理回调（删除整个临时工作目录）。

    Raises:
        RuntimeError: 转换失败（此时临时目录已被清理）。
    """
    workdir = Path(mkdtemp(prefix='ft_doc_'))
    docx_path = workdir / (doc_path.stem + '.docx')
    try:
        _convert(doc_path, docx_path, 'docx')
    except Exception:
        rmtree_retry(workdir)
        raise

    return docx_path, lambda: rmtree_retry(workdir)


def convert_to_doc(src_docx: Path, dst_doc: Path) -> None:
    """
    ## 将已填写的 `.docx` 转换回 `.doc`（用于用户要求 `.doc` 输出时）。

    Args:
        src_docx (Path): 已填写的 `.docx` 路径。
        dst_doc (Path): 目标 `.doc` 路径（父目录需已存在）。

    Raises:
        RuntimeError: 转换失败。
    """
    _convert(src_docx, dst_doc, 'doc')
