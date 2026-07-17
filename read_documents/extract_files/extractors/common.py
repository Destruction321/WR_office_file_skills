"""
# 公共工具 — COM 僵尸清理、脚本路径解析、提取作业封装。
- 各格式提取器共用这里的函数、路径常量和 ExtractJob 类型。
"""

from dataclasses import dataclass
from pathlib import Path
from subprocess import run
from sys import platform


@dataclass(frozen=True)
class ExtractJob:
    """
    ## 单文件提取作业：文件路径 + 资源目录 + 强制渲染页码（仅 PDF）。
    
    Attributes:
        filepath (Path): 待提取文件路径。
        assets_dir (Path | None): 资源提取目标目录（可选）
        render_pages (list[int] | None): 强制渲染为图片的页码列表（仅 PDF，1-based），
            None 表示不强制渲染。
    """
    filepath: Path
    assets_dir: Path | None = None
    render_pages: list[int] | None = None


def kill_orphan_com(process_name: str) -> None:
    """
    ## 清理超时后残留的 Office COM 僵尸进程。

    - 当 subprocess.run 抛出 TimeoutExpired 时子进程已被杀掉，但
      PowerShell 启动的 COM 服务器可能仍残留在系统中。
    - 使用 taskkill /FI
      "STATUS eq NOT RESPONDING"，只杀死无响应的进程，不影响用户正在使用的 Office 窗口。

    Args:
        process_name (str): 进程映像名称（如 'WINWORD.EXE'）。
    """
    if platform != 'win32':
        return
    try:
        run(
            ['taskkill', '/F',
             '/FI', f'IMAGENAME eq {process_name}',
             '/FI', 'STATUS eq NOT RESPONDING'],
            capture_output=True, timeout=5,
        )
    except Exception:
        pass  # 尽力而为，不因清理失败而崩溃


# COM 辅助脚本路径 — ps1_scripts/ 目录
_SCRIPT_DIR = Path(__file__).resolve().parent.parent.parent / 'ps1_scripts'
PPT_SCRIPT: Path = _SCRIPT_DIR / 'extract_ppt.ps1'
DOC_SCRIPT: Path = _SCRIPT_DIR / 'extract_doc.ps1'
XLS_SCRIPT: Path = _SCRIPT_DIR / 'extract_xls.ps1'
