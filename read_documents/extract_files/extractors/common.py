"""
# 公共工具 — COM 僵尸清理、提取作业封装。
- 各格式提取器共用这里的函数和 ExtractJob 类型。
- COM 调用（pywin32）直接在进程内完成，不再需要外部 ps1 脚本。
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
    ## 清理 COM 调用失败后残留的 Office 僵尸进程。

    - 当 ``Application.Quit()`` 失败或进程被中断时，Office 进程可能仍残留。
    - 使用 ``taskkill /FI "STATUS eq NOT RESPONDING"``，只杀死无响应的进程，
      不影响用户正在使用的 Office 窗口。

    Args:
        process_name (str): 进程映像名称（如 'WINWORD.EXE'）。
    """
    if platform != 'win32':
        return
    try:
        run(
            [
                'taskkill', '/F',
                '/FI', f'IMAGENAME eq {process_name}',
                '/FI', 'STATUS eq NOT RESPONDING'
            ],
            capture_output=True, timeout=5,
        )
    except Exception:
        pass  # 尽力而为，不因清理失败而崩溃
