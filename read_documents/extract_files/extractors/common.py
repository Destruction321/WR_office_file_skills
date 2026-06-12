"""
公共工具 — COM 僵尸清理、脚本路径解析。

各格式提取器共用这里的函数和路径常量。
"""
from pathlib import Path
from subprocess import run
from sys import platform


def kill_orphan_com(process_name: str) -> None:
    """
    清理超时后残留的 Office COM 僵尸进程。

    当 subprocess.run 抛出 TimeoutExpired 时子进程已被杀掉，
    但 PowerShell 启动的 COM 服务器可能仍残留在系统中。

    使用 taskkill /FI "STATUS eq NOT RESPONDING"，
    只杀死无响应的进程，不影响用户正在使用的 Office 窗口。
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
PPT_SCRIPT: str = str(_SCRIPT_DIR / 'extract_ppt.ps1')
DOC_SCRIPT: str = str(_SCRIPT_DIR / 'extract_doc.ps1')
XLS_SCRIPT: str = str(_SCRIPT_DIR / 'extract_xls.ps1')
