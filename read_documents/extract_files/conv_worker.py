"""
# COM 转换子进程入口（由 convert.py 以 subprocess 带超时调用）。

将 COM 转换隔离到子进程：Word/PowerPoint/Excel 一旦卡死（弹窗、无响应），
父进程可以超时终止本进程并清理 Office 僵尸，避免整个提取/填写流程被阻塞。
"""

import sys
from pathlib import Path

from .convert import _APPS, _convert_via_com


def main() -> None:
    src, dst = Path(sys.argv[1]), Path(sys.argv[2])
    progid, save_method, fmt = _APPS[src.suffix.lower()]
    _convert_via_com(src, dst, progid, save_method, fmt)


if __name__ == '__main__':
    main()
