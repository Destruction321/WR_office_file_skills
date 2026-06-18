"""依赖管理 — 自动安装缺失的 Python 包。用的时候发现缺了再装。"""
from importlib import import_module
from subprocess import check_call
from sys import executable, stderr
from typing import Any


def _install_one(pip_name: str) -> None:
    """安装单个 pip 包，失败时抛异常。"""
    pip_cmd = [executable, '-m', 'pip', 'install', pip_name]
    print(f'正在安装 {pip_name} ...', file=stderr)
    check_call(pip_cmd, timeout=180)


def ensure_import(pip_name: str, import_name: str | None = None, attr: str | None = None) -> Any:
    """
    ## 确保某个包可导入。如果缺了，自动 pip install 之后再试。

    Args:
        pip_name (str): pip 上的包名（如 'python-docx'）。
        import_name (str | None): Python import 名（如 'docx'），默认等于 pip_name。
        attr (str | None): 如果指定，返回模块的该属性（如 'Document'），否则返回模块本身。

    Returns:
        attr (Any): 模块对象（attr=None）或模块中的指定属性（attr=...）。

    Raises:
        ImportError (ImportError): 安装后仍然导入失败。
    """
    name = import_name or pip_name
    try:
        mod = import_module(name)
    except ImportError:
        _install_one(pip_name)
        mod = import_module(name)

    if attr:
        return getattr(mod, attr)
    return mod
