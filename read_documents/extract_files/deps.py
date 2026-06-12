"""依赖管理 — 自动安装缺失的 Python 包。"""
from pathlib import Path
from subprocess import run, TimeoutExpired
from sys import executable, stderr


# ===================================================================
#  pip 安装时的包名 → Python __import__() 所需的模块名。
#  大部分一致，少数例外：PyMuPDF 的 import 名是 fitz。
# ===================================================================

PIP_TO_IMPORT: dict[str, str] = {
    'python-docx': 'docx',
    'python-pptx': 'pptx',
    'pdfplumber': 'pdfplumber',
    'PyPDF2': 'PyPDF2',
    'PyMuPDF': 'fitz',
    'openpyxl': 'openpyxl',
    'xlrd': 'xlrd',
    'olefile': 'olefile',
}

# ===================================================================
#  文件扩展名 → 该格式所需安装的 pip 包列表。
#  一种格式可能依赖多个包（如 .pdf），也可能多个格式共用同一包（如 .docx / .doc）。
# ===================================================================

FORMAT_DEPS: dict[str, list[str]] = {
    '.docx': ['python-docx'],
    '.doc': ['python-docx'],
    '.pptx': ['python-pptx'],
    '.ppt': ['python-pptx'],
    '.pdf': ['pdfplumber', 'PyMuPDF'],
    '.xlsx': ['openpyxl'],
    '.xls': ['xlrd'],
}


def check_and_install_deps(file_list: list[str]) -> None:
    """
    检查并按需自动安装缺失的 Python 包。

    Args:
        file_list (list[str]): 文档文件路径列表，用于推断需要哪些包。
    """
    needed: set[str] = set()
    for ext in {Path(f).suffix.lower() for f in file_list}:
        needed.update(FORMAT_DEPS.get(ext, []))
    
    needed.add('olefile')  # 所有 OOXML 格式均可受益
    missing: list[str] = []
    for pkg in sorted(needed):
        import_name = PIP_TO_IMPORT.get(pkg, pkg)
        try:
            __import__(import_name)
        
        except ImportError:
            missing.append(pkg)

    if not missing:
        return

    pip_cmd = [executable, '-m', 'pip', 'install'] + missing
    print(f'正在安装缺失的包: {" ".join(missing)}', file=stderr)
    try:
        result = run(pip_cmd, capture_output=True, text=True, timeout=180)
        if result.returncode == 0:
            print(f'  安装成功。', file=stderr)
            return

        print(f'  安装失败，请手动执行:', file=stderr)
        print(f'  {" ".join(pip_cmd)}', file=stderr)
    
    except TimeoutExpired:
        print(f'  安装超时，请手动执行:', file=stderr)
        print(f'  {" ".join(pip_cmd)}', file=stderr)
    
    except Exception as e:
        print(f'  安装失败: {e}', file=stderr)
        print(f'  请手动执行: {" ".join(pip_cmd)}', file=stderr)
