"""文件发现 — 遍历目录或读取路径文件。"""
from os import environ, sep
from pathlib import Path
from re import match
from sys import platform, stderr

EXTENSIONS: set[str] = {'.doc', '.docx', '.pptx', '.ppt', '.pdf', '.xlsx', '.xls'}


def normalize_path(path: str) -> str:
    """
    ## 将 MSYS / Cygwin 风格的路径转换为 Windows 原生路径。

    - 在 Windows 的 Git Bash（MSYS2）下，Shell 将 /tmp/ 视为 Windows TEMP 目录，
      但 Python（原生 Windows 进程）不识别 /tmp/，
      它只理解 C:\\Users\\\\...\\Temp\\ 这样的 Windows 绝对路径。
    - 此函数检测常见 MSYS 模式并将其转换，使 Python 能打开文件。
      在非 Windows 平台上原样返回。

    ## 支持的转换（仅 Windows）:
    1. /tmp/...  ->  %TEMP%\\\\...
    2. /c/...    ->  C:\\\\...  （盘符映射）
    3. /home/user/...  ->  %USERPROFILE%\\\\...

    - 如果路径看起来已经是 Windows 路径（首字符后有 :，如 C:\\\\...），则原样返回。

    Args:
        path (str): 可能是 MSYS 风格的路径。

    Returns:
        path (str): Windows 原生路径（如果在 Windows 上）或原样路径（非 Windows）。
    """
    # 非 Windows：无需处理
    if platform != 'win32':
        return path

    # 已经是 Windows 路径？（首字符后有 ':'，如 C:\...）
    if ':' in path[1:]:
        return path

    # UNC 路径（\\server\share\...）—— 已经是原生 Windows
    if path.startswith('\\\\'):
        return path

    # /tmp/ -> Windows TEMP 目录
    if path.startswith('/tmp/') or path == '/tmp':
        suffix = path[5:] if path.startswith('/tmp/') else ''
        default_temp = Path(environ.get('USERPROFILE', '')) / 'AppData' / 'Local' / 'Temp'
        return str(Path(environ.get('TEMP', str(default_temp))) / suffix)

    # /c/... -> C:\... （MSYS 盘符映射）
    m = match(r'^/([a-zA-Z])/(.*)', path)
    if m:
        rest = m.group(2).replace('/', sep)
        return f'{m.group(1).upper()}:{sep}{rest}'

    # /home/user/... -> USERPROFILE
    if path.startswith('/home/'):
        rest = path[6:]  # 去掉 /home/
        return str(Path(environ.get('USERPROFILE', Path.home())) / rest.replace('/', sep))

    # 回退：原样返回
    return path


def find_files(root: str, pattern: str | None = None, max_depth: int = 6) -> list[str]:
    """
    ## 递归查找 root 下所有支持的文档文件。

    Args:
        root (str): 搜索的根目录。
        pattern (str | None): 可选的子串模式（不区分大小写），仅匹配文件名或路径中包含该子串的文件。
        max_depth (int): 最大目录深度，超过则不再进入子目录（默认 6）。

    Returns:
        paths (list[str]): 找到的文件路径列表。
    """
    results: list[str] = []
    base_depth = root.rstrip(sep).count(sep)

    def on_error(err: OSError) -> None:
        print(f'  [警告] 无法访问: {err}', file=stderr)

    for dirpath, dirnames, filenames in Path(root).walk(on_error=on_error):
        dir_path = str(dirpath)
        depth = dir_path.count(sep) - base_depth
        if depth > max_depth:
            dirnames.clear()
            continue

        for f in filenames:
            if f.startswith('~$'):
                continue  # 跳过 Office 临时锁文件
            if Path(f).suffix.lower() not in EXTENSIONS:
                continue
            if pattern:
                pat_low = pattern.lower()
                if pat_low not in f.lower() and pat_low not in dir_path.lower():
                    continue
 
            results.append(str(dirpath / f))

    return results


def read_paths_file(path: str) -> list[str]:
    """
    ## 读取 UTF-8 路径文件，每行一个绝对路径，并标准化每个路径。

    Args:
        path (str): 路径文件路径。

    Returns:
        paths (list[str]): 解析并标准化后的路径列表。
    """
    resolved = normalize_path(path)
    paths: list[str] = []
    with open(resolved, 'r', encoding='utf-8') as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            
            # 标准化文件内的每个路径
            paths.append(normalize_path(line))
    
    return paths
