"""OLE 复合文档分解 — 从 .bin 嵌入对象中提取原生数据。"""
from pathlib import Path
from sys import stderr

from .deps import ensure_import
from .util import identify_data, MAGIC_SIGNATURES


def decompose_ole_object(filepath: Path, out_dir: Path) -> list[tuple[str | None, str]]:
    """
    ## 解析 OLE 嵌入对象（来自 OOXML embeddings/ 的 .bin），提取原生数据。
    - 通过 `olefile` 库打开 OLE 容器，读取 Ole10Native 流，识别嵌入文件的真实类型并提取到输出目录。

    Args:
        filepath (Path): OLE 对象的 .bin 文件路径。
        out_dir (Path): 提取出的文件输出目录（自动创建）。

    Returns:
        paths (list[tuple[str | None, str]]): 文件路径或
            None, 描述元组的列表，None 表示提取失败或加密条目。
    """
    extracted: list[tuple[str | None, str]] = []
    if not filepath.exists() or filepath.stat().st_size < 64:
        return extracted

    try:
        OleFileIO = ensure_import('olefile', attr='OleFileIO')
    except ImportError:
        return extracted
    
    try:
        ole = OleFileIO(str(filepath))
        native_data: bytes | None = None
        for parts in ole.listdir():
            name = parts[-1] if parts else ''
            if 'Ole10Native' in name:
                native_data = ole.openstream(parts).read()
                break

        ole.close()
        if native_data is None:
            return extracted

        offset_info = _find_ole_embedded_offset(native_data)
        if offset_info is None:
            return extracted

        strings, data_start = offset_info
        raw_data = native_data[data_start:]
        ext, desc = identify_data(raw_data)
        filename = strings[0] if strings else 'embedded_object'
        stem = Path(filename).stem if '.' in filename else filename
        stem = stem.strip().replace('\x00', '').replace('\x01', '') or 'embedded'
        out_path = out_dir / f'{stem}{ext}'
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(raw_data)
        extracted.append((str(out_path), desc))

    except Exception as e:
        print(f'  [警告] OLE 分解失败: {e}', file=stderr)

    return extracted


def _find_ole_embedded_offset(native_data: bytes) -> tuple[list[str], int] | None:
    """
    扫描 Ole10Native 流，定位嵌入文件的起始偏移和文件名。

    - 在 Ole10Native 二进制数据中，嵌入文件的原始文件名以空字节结尾的
      ASCII 字符串形式存储在头部，实际文件数据紧随其后。
      函数先扫描字符串区域，再用魔数定位数据起始位置。
    """
    if len(native_data) < 16:
        return None

    # 扫描空字节分隔的 ASCII 字符串（通常是原始文件名）
    strings: list[str] = []
    pos = 8
    while pos < min(len(native_data), 1024):
        end = native_data.find(b'\x00', pos)
        if end == -1 or end - pos > 512:
            break
        
        s = native_data[pos:end]
        if s:
            try:
                strings.append(s.decode('ascii', errors='replace'))
            except Exception:
                strings.append(repr(s))
        
        pos = end + 1
        if pos < len(native_data) and native_data[pos:pos + 1] == b'\x00':
            break

    # 在字符串区域之后找已知魔数，确定数据起始偏移
    data_start: int | None = None
    search_start = max(0, pos - 10)
    for magic, _, _ in MAGIC_SIGNATURES:
        idx = native_data.find(magic, search_start)
        if idx != -1 and (data_start is None or idx < data_start):
            data_start = idx

    if data_start is None:
        data_start = pos + 2
    if data_start >= len(native_data):
        return None

    return strings, data_start
