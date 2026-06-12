"""共享工具 — 安全路径处理和二进制识别。"""
from contextlib import contextmanager
from pathlib import Path
from shutil import copy2
from tempfile import mkdtemp


def mktemp_in_dir(filepath: Path, prefix: str = 'tmp_') -> Path:
    """
    在源文件目录的 temp/ 子目录下创建唯一临时文件夹。
    
    Args:
        filepath (Path): 参考文件路径，临时目录将创建在该文件所在目录的 temp/ 子目录下。
        prefix (str): 临时目录文件名称前缀，默认为 'tmp_'。
    
    Returns:
        tmp_path (Path): 创建的临时目录路径。
    """
    temp_base = filepath.parent / 'temp'
    temp_base.mkdir(parents=True, exist_ok=True)
    return Path(mkdtemp(prefix=prefix, dir=str(temp_base)))


@contextmanager
def safe_open_path(filepath: Path):
    """
    当原始路径包含非 ASCII 字符时，复制到纯 ASCII 临时路径再打开。

    python-docx / python-pptx 共用的 OPC 层在处理非 ASCII 路径时会报错，
    该上下文管理器透明地将文件复制到纯 ASCII 名称的临时目录，返回该路径，使用后自动清理。

    Args:
        filepath (Path): 原始文件路径，可能包含非 ASCII 字符。
    """
    if not _has_nonascii(str(filepath)):
        yield filepath
        return
    tmp_dir = mktemp_in_dir(filepath, prefix='tmp_doc_')
    tmp_path = tmp_dir / f'doc{filepath.suffix}'
    copy2(filepath, tmp_path)
    yield tmp_path


def _has_nonascii(s: str) -> bool:
    """检查 *s* 是否包含非 ASCII 字符。"""
    try:
        s.encode('ascii')
        return False
    
    except UnicodeEncodeError:
        return True


# ===================================================================
#  魔数识别
# ===================================================================

MAGIC_SIGNATURES: list[tuple[bytes, str, str]] = [
    (b'\x89PNG\r\n\x1a\n',   '.png',  'PNG 图片'),
    (b'\xff\xd8\xff',        '.jpg',  'JPEG 图片'),
    (b'GIF8',                '.gif',  'GIF 图片'),
    (b'BM',                  '.bmp',  'BMP 图片'),
    (b'PK\x03\x04',          '.zip',  'ZIP 压缩包'),
    (b'%PDF',                '.pdf',  'PDF 文档'),
    (b'\xd0\xcf\x11\xe0',    '.ole',  'OLE 复合文档'),
    (b'RIFF',                '.dat',  'RIFF 容器'),
    (b'MZ',                  '.exe',  'PE 可执行文件'),
    (b'\x1f\x8b',            '.gz',   'GZIP 压缩'),
    (b'\x00\x00\x01\xba',    '.mpg',  'MPEG 视频'),
    (b'\x00\x00\x01\xb3',    '.mpg',  'MPEG 视频'),
    (b'ID3',                 '.mp3',  'MP3 音频'),
    (b'OggS',                '.ogg',  'OGG 媒体'),
    (b'ftyp',                '.mp4',  'MP4 视频（片段）'),
    (b'WEBP',                '.webp', 'WebP 图片'),
]


def identify_data(data: bytes) -> tuple[str, str]:
    """
    根据魔数返回二进制数据的（扩展名，描述）。
    
    Args:
        data (bytes): 待识别的二进制数据。
        
    Returns:
        ext, desc (tuple[str, str]): 识别出的文件扩展名和描述，未识别时返回 ('.bin', '未知二进制')。
    """
    for magic, ext, desc in MAGIC_SIGNATURES:
        if data[:len(magic)] == magic:
            return ext, desc
    
    return '.bin', '未知二进制'
