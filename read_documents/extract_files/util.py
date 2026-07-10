"""共享工具 — 临时目录与魔数识别。"""
from pathlib import Path
from tempfile import mkdtemp


def mktemp_in_dir(filepath: Path, prefix: str = 'tmp_') -> Path:
    """
    ## 在源文件目录的 temp/ 子目录下创建唯一临时文件夹。

    Args:
        filepath (Path): 参考文件路径。
        prefix (str): 临时目录名称前缀，默认为 'tmp_'。

    Returns:
        temp_path (Path): 创建的临时目录路径。
    """
    temp_base = filepath.parent / 'temp'
    temp_base.mkdir(parents=True, exist_ok=True)
    return Path(mkdtemp(prefix=prefix, dir=str(temp_base)))


# ===================================================================
#  魔数识别 — 通过文件头部字节识别文件类型
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
    ## 根据头部魔数识别二进制数据的文件类型。

    Args:
        data (bytes): 待识别的二进制数据（至少前 64 字节）。

    Returns:
        (ext, desc) (tuple[str, str]): (扩展名, 描述) 元组。未识别时返回 ('.bin', '未知二进制')。
    """
    for magic, ext, desc in MAGIC_SIGNATURES:
        if data[:len(magic)] == magic:
            return ext, desc
    return '.bin', '未知二进制'
