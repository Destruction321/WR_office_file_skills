"""资源提取 — 从 OOXML 和 PDF 中提取图片、媒体和嵌入对象。"""
from pathlib import Path
from sys import stderr
from zipfile import ZipFile, BadZipFile

from .deps import ensure_import
from .ole import decompose_ole_object


# ===================================================================
#  各格式的 OOXML 资源目录和扩展名分类（图片/媒体/嵌入对象/其他）
# ===================================================================

OOXML_ASSET_FOLDERS: dict[str, list[str]] = {
    '.docx': ['word/media/', 'word/embeddings/'],
    '.pptx': ['ppt/media/', 'ppt/embeddings/'],
    '.xlsx': ['xl/media/', 'xl/embeddings/'],
    '.xls':  ['xl/media/', 'xl/embeddings/'],
}

IMAGE_EXTS: set[str] = {
    '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff',
    '.tif', '.webp', '.emf', '.wmf', '.svg', '.ico',
}
MEDIA_EXTS: set[str] = {
    '.mp3', '.wav', '.wma', '.aac', '.flac', '.ogg', '.wmv',
    '.mp4', '.avi', '.mov', '.mkv', '.webm', '.flv', '.m4v',
}
EMBED_EXTS: set[str] = {
    '.bin', '.ole', '.dat', '.pdf',
    '.doc', '.xls', '.ppt', '.zip',
}


# ===================================================================
#  公开 API
# ===================================================================

def extract_ooxml_assets(filepath: Path, assets_dir: Path, ext_key: str) -> dict[str, list[str]]:
    """
    ## 提取 OOXML 文件中的所有资源，按类别存入子目录。
    - 从 ZIP 包中提取 media/ 和 embeddings/
    目录下的文件，按图片/媒体/嵌入对象/其他分类，处理文件名冲突。

    Args:
        filepath (Path): OOXML 文件路径。
        assets_dir (Path): 保存提取资源的目录（按需创建）。
        ext_key (str): 文件扩展名 key，用于确定资源目录（如 '.docx'）。

    Returns:
        files (dict[str, list[str]]): 按类别分类的资源字典，值为描述性标签列表。
    """
    result: dict[str, list[str]] = {'images': [], 'media': [], 'embeddings': [], 'other': []}
    if not filepath.exists():
        return result

    folders = OOXML_ASSET_FOLDERS.get(ext_key, ['word/media/', 'word/embeddings/'])
    assets_dir.mkdir(parents=True, exist_ok=True)

    try:
        _extract_from_zip(filepath, folders, assets_dir, result)
    
    except BadZipFile:
        if ext_key not in ('.xls', '.doc'):
            print(f'  [警告] 非有效 ZIP/OOXML 文件: {filepath.name}', file=stderr)
    
    except Exception as e:
        print(f'  [警告] 资源提取失败 {filepath.name}: {e}', file=stderr)
    
    return result


def extract_pdf_assets(filepath: Path, assets_dir: Path) -> dict[str, list[str]]:
    """
    ## 使用 `PyMuPDF` 提取 PDF 中的图片。
    - `PyMuPDF` 是可选依赖，未安装时静默返回空结果。

    Args:
        filepath (Path): PDF 文件路径。
        assets_dir (Path): 保存提取图片的目录。

    Returns:
        result (dict[str, list[str]]): 按类别分类的资源字典。
    """
    result: dict[str, list[str]] = {'images': [], 'media': [], 'embeddings': [], 'other': []}
    img_dir = assets_dir / 'images'
    img_dir.mkdir(parents=True, exist_ok=True)

    try:
        fitz_open = ensure_import('PyMuPDF', 'fitz', attr='open')
    except ImportError:
        fitz_open = None

    if fitz_open:
        try:
            doc = fitz_open(filepath)
            for page_idx in range(len(doc)):
                _extract_page_images(doc, page_idx, img_dir, result)
            doc.close()
        except Exception as e:
            print(f'  [警告] PDF 图片提取失败: {e}', file=stderr)
    
    return result


def append_assets_summary(lines: list[str], assets_result: dict[str, list[str]]) -> None:
    """
    ## 将提取资源的结构化摘要追加到文本行末尾。

    Args:
        lines (list[str]): 输出文本行列表（直接原地修改）。
        assets_result (dict[str, list[str]]): 要摘要的资源分类字典。
    """
    lines.append('')
    lines.append('=' * 40)
    lines.append('[提取的嵌入文件]')

    cat_labels = [
        ('图片', 'images'), ('媒体', 'media'),
        ('嵌入对象', 'embeddings'), ('其他', 'other'),
    ]

    for cat, key in cat_labels:
        if not assets_result.get(key):
            continue
        lines.append(f'  {cat} ({len(assets_result[key])}):')
        for item in assets_result[key]:
            lines.append(f'    {item}')

    lines.append('=' * 40)


# ===================================================================
#  内部辅助函数
# ===================================================================

def _extract_from_zip(filepath: Path,
                      folders: list[str],
                      assets_dir: Path,
                      result: dict[str, list[str]]) -> None:
    """文件提取核心逻辑"""
    with ZipFile(filepath, 'r') as z:
        for name in z.namelist():
            if name.endswith('/'):
                continue

            matched = False
            for prefix in folders:
                if not name.startswith(prefix):
                    continue
                matched = True
                break
            if not matched:
                continue

            basename = Path(name).name
            ext = Path(basename).suffix.lower()
            cat, out_subdir, label = _classify_asset(ext, name, assets_dir, basename)
            out_subdir.mkdir(parents=True, exist_ok=True)
            out_path = out_subdir / basename
            counter = 1
            
            while out_path.exists():
                out_path = out_subdir / f'{Path(basename).stem}_{counter}{ext}'
                counter += 1
            out_path.write_bytes(z.read(name))
            result[cat].append(label)

            _try_decompose_ole(out_path, assets_dir, cat, result)


def _classify_asset(ext: str, name: str, assets_dir: Path, basename: str) -> tuple[str, Path, str]:
    """根据扩展名和路径将资源分类，返回（类别 key，输出目录路径，显示标签）。"""
    if ext in IMAGE_EXTS:
        return 'images', assets_dir / 'images', f'图片: {basename}'
    elif ext in MEDIA_EXTS:
        return 'media', assets_dir / 'media', f'媒体: {basename}'
    elif ext in EMBED_EXTS or 'embedding' in name.lower():
        return 'embeddings', assets_dir / 'embeddings', f'嵌入对象: {basename}'
    else:
        return 'other', assets_dir / 'other', f'其他: {basename}'


def _try_decompose_ole(out_path: Path, assets_dir: Path, cat: str, result: dict[str, list[str]]) -> None:
    """
    如果 out_path 是嵌入对象的 .bin，尝试 OLE 复合文档分解,
    分解结果追加到 result['embeddings'] 中。
    
    此为尽力而为步骤，失败只记标签，不向上抛异常。
    """
    if cat != 'embeddings' or out_path.suffix.lower() != '.bin':
        return

    try:
        ole_assets = decompose_ole_object(out_path, assets_dir / 'embeddings')
        for ole_path, ole_desc in ole_assets:
            if ole_path is None:
                result['embeddings'].append(ole_desc)
            else:
                result['embeddings'].append(f'  L OLE分解: {Path(ole_path).name} ({ole_desc})')
    
    except Exception:
        result['embeddings'].append('  L OLE分解失败')


def _extract_page_images(doc, page_idx: int, img_dir: Path, result: dict[str, list[str]]) -> None:
    """提取 PDF 第 page_idx 页中的所有图片，保存到 img_dir。"""
    for j, img in enumerate(doc[page_idx].get_images(full=True)):
        try:
            base_image = doc.extract_image(img[0])
            img_ext = base_image['ext']
            out_path = img_dir / f'page{page_idx+1}_img{j+1}.{img_ext}'
            out_path.write_bytes(base_image['image'])
            result['images'].append(f'图片: page{page_idx+1}_img{j+1}.{img_ext}')
        except Exception:
            pass
