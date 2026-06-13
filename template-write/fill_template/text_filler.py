"""
# 纯文本模板填写器（md/txt/csv）

- 直接替换全文中的占位符。
- Windows 中文环境下写 UTF-8 BOM，避免 GBK 乱码。
"""

from locale import getdefaultlocale
from os import name
from pathlib import Path
from re import Pattern, Match


def fill_text(template_path: Path,
              output_path: Path,
              content_map: dict[str, str],
              pattern: Pattern[str]) -> None:
    """
    ## 在纯文本模板文件中替换占位符。

    Args:
        template_path (Path): 源模板文件路径。
        output_path (Path): 输出文件路径。
        content_map (dict[str, str]): 占位符名称到替换文本的映射。
        pattern (Pattern[str]): 占位符正则。
    """
    encoding = _detect_encoding(template_path)

    with open(template_path, 'r', encoding=encoding) as f:
        content = f.read()

    content = _replace_all(content, content_map, pattern)

    # Windows 中文系统：写 UTF-8 BOM 避免 Excel 打开乱码
    use_bom = _needs_bom()

    with open(output_path, 'wb') as f:
        if use_bom:
            f.write(b'\xef\xbb\xbf')
        f.write(content.encode('utf-8'))


def _detect_encoding(path: Path) -> str:
    """检测文件编码，优先 chardet，回退到根据系统区域设置猜测。"""
    try:
        import chardet
        with open(path, 'rb') as f:
            raw = f.read(4096)
        result = chardet.detect(raw)
        return result.get('encoding') or 'utf-8'
    
    except ImportError:
        # chardet 不可用时，根据系统区域猜测
        try:
            lang = getdefaultlocale()[0]
            return 'gbk' if lang and 'zh' in lang else 'utf-8'
        
        except Exception:
            return 'utf-8'


def _needs_bom() -> bool:
    """检查是否需要在中文 Windows 系统上写 UTF-8 BOM。"""
    if name != 'nt':
        return False
    try:
        lang = getdefaultlocale()[0]
        return bool(lang and 'zh' in lang)
    
    except Exception:
        return False


def _replace_all(text: str, content_map: dict[str, str], pattern: Pattern[str]) -> str:
    """将 text 中所有占位符替换为 content_map 中的值。"""
    def _replacer(match: Match[str]) -> str:
        name = match.group(1)
        return content_map.get(name, match.group(0))
    
    return pattern.sub(_replacer, text)
