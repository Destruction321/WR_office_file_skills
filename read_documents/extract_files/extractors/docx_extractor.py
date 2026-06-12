"""
.doc / .docx 提取器。

先尝试 python-docx（可处理伪装成 .doc 的 .docx），
再 COM 回退处理旧 .doc 格式。
"""
from pathlib import Path
from subprocess import run, DEVNULL, TimeoutExpired
from sys import platform

from .common import kill_orphan_com, DOC_SCRIPT
from .. import assets
from ..deps import ensure_import
from ..section import detect_chinese_heading
from ..util import safe_open_path, mktemp_in_dir


def extract_doc(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """
    先尝试 python-docx（可处理伪装成 .doc 的 .docx），再 COM 回退。

    如果 python-docx 只返回了错误信息（全是 [Error...] 或空），
    说明不是 .docx 变体，走 COM 路径。
    """
    try:
        result = extract_docx(filepath, assets_dir)
        if all(not line or line.startswith('[') for line in result):
            raise ValueError('python-docx 仅返回了错误信息')
        return result
    except Exception:
        return _extract_doc_com(filepath, assets_dir)


def extract_docx(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """通过 python-docx 提取 .docx 文件，保留段落/表格交错顺序和标题样式。"""
    lines: list[str] = []
    assets_result: dict[str, list[str]] = {}

    if assets_dir:
        assets_result = assets.extract_ooxml_assets(
            filepath, assets_dir / filepath.stem, '.docx')

    try:
        Document = ensure_import('python-docx', 'docx', attr='Document')  # type: ignore[assignment]
    except ImportError:
        return ['[Error: python-docx 未安装。执行: pip install python-docx]']

    with safe_open_path(filepath) as safe_path:
        try:
            doc = Document(str(safe_path))  # type: ignore[operator]
        except Exception as e:
            return [f'[Error: 用 python-docx 打开 .docx 失败: {e}]']

        # 始终使用有序提取——保留段落/表格的真实交错顺序
        _extract_docx_body_ordered(doc, lines)

    if assets_result:
        assets.append_assets_summary(lines, assets_result)

    return lines


def _extract_doc_com(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """通过 Windows COM 提取旧 .doc 文件。"""
    if platform != 'win32':
        return ['[Error: 旧格式 .doc 提取需要 Windows + Microsoft Office]']

    if not Path(DOC_SCRIPT).exists():
        return ['[Error: 找不到 Word 提取脚本]']

    tmp_out = mktemp_in_dir(filepath, prefix='tmp_doc_') / 'output.txt'
    try:
        run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-File', DOC_SCRIPT,
             '-DocPath', str(filepath), '-OutFile', str(tmp_out)],
            stdout=DEVNULL, stderr=DEVNULL, timeout=120
        )
        if tmp_out.exists():
            lines = tmp_out.read_text(encoding='utf-8-sig').splitlines()
            if assets_dir:
                lines.append('[提示: 旧格式 .doc 暂不支持嵌入文件提取]')
            return lines
        return ['[Error: Word 提取未产生输出]']

    except TimeoutExpired:
        kill_orphan_com('WINWORD.EXE')
        return ['[Error: Word 提取超时]']

    except Exception as e:
        return [f'[Error: 通过 COM 提取 DOC 失败: {e}]']


def _extract_docx_body_ordered(doc, lines: list[str]) -> None:
    """
    按文档顺序提取 docx 内容，添加 Markdown 标题标记。

    通过迭代 XML body 使段落和表格以真实文档顺序出现
    （有别于 doc.paragraphs + doc.tables 这种分离序列的方式）。
    """
    from docx.oxml.ns import qn

    def _text(elem) -> str:
        """收集 elem 内所有 w:t 节点文本。"""
        return ''.join(t.text or '' for t in elem.iter(qn('w:t')))

    body = doc.element.body
    for child in body:
        if child.tag == qn('w:p'):
            text = _text(child).strip()
            if not text:
                continue
            heading_level = _get_heading_style_level(child)
            if heading_level is None:
                heading_level = detect_chinese_heading(text)
            if heading_level is not None:
                lines.append(f'{"#" * min(heading_level, 6)} {text}')
            else:
                lines.append(text)

        elif child.tag == qn('w:tbl'):
            lines.append('')
            for row_elem in child.iter(qn('w:tr')):
                cells = [
                    _text(tc).strip().replace('\n', ' ').replace('\r', '')
                    for tc in row_elem.iter(qn('w:tc'))
                ]
                if any(c for c in cells):
                    lines.append(' | '.join(cells))
            lines.append('')


def _get_heading_style_level(child) -> int | None:
    """
    检查段落是否应用了 Word 标题样式。

    匹配 "Heading 1"、"Heading 2" 等样式名，返回对应级别（1-9）。
    没有标题样式则返回 None。
    """
    from docx.oxml.ns import qn
    pPr = child.find(qn('w:pPr'))
    if pPr is None:
        return None
    pStyle = pPr.find(qn('w:pStyle'))
    if pStyle is None:
        return None
    style_val = pStyle.get(qn('w:val'), '')
    if not style_val.lower().startswith('heading'):
        return None
    try:
        return int(style_val.split()[-1])
    except ValueError:
        return 1  # 解析失败默认返回级别 1
