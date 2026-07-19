"""
# `.ppt` / `.pptx` 提取器。
- 先尝试 `python-pptx`（可处理伪装成 `.ppt` 的 `.pptx`），再 COM 回退处理旧 `.ppt` 格式。
"""

from subprocess import run, DEVNULL, TimeoutExpired
from sys import platform

from .common import ExtractJob, PPT_SCRIPT, kill_orphan_com
from .. import assets
from ..deps import ensure_import
from ..util import mktemp_in_dir


def extract_pptx(job: ExtractJob) -> list[str]:
    """
    ## 通过 `python-pptx` 提取 `.pptx` 文件，逐幻灯片提取文字和表格。

    Args:
        job (ExtractJob): 待提取作业，包含文件路径和资源目录。

    Returns:
        lines (list[str]): 提取出的文本行，失败时返回错误信息。
    """
    filepath, assets_dir = job.filepath, job.assets_dir
    assets_result: dict[str, list[str]] = {}
    if assets_dir:
        assets_result = assets.extract_ooxml_assets(filepath, assets_dir / filepath.stem, '.pptx')

    try:
        Presentation = ensure_import('python-pptx', 'pptx', attr='Presentation')
    except ImportError:
        return ['[Error: python-pptx 未安装。执行: pip install python-pptx]']

    try:
        prs = Presentation(str(filepath))
    except Exception:
        return ['[Error: 用 python-pptx 打开 .pptx 失败]']

    lines: list[str] = []
    for i, slide in enumerate(prs.slides, 1):
        lines.append(f'--- Slide {i} ---')
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    t = p.text.strip()
                    if t:
                        lines.append(t)

            if shape.has_table:
                for row in shape.table.rows:
                    cells = [
                        cell.text.strip().replace('\n', ' ').replace('\r', '')
                        for cell in row.cells
                    ]
                    lines.append(' | '.join(cells))

    if assets_result:
        assets.append_assets_summary(lines, assets_result)

    return lines


def extract_ppt(job: ExtractJob) -> list[str]:
    """
    ## 先尝试 `python-pptx`，失败时回退 COM（旧格式兼容）。
    - `python-pptx` 能打开部分旧 `.ppt` 文件（OOXML 变体），
      真正的旧 `.ppt`（二进制格式）才会走到 COM 路径。

    Args:
        job (ExtractJob): 待提取作业，包含文件路径和资源目录。

    Returns:
        lines (list[str]): 提取出的文本行，失败时返回错误信息。
    """
    try:
        result = extract_pptx(job)
        if all(not line or line.startswith('[') for line in result):
            raise ValueError('python-pptx 仅返回了错误信息')
        return result
    except Exception:
        return _extract_ppt_com(job)


def _extract_ppt_com(job: ExtractJob) -> list[str]:
    """通过 Windows COM 提取旧 .ppt 文件。"""
    filepath, assets_dir = job.filepath, job.assets_dir
    if platform != 'win32':
        return ['[Error: 旧格式 .ppt 提取需要 Windows + Microsoft Office]']
    if not PPT_SCRIPT.exists():
        return ['[Error: 找不到 PowerPoint 提取脚本]']

    tmp_out = mktemp_in_dir(filepath, prefix='tmp_ppt_') / 'output.txt'
    try:
        run(
            [
                'powershell', '-ExecutionPolicy', 'Bypass',
                '-File', str(PPT_SCRIPT),
                '-PptPath', str(filepath), '-OutFile', str(tmp_out)
            ],
            stdout=DEVNULL, stderr=DEVNULL, timeout=120
        )
        if tmp_out.exists():
            lines = tmp_out.read_text(encoding='utf-8-sig').splitlines()
            if assets_dir:
                lines.append('[提示: 旧格式 .ppt 暂不支持嵌入文件提取]')
            return lines
        return ['[Error: PowerPoint 提取未产生输出]']

    except TimeoutExpired:
        kill_orphan_com('POWERPNT.EXE')
        return ['[Error: PowerPoint 提取超时]']
    except Exception as e:
        return [f'[Error: 通过 COM 提取 PPT 失败: {e}]']
