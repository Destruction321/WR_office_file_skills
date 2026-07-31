"""
# `.ppt` / `.pptx` 提取器。
- 先尝试 `python-pptx`（可处理伪装成 `.ppt` 的 `.pptx`），失败时通过
  pywin32 COM SaveAs 将 `.ppt` 转换为 `.pptx` 后复用 python-pptx 提取路径。
- 转换在系统临时目录进行，原始 `.ppt` 只读打开、永不被修改。
"""

from .common import ExtractJob
from .. import assets
from ..deps import ensure_import


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
    ## 提取旧格式 `.ppt` 文件。

    - 先尝试 `python-pptx`（可处理伪装成 `.ppt` 的 `.pptx`）。
    - 失败（真二进制 `.ppt`）时通过 COM SaveAs 转换为 `.pptx`，
      再复用 `extract_pptx` 的 python-pptx 提取路径。

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
        return _extract_ppt_converted(job)


def _extract_ppt_converted(job: ExtractJob) -> list[str]:
    """通过 COM 转换 .ppt -> .pptx 后，复用 python-pptx 提取（含资源提取）。"""
    from ..convert import convert_to_modern
    try:
        modern, cleanup = convert_to_modern(job.filepath)
    except RuntimeError as e:
        return [f'[Error: {e}]']
    
    try:
        lines = extract_pptx(ExtractJob(filepath=modern, assets_dir=job.assets_dir))
        if not lines or all(not line or line.startswith('[') for line in lines):
            return lines
        lines.append('[提示: 已通过 PowerPoint COM 转换为 .pptx 后提取]')
        return lines
    finally:
        cleanup()
