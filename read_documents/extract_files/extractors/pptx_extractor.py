"""
.ppt / .pptx 提取器。

先尝试 python-pptx（可处理伪装成 .ppt 的 .pptx），
再 COM 回退处理旧 .ppt 格式。
"""
from pathlib import Path
from subprocess import run, DEVNULL, TimeoutExpired
from sys import platform

from .common import kill_orphan_com, PPT_SCRIPT
from .. import assets
from ..deps import ensure_import
from ..util import safe_open_path, mktemp_in_dir


def extract_pptx(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """通过 python-pptx 提取 .pptx 文件，逐幻灯片提取文字和表格。"""
    assets_result: dict[str, list[str]] = {}

    if assets_dir:
        assets_result = assets.extract_ooxml_assets(
            filepath, assets_dir / filepath.stem, '.pptx')

    try:
        Presentation = ensure_import('python-pptx', 'pptx', attr='Presentation')  # type: ignore[assignment]
    except ImportError:
        return ['[Error: python-pptx 未安装。执行: pip install python-pptx]']

    with safe_open_path(filepath) as safe_path:
        try:
            prs = Presentation(str(safe_path))  # type: ignore[operator]
        
        except Exception:
            return ['[Error: 用 python-pptx 打开 .pptx 失败]']

        lines: list[str] = []
        for i, slide in enumerate(prs.slides, 1):
            lines.append(f'--- Slide {i} ---')
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for p in shape.text_frame.paragraphs:  # type: ignore[attr-defined]
                        t = p.text.strip()
                        if not t:
                            continue
                        lines.append(t)

                if shape.has_table:
                    for row in shape.table.rows:  # type: ignore[attr-defined]
                        cells = [
                            cell.text.strip().replace('\n', ' ').replace('\r', '')
                            for cell in row.cells
                        ]
                        lines.append(' | '.join(cells))

    if assets_result:
        assets.append_assets_summary(lines, assets_result)

    return lines


def extract_ppt(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """
    先尝试 python-pptx，失败时回退 COM（旧格式兼容）。

    python-pptx 能打开部分旧 .ppt 文件（OOXML 变体），
    真正的旧 .ppt（二进制格式）才会走到 COM 路径。
    """
    try:
        result = extract_pptx(filepath, assets_dir)
        if all(not line or line.startswith('[') for line in result):
            raise ValueError('python-pptx 仅返回了错误信息')
        return result
    except Exception:
        return _extract_ppt_com(filepath, assets_dir)


def _extract_ppt_com(filepath: Path, assets_dir: Path | None = None) -> list[str]:
    """通过 Windows COM 提取旧 .ppt 文件。"""
    if platform != 'win32':
        return ['[Error: 旧格式 .ppt 提取需要 Windows + Microsoft Office]']

    if not Path(PPT_SCRIPT).exists():
        return ['[Error: 找不到 PowerPoint 提取脚本]']

    tmp_out = mktemp_in_dir(filepath, prefix='tmp_ppt_') / 'output.txt'
    try:
        run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-File', PPT_SCRIPT,
             '-PptPath', str(filepath), '-OutFile', str(tmp_out)],
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
