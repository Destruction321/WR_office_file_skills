"""
# docx 占位符替换（placeholder）。

- 关键设计：逐 run 替换占位符，不破坏段落格式。
如果占位符跨 run 分散（例如 "{{na" 在 run 1，"me}}" 在 run 2），
会合并相邻 run 来处理。
- 依赖：`python-docx`（缺失时自动安装）
"""

from pathlib import Path
from re import Pattern
from typing import Any

from ..deps import ensure_import
from ..replacer import Replacer


def fill_docx(output_path: Path, replacer: Replacer) -> None:
    """
    ## 在 .docx 文件中替换占位符（原地修改）。

    - 保留格式的关键：操作级别是 run，不是 paragraph。
    - 对于跨 run 的占位符，合并相邻 run 后统一替换。

    Args:
        output_path (Path): 已复制的 .docx 文件路径。
        replacer (Replacer): 占位符替换器（content_map + pattern）。
    """
    Document = ensure_import('python-docx', 'docx', attr='Document')
    doc = Document(str(output_path))

    # 替换正文段落
    for para in doc.paragraphs:
        _fill_paragraph(para, replacer)

    # 替换表格中的文字
    for table in doc.tables:
        _fill_table(table, replacer)

    # 替换页眉页脚
    for section in doc.sections:
        for para in section.header.paragraphs:
            _fill_paragraph(para, replacer)

        for para in section.footer.paragraphs:
            _fill_paragraph(para, replacer)

    doc.save(str(output_path))


def _fill_paragraph(para: Any, replacer: Replacer) -> None:
    """
    替换单个段落中的占位符。

    策略：
    1. 快速检查——没有占位符就跳过。
    2. 如果有跨 run 的占位符 -> 合并所有 run 到第一个，再替换。
    3. 否则在每个 run 内单独替换。
    """
    full_text = ''.join(run.text for run in para.runs)
    if not replacer.pattern.search(full_text):
        return

    if not _has_multi_run_match(para.runs, full_text, replacer.pattern):
        for run in para.runs:
            run.text = replacer.replace(run.text)
        return

    runs = para.runs
    if not runs:
        return

    first_run = runs[0]
    first_run.text = replacer.replace(full_text)

    for run in runs[1:]:
        run.text = ''


def _fill_table(table: Any, replacer: Replacer) -> None:
    """替换表格中所有单元格的占位符。"""
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                _fill_paragraph(para, replacer)


def _has_multi_run_match(runs: Any, full_text: str, pattern: Pattern[str]) -> bool:
    """
    检查是否有占位符跨多个 run。

    在 full_text 中找匹配，再看该匹配覆盖了多少个 run，
    超过 1 个就是跨 run 占位符。
    """
    for match in pattern.finditer(full_text):
        start, end = match.start(), match.end()
        char_pos = 0
        runs_covered = 0
        for run in runs:
            run_len = len(run.text)
            if start < char_pos + run_len:
                runs_covered += 1

            char_pos += run_len
            if char_pos > end:
                break

        if runs_covered > 1:
            return True

    return False
