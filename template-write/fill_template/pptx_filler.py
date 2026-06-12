"""
pptx 模板填写器。

遍历幻灯片的所有形状（文本框 + 表格），替换占位符。

依赖：python-pptx（缺失时自动安装）
"""
from .deps import ensure_import


def fill_pptx(output_path, content_map, pattern):
    """
    在 .pptx 文件中替换占位符（原地修改）。

    Args:
        output_path: 已复制的 .pptx 文件路径。
        content_map: dict，占位符名称 -> 替换文本。
        pattern: 编译好的占位符正则。
    """
    Presentation = ensure_import('python-pptx', 'pptx', attr='Presentation')  # type: ignore[assignment]

    prs = Presentation(str(output_path))  # type: ignore[operator]

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                _fill_runs(shape.text_frame.paragraphs, content_map, pattern)
            if shape.has_table:
                _fill_table(shape.table, content_map, pattern)

    prs.save(str(output_path))


def _fill_runs(paragraphs, content_map, pattern):
    """替换一组段落中所有 run 的占位符。"""
    for para in paragraphs:
        for run in para.runs:
            run.text = _replace_all(run.text, content_map, pattern)


def _fill_table(table, content_map, pattern):
    """替换表格中所有单元格的占位符。"""
    for row in table.rows:
        for cell in row.cells:
            _fill_runs(cell.text_frame.paragraphs, content_map, pattern)


def _replace_all(text, content_map, pattern):
    """将 text 中所有占位符替换为 content_map 中的值。"""
    def _replacer(match):
        name = match.group(1)
        return content_map.get(name, match.group(0))
    return pattern.sub(_replacer, text)
