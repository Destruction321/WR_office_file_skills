"""
模板填写主入口 — 格式分发。

fill_template(template_path, output_path, content_map, placeholder_pattern=None)
    解析模板格式，调用对应格式的 filler。
    各 filler 内部自行处理缺失依赖的自动安装。
"""
from os import makedirs
from pathlib import Path
from re import compile
from shutil import copy2

# 默认占位符模式：{{ name }}、{{name}} 等
DEFAULT_PATTERN = r'\{\{\s*(\w+)\s*\}\}'


def fill_template(template_path, output_path, content_map, placeholder_pattern=None):
    """
    在模板文件中替换占位符，生成新文档。

    Args:
        template_path: 模板文件路径。
        output_path: 输出文件路径（默认不覆盖已存在的文件）。
        content_map: dict，占位符名称 -> 替换文本。
        placeholder_pattern: 占位符正则模式。Group(1) 必须捕获占位符名称。
                            默认为 r'\\{\\{\\s*(\\w+)\\s*\\}\\}'，匹配 {{ name }}。

    Returns:
        输出文件的 Path 对象。

    Raises:
        FileNotFoundError: 模板不存在。
        FileExistsError: 输出文件已存在。
        ValueError: 不支持的格式。
    """
    template_path = Path(template_path)
    output_path = Path(output_path)

    if not template_path.exists():
        raise FileNotFoundError(f'模板不存在: {template_path}')
    if output_path.exists():
        raise FileExistsError(f'输出文件已存在: {output_path}')

    ext = template_path.suffix.lower()

    makedirs(output_path.parent, exist_ok=True)

    # 编译占位符正则
    pat = placeholder_pattern or DEFAULT_PATTERN
    compiled = compile(pat)

    # 二进制格式：先复制模板，再原地修改
    # 文本格式：读 -> 替换 -> 重新写入
    if ext == '.docx':
        copy2(template_path, output_path)
        from .docx_filler import fill_docx
        fill_docx(output_path, content_map, compiled)
    elif ext == '.xlsx':
        copy2(template_path, output_path)
        from .xlsx_filler import fill_xlsx
        fill_xlsx(output_path, content_map, compiled)
    elif ext == '.pptx':
        copy2(template_path, output_path)
        from .pptx_filler import fill_pptx
        fill_pptx(output_path, content_map, compiled)
    elif ext in ('.md', '.txt', '.csv'):
        from .text_filler import fill_text
        fill_text(template_path, output_path, content_map, compiled, ext)
    else:
        raise ValueError(f'不支持的格式: {ext}')

    return output_path
