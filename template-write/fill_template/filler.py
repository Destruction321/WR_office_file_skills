"""
# 模板填写主入口 — 格式分发。

- `fill_template(template_path, output_path, content_map, placeholder_pattern=None)`:
  解析模板格式，调用对应格式的 filler，各 filler 内部自行处理缺失依赖的自动安装。
- 仅支持 `.docx` 和纯文本（`.md`/`.txt`）。
- `.xlsx`、`.pptx`、`.csv` 不纳入支持——前者没有模板场景，pptx 模板过于复杂，csv 是数据格式不是文档。
"""

from os import makedirs
from pathlib import Path
from re import compile
from shutil import copy2

# 默认占位符模式：{{ name }}、{{name}} 等
DEFAULT_PATTERN = r'\{\{\s*(\w+)\s*\}\}'

# 支持的格式
_DOCX_EXTS = {'.docx'}
_TEXT_EXTS = {'.md', '.txt'}


def fill_template(template_path: str | Path,
                  output_path: str | Path,
                  content_map: dict[str, str],
                  placeholder_pattern: str | None = None) -> Path:
    """
    ## 在模板文件中替换占位符，生成新文档。

    Args:
        template_path (str | Path): 模板文件路径。
        output_path (str | Path): 输出文件路径（默认不覆盖已存在的文件）。
        content_map (dict[str, str]): 占位符名称 -> 替换文本。
        placeholder_pattern (str | None): 占位符正则模式。Group(1) 必须捕获占位符名称，默认匹配 {{ name }}。

    Returns:
        output_path (Path): 输出文件的 Path 对象。

    Raises:
        FileNotFoundError (FileNotFoundError): 模板不存在。
        FileExistsError (FileExistsError): 输出文件已存在。
        ValueError (ValueError): 不支持的格式。
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
    compiled = compile(placeholder_pattern or DEFAULT_PATTERN)

    if ext in _DOCX_EXTS:
        # docx：先复制模板，再原地替换
        copy2(template_path, output_path)
        from .docx_filler import fill_docx
        fill_docx(output_path, content_map, compiled)

    elif ext in _TEXT_EXTS:
        # 纯文本：读 -> 替换 -> 重新写入
        from .text_filler import fill_text
        fill_text(template_path, output_path, content_map, compiled)

    else:
        raise ValueError(
            f'不支持的格式: {ext}。'
            f'仅支持 docx 和纯文本（md/txt）。'
        )

    return output_path
