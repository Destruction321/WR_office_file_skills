"""
# 占位符替换器 — 聚合 content_map + pattern，承载替换行为。

- `Replacer.replace(text)` 取代原先在 `placeholder.py` 与 `text_filler.py` 中
  逐字重复的 `_replace_all`，作为唯一替换入口。
- leaf 模块（不依赖 filler/docx/text_filler），避免循环导入。
"""

from dataclasses import dataclass
from re import Pattern, Match


@dataclass(frozen=True)
class Replacer:
    """
    占位符替换规格：名称映射 + 占位符正则（group(1) 捕获占位符名称）。
    
    Attributes:
        content_map (dict[str, str]): 占位符名称 -> 替换内容
        pattern (Pattern[str]): 占位符正则，group(1) 捕获占位符名称
    """
    content_map: dict[str, str]
    pattern: Pattern[str]

    def replace(self, text: str) -> str:
        """
        将 text 中所有占位符替换为 content_map 中的值，未命中则原样保留。
        
        Args:
            text (str): 待替换的文本
        
        Returns:
            replaced_text (str): 替换后的文本
        """
        def replacer(match: Match[str]) -> str:
            name = match.group(1)
            return self.content_map.get(name, match.group(0))

        return self.pattern.sub(replacer, text)
