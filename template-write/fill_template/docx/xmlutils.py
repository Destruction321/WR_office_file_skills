"""
# docx XML 工具 — `w:p` 元素的样式与文本提取。

底层 OXML 访问，`scanner` 与 `locator` 共用。
"""

def get_style_name(p_elem, qn) -> str:
    """
    ## 提取段落的 `w:pStyle 值`，无样式时返回空字符串。
    
    Args:
        p_elem: `w:p` 元素
        qn: docx.oxml.ns.qn 函数
        
    Returns:
        style (str): 段落样式名，未设置时返回空字符串。
    """
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return ""
    
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        return ""
    
    return pStyle.get(qn("w:val"), "")


def text_of(p_elem, qn) -> str:
    """
    ## 收集 `w:p` 元素内所有 `w:t` 文本。
    
    Args:
        p_elem: `w:p` 元素
        qn: docx.oxml.ns.qn 函数
        
    Returns:
        text (str): 段落文本，未设置时返回空字符串。
    """
    return "".join(t.text or "" for t in p_elem.iter(qn("w:t")))
