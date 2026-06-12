# Claude Skills

Claude Code 技能集合 —— 增强 AI 对文档文件的读写能力。

## 技能一览

| 技能                                      | 描述                                                                     |
| ----------------------------------------- | ------------------------------------------------------------------------ |
| [read_documents](read_documents/SKILL.md) | 读取中文路径的 Office 文档和 PDF（docx, pptx, xlsx, pdf, doc, ppt, xls） |
| [template-write](template-write/SKILL.md) | 基于模板填写内容生成新文档，原模板保持不动                               |

---

## read_documents — 文档读取

读取中文路径（或其他非 ASCII 路径）的文档文件，避免 shell 编码损坏和 OPC 库编码问题。

**支持格式：**

| 格式                        | 方案                                    |
| --------------------------- | --------------------------------------- |
| `.docx` / `.pptx` / `.xlsx` | OOXML 格式，Python 库直接读取（跨平台） |
| `.pdf`                      | pdfplumber + PyMuPDF（跨平台）          |
| `.xls`                      | xlrd 直接处理（跨平台）                 |
| `.doc` / `.ppt`             | COM 回退（仅 Windows + Office）         |

**特性：**

- 自动安装缺失的 Python 依赖
- 支持按关键字搜索文件（`--root` + `--glob`）
- 支持章节提取（`--section`）节省 tokens
- 提取图片/媒体/OLE 嵌入资源
- 用完清理临时文件

---

## template-write — 模板填写

基于已有模板文件，读取 → 复制 → 填写，生成新文档。

**工作流程：**

1. 读取模板，识别占位符（`{{xxx}}`、`[xxx]`、`<xxx>`、下划线留白等）
2. 复制一份副本到同目录（原模板不动）
3. 用 `python-docx` / `openpyxl` / `python-pptx` 填入内容
4. 读回验证

**支持格式：** `.docx`、`.xlsx`、`.pptx`、`.md`、`.txt`、`.csv`

---

## 安装

```bash
# 克隆到 Claude Code 技能目录
git clone https://github.com/Destruction321/read_documents.git ~/.claude/skills

# 安装全部依赖（如需）
pip install python-docx python-pptx pdfplumber PyMuPDF openpyxl olefile xlrd
```

> 各 skill 首次使用时若发现缺失包会自动安装，无需手动预装。

---

## 项目结构

```txt
~/.claude/skills/
├── README.md                    # 本文件 — 项目说明
├── .gitignore
├── read_documents/              # 文档读取技能
│   ├── SKILL.md                 # 给 AI 的调用指引
│   ├── extract_files/           # Python 包
│   │   ├── __init__.py          # 公开 API：extract_file, EXTRACTORS
│   │   ├── __main__.py          # CLI 入口（python -m）
│   │   ├── cli.py               # 参数解析、流程编排
│   │   ├── discovery.py         # 文件发现、MSYS 路径转换
│   │   ├── deps.py              # 自动安装依赖
│   │   ├── assets.py            # 图片/媒体/OLE 资源提取
│   │   ├── ole.py               # OLE 复合文档分解
│   │   ├── section.py           # 按关键字过滤小节
│   │   ├── util.py              # 安全路径、魔数识别
│   │   └── extractors/          # 按格式拆分的提取器子包
│   │       ├── __init__.py      # 格式分发器 + extract_file()
│   │       ├── common.py        # COM 清理、脚本路径
│   │       ├── docx_extractor.py
│   │       ├── pptx_extractor.py
│   │       ├── xlsx_extractor.py
│   │       └── pdf_extractor.py
│   └── ps1_scripts/             # COM 回退脚本（仅 Windows）
│       ├── extract_doc.ps1
│       ├── extract_ppt.ps1
│       └── extract_xls.ps1
└── template-write/              # 模板填写技能
    ├── SKILL.md                 # 给 AI 的调用指引
    └── fill_template/           # 可复用的模板填写包
        ├── __init__.py
        ├── __main__.py          # CLI 入口（python -m）
        ├── cli.py               # 参数解析
        ├── filler.py            # 主入口、格式分发
        ├── deps.py              # 自动安装依赖
        ├── docx_filler.py       # 保留格式的 run 级替换
        ├── xlsx_filler.py       # 单元格级替换
        ├── pptx_filler.py       # 幻灯片占位符替换
        └── text_filler.py       # md/txt/csv 替换
```
