# OpenCode Skills

OpenCode 技能集合 —— 增强 AI 对文档文件的读写能力。

~~没啥大用，孩子不懂事写着玩的（~~

## 技能一览

| 技能                                      | 描述                                                                     |
| ----------------------------------------- | ------------------------------------------------------------------------ |
| [read_documents](read_documents/SKILL.md) | 读取中文路径的 Office 文档和 PDF（docx, pptx, xlsx, pdf, doc, ppt, xls） |
| [template-write](template-write/SKILL.md) | docx 节级填充（扫描/验证）与占位符替换，原模板不动                       |

---

## read_documents — 文档读取

读取中文路径（或其他非 ASCII 路径）的文档文件，避免 shell 编码损坏。

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

向已有模板写入内容生成新文档（原模板不动）。两种模式：

**节级填充**（主要能力，面向结构化报告）：`--scan` 输出标题结构与样式 -> AI 据此决策 -> `--section-data-file` 按标题定位节、以样式名定边界、清空旧内容、插入文本/图片、读回验证。工具只做机械操作，智能判断交给 AI。支持 `父标题 / 子标题` 限定语法消歧重名标题。

**占位符替换**（简单场景）：识别 `{{name}}` 等占位符并替换。

| 模式       | 关键参数                                                                |
| ---------- | ----------------------------------------------------------------------- |
| 扫描结构   | `--scan`                                                                |
| 节级填充   | `--section-data-file`、`--heading-style`、`--section-mode`、`--dry-run` |
| 占位符替换 | `--set` / `--data` / `--data-file`、`--pattern`                         |

**支持格式：** `.docx`（节级填充 + 占位符替换）、`.md`/`.txt`（占位符替换）

> `.xlsx`/`.pptx`/`.csv` 不纳入支持——无模板场景或模板过于复杂。

---

## 安装

```bash
# 克隆到 opencode 技能目录
git clone https://github.com/Destruction321/WR_office_file_skills.git ~/.config/opencode/skills

# 安装全部依赖（如需）
pip install python-docx python-pptx PyMuPDF openpyxl chardet olefile xlrd
```

> 各 skill 首次使用时若发现缺失包会自动安装，无需手动预装。

---

## 项目结构

```txt
~/.config/opencode/skills/
├── README.md                    # 本文件 — 项目说明
├── AGENTS.md                    # 项目级 AI 指引（测试规则等）
├── .gitignore
├── read_documents/              # 文档读取技能
│    ├── SKILL.md                 # 给 AI 的调用指引
│    ├── extract_files/           # Python 包
│    │    ├── __init__.py          # 公开 API：extract_file, EXTRACTORS
│    │    ├── __main__.py          # CLI 入口（python -m）
│    │    ├── cli.py               # 参数解析、流程编排
│    │    ├── discovery.py         # 文件发现、MSYS 路径转换
│    │    ├── deps.py              # 自动安装依赖
│    │    ├── assets.py            # 图片/媒体/OLE 资源提取
│    │    ├── ole.py               # OLE 复合文档分解
│    │    ├── section.py           # 按关键字过滤小节
│    │    ├── util.py              # 临时目录、魔数识别
│    │    └── extractors/          # 按格式拆分的提取器子包
│    │         ├── __init__.py      # 格式分发器 + extract_file()
│    │         ├── common.py        # COM 清理、脚本路径
│    │         ├── docx_extractor.py
│    │         ├── pptx_extractor.py
│    │         ├── xlsx_extractor.py
│    │         └── pdf_extractor.py
│    └── ps1_scripts/             # COM 回退脚本（仅 Windows）
│         ├── extract_doc.ps1
│         ├── extract_ppt.ps1
│         └── extract_xls.ps1
└── template-write/              # 模板填写技能
     ├── SKILL.md                 # 给 AI 的调用指引
     └── fill_template/           # 可复用的模板填写包
          ├── __init__.py
          ├── __main__.py            # CLI 入口（python -m）
          ├── cli.py                 # 参数解析、模式分发
          ├── filler.py              # 占位符替换主入口、格式分发
          ├── deps.py                # 自动安装依赖
          ├── md_parser.py           # Markdown -> 节内容解析
          ├── text_filler.py         # md/txt 占位符替换
          └── docx/                  # docx 操作子包
               ├── __init__.py          # 公开 API：scan_docx / fill_docx_sections / fill_docx
               ├── section_filler.py    # 节级填充编排（会话 + 流程 + 验证）
               ├── scanner.py           # 模板结构分析与样式提示
               ├── locator.py           # 标题索引查找与节边界检测
               ├── elements.py          # 段落构建与插入/删除
               ├── xmlutils.py          # OXML 工具（get_style_name / text_of）
               └── placeholder.py       # 占位符替换（run 级）
```
