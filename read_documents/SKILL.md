---
name: read_documents
description: >
    Read document files (doc, docx, ppt, pptx, xls, xlsx, pdf) when file paths
    contain Chinese or other non-ASCII characters. Avoids shell encoding corruption.
    Extracts content to current working directory
    and cleans up temp files. Supports --section for extracting only a specific
    section (saves tokens). Auto-installs missing Python packages.
triggers:
    - filePatterns: "*.doc,*.docx,*.ppt,*.pptx,*.xls,*.xlsx,*.pdf"
      autoSuggest: false
tools:
    - read
    - bash
    - write
---

# Read Documents Skills

## Auto-invoke rule

Invoke `Skill(skill="read_documents")` when you need to **read** a document file (`.doc`, `.docx`, `.ppt`, `.pptx`, `.xls`, `.xlsx`, `.pdf`) **and** the file path contains Chinese or other non-ASCII characters — or when reading a document file with the Read tool fails due to encoding issues. Do NOT invoke for ASCII-only paths where simpler approaches work.

Supported formats: `.doc` `.docx` `.pptx` `.ppt` `.pdf` `.xlsx` `.xls`

## ⚠️ Reading a docx to FILL it? Use `template-write` instead

If the goal is to **write content into** a docx template (fill sections, insert text/images), do NOT:

- probe paragraph indices with inline `python-docx` scripts
- hand-write `insert_paragraph_before` loops

That path reliably produces reversed content, wrong styles, and missing images. Instead use the **`template-write`** skill: `python -m fill_template --scan` shows the structure + styles, then `--section-data-file` fills it correctly with image embedding and auto-verification.

> **Location**: `template-write` lives in the **same skills directory as this skill** — i.e. `~/.config/opencode/skills/template-write/` (or `~/.claude/skills/template-write/` in a Claude Code install). Do NOT assume it is elsewhere; if your install path for this skill is X, then `template-write` is at `X/../template-write/`.

`read_documents` is for **reading** (understanding what's there). `template-write` is for **writing** (putting content in). Don't blur the two.

---

All files live under `~/.config/opencode/skills/read_documents/`. Invoke via:

```bash
# workdir = <skill 目录>（~/.config/opencode/skills/read_documents/）
python -m extract_files \
  --paths-file "<PROJECT>/temp/tmp_targets.txt"
```

> **SKILL_DIR:** adjust if installed elsewhere. Use **absolute paths** for `--paths-file` and `--assets-dir` so output lands in the project directory.
>
> **Shell compatibility:** the command above is identical in bash and PowerShell — run it with the working directory set to the skill folder (the `workdir` parameter), NOT via `cd ... && ...`. **PowerShell 5.1 does not support `&&` or `$(pwd)`**; always pass absolute paths instead of `$_P`-style variables. All examples below follow this form.

## Format support

| Format  | Text                         | Assets                             | Notes                   |
| ------- | ---------------------------- | ---------------------------------- | ----------------------- |
| `.docx` | `python-docx`                | `word/media/` + `word/embeddings/` | Pure Python (no Office) |
| `.doc`  | `python-docx` -> COM convert | ZIP/XML (converted)                | COM needs Office (Win)  |
| `.pptx` | `python-pptx`                | `ppt/media/` + `ppt/embeddings/`   | Pure Python (no Office) |
| `.ppt`  | `python-pptx` -> COM convert | ZIP/XML (converted)                | COM needs Office (Win)  |
| `.pdf`  | `PyMuPDF`                    | PyMuPDF per-page                   | Pure Python (no Office) |
| `.xlsx` | `openpyxl`                   | `xl/media/`                        | Pure Python (no Office) |
| `.xls`  | `xlrd` -> COM convert        | ZIP/XML (converted)                | COM needs Office (Win)  |

Encrypted ZIP entries are noted but never block extraction.

---

## Section extraction (`--section`)

Use `--section KEYWORD` to extract only matching sections — **saves output tokens**.

| Format           | Boundary detection                                                        |
| ---------------- | ------------------------------------------------------------------------- |
| `.docx`          | Heading styles + Chinese patterns -> markdown headings -> filter to match |
| `.doc`           | Same as docx (fallback = full text)                                       |
| `.pptx`          | Slide markers -> matches + 1 adjacent                                     |
| `.ppt`           | Same as pptx (fallback = full text)                                       |
| `.pdf`           | Page markers -> matches + 1 adjacent                                      |
| `.xlsx` / `.xls` | Sheet markers whose name/content matches                                  |

```bash
python -m extract_files \
  --paths-file "<PROJECT>/temp/tmp_targets.txt" \
  --section "<KEYWORD>"
```

---

## Complex page rendering (PDF only)

PDF pages with complex layouts (flowcharts, diagrams, multi-column graphics) produce garbled text. Two layers handle this:

### Layer 1: Auto-detection (code heuristic)

When `--assets-dir` is provided, pages with high text fragmentation (blocks>10, ≥40% short blocks <5 chars, x-spread >50% page width, no tables) are **automatically rendered as PNG** at 150 DPI. The text output shows:

```txt
--- Page 11 ---

[此页为复杂排版，已渲染为图片（图片: pages/page_11.png）]
[meta] blocks=20 avg_len=12 short=8 tables=0 x_spread=83%
```

Pages with tables are **never** auto-rendered — structured table data is more valuable than an image. The x-spread condition excludes directory/index pages (narrow column) from false positives.

### Layer 2: AI on-demand (`--render-page`)

Every page includes `[meta]` metadata at the end. If you (the AI) judge a page needs rendering that Layer 1 missed, re-run with `--render-page`:

```bash
python -m extract_files \
  --paths-file "<PROJECT>/temp/tmp_targets.txt" \
  --assets-dir "<PROJECT>/temp/assets" \
  --render-page 7 14
```

This forces the specified pages (1-based) to render as images, **overriding** the table exclusion. Requires `--assets-dir`.

### Meta format

```txt
[meta] blocks=N avg_len=M short=K tables=T x_spread=PP%
```

| Field      | Meaning                                     |
| ---------- | ------------------------------------------- |
| `blocks`   | Merged text block count                     |
| `avg_len`  | Average text length per block               |
| `short`    | Blocks with <5 chars (fragmentation)        |
| `tables`   | Table count on this page                    |
| `x_spread` | Text x-coordinate spread as % of page width |

---

## Modes

| Situation                 | Mode                     | Flags                                    |
| ------------------------- | ------------------------ | ---------------------------------------- |
| Search files by keyword   | **SEARCH**               | `--root` + `--glob`                      |
| Read specific known paths | **DIRECT**               | `--paths-file`                           |
| Read only a section       | **DIRECT + --section**   | `--paths-file` + `--section "<KEYWORD>"` |
| Just list files           | **SEARCH + --list-only** | `--root` + `--glob` + `--list-only`      |

### ⚠️ Anti-patterns

1. **Don't Glob + concat + DIRECT.** Use SEARCH mode (`--root` + `--glob`). The built-in Glob tool may return relative/truncated paths.
2. **Don't invoke once per file.** Write all paths into one `--paths-file` and make one call. Avoids redundant Python startup.

---

## SEARCH mode

When you need to **find** files by keyword. `--root` must be ASCII-safe; Chinese matching via `--glob`:

```bash
python -m extract_files \
  --root "<ASCII-safe ancestor>" \
  --glob "<Chinese keyword>"
```

For listing only: add `--list-only` (output to stdout — do NOT redirect to a file, bash corrupts Chinese).

---

## DIRECT mode

When you **already know the exact path**. Write a paths file with the **Write tool** (UTF-8):

```txt
Path: temp/tmp_targets.txt
Content:
<absolute path to document>
```

> Write the paths file in the project directory's **temp/** subdirectory.

```bash
python -m extract_files \
  --paths-file "<PROJECT>/temp/tmp_targets.txt"
```

---

## After extraction

- **DIRECT mode**: output is at `temp/tmp_output.txt` (same directory as paths file) — read it, then `rm -rf temp/`
- **SEARCH mode**: output path is printed as `OUTPUT_PATH:` on stderr — read it, then clean up accordingly
- Keep the `--assets-dir` output — that's user data

---

## Parameters

| Flag                      | Mode   | Description                                          |
| ------------------------- | ------ | ---------------------------------------------------- |
| `--root PATH`             | SEARCH | ASCII-safe ancestor directory                        |
| `--glob STR`              | SEARCH | Case-insensitive substring match                     |
| `--paths-file PATH`       | DIRECT | UTF-8 file, one path per line                        |
| `--section STR`           | both   | Only extract matching section                        |
| `--list-only`             | SEARCH | List files, skip extraction                          |
| `--assets-dir PATH`       | both   | Extract assets (use absolute path)                   |
| `--render-page N [N ...]` | DIRECT | Force-render PDF pages as images (1-based, PDF only) |
| `--max-depth N`           | SEARCH | Max directory depth (default 6)                      |
