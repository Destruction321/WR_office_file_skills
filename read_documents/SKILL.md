---
name: read_documents
description: >
    Read document files (doc, docx, ppt, pptx, xls, xlsx, pdf) when file paths
    contain Chinese or other non-ASCII characters. Avoids shell encoding corruption
    and OPC library encoding bugs. Extracts content to current working directory
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

---

All files live under `~/.claude/skills/read_documents/`. Invoke via:

```bash
_P="$(pwd)" && cd ~/.claude/skills/read_documents && python -m extract_files \
  --paths-file "$_P/temp/tmp_targets.txt"
```

> **SKILL_DIR:** adjust if installed elsewhere. Use **absolute paths** for `--paths-file` and `--assets-dir` so output lands in the project directory.

## Format support

| Format  | Text                         | Assets                             | Notes                   |
| ------- | ---------------------------- | ---------------------------------- | ----------------------- |
| `.docx` | `python-docx`                | `word/media/` + `word/embeddings/` | Safe-path for non-ASCII |
| `.doc`  | `python-docx` → COM fallback | ZIP/XML or **NO**                  | COM needs Office (Win)  |
| `.pptx` | `python-pptx`                | `ppt/media/` + `ppt/embeddings/`   | Safe-path for non-ASCII |
| `.ppt`  | COM via PowerShell           | **NO**                             | COM needs Office (Win)  |
| `.pdf`  | `pdfplumber` → `PyPDF2`      | PyMuPDF per-page                   | Safe-path for non-ASCII |
| `.xlsx` | `openpyxl`                   | `xl/media/`                        | Safe-path for non-ASCII |
| `.xls`  | `xlrd` → COM fallback        | **NO**                             | COM needs Office (Win)  |

Encrypted ZIP entries are noted but never block extraction.

---

## Section extraction (`--section`)

Use `--section KEYWORD` to extract only matching sections — **saves output tokens**.

| Format           | Boundary detection                                                      |
| ---------------- | ----------------------------------------------------------------------- |
| `.docx`          | Heading styles + Chinese patterns → markdown headings → filter to match |
| `.doc`           | Same as docx (fallback = full text)                                     |
| `.pptx`          | Slide markers → matches + 1 adjacent                                    |
| `.ppt`           | Same as pptx (fallback = full text)                                     |
| `.pdf`           | Page markers → matches + 1 adjacent                                     |
| `.xlsx` / `.xls` | Sheet markers whose name/content matches                                |

```bash
_P="$(pwd)" && cd ~/.claude/skills/read_documents && python -m extract_files \
  --paths-file "$_P/temp/tmp_targets.txt" \
  --section "<KEYWORD>"
```

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
_P="$(pwd)" && cd ~/.claude/skills/read_documents && python -m extract_files \
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
_P="$(pwd)" && cd ~/.claude/skills/read_documents && python -m extract_files \
  --paths-file "$_P/temp/tmp_targets.txt"
```

---

## After extraction

- **DIRECT mode**: output is at `temp/tmp_output.txt` (same directory as paths file) — read it, then `rm -rf temp/`
- **SEARCH mode**: output path is printed as `OUTPUT_PATH:` on stderr — read it, then clean up accordingly
- Keep the `--assets-dir` output — that's user data

---

## Parameters

| Flag                | Mode   | Description                                  |
| ------------------- | ------ | -------------------------------------------- |
| `--root PATH`       | SEARCH | ASCII-safe ancestor directory                |
| `--glob STR`        | SEARCH | Case-insensitive substring match             |
| `--paths-file PATH` | DIRECT | UTF-8 file, one path per line                |
| `--section STR`     | both   | Only extract matching section                |
| `--list-only`       | SEARCH | List files, skip extraction                  |
| `--assets-dir PATH` | both   | Extract assets (use `$_P/...` absolute path) |
| `--max-depth N`     | SEARCH | Max directory depth (default 6)              |
