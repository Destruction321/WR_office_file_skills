---
name: template-write
description: >
    Write reports or documents based on a known file template. Reads the template,
    copies it to a new file in the same directory, then fills in content in the
    copy. Preserves the original template untouched.
triggers:
    - keywords: "模板,template,填写模板,按模板写,写报告,写文件,用模板,生成报告"
    - descriptionMatches:
        - "write.*template"
        - "fill.*template"
        - "report.*template"
        - "按.*模板"
        - "根据.*模板"
tools:
    - read
    - bash
    - write
---

# Template Write — Template-Based Document Writing

## When to invoke

Invoke when the user provides a **template file path** and **content instructions** (what to fill in).

## Workflow

### Step 0 — Validate template path

Check the template exists before proceeding:

```bash
test -f "<TEMPLATE_PATH>" && echo "exists" || echo "not found"
```

If it doesn't exist, report the error and ask the user for the correct path.

### Step 1 — Read the template

Read the template to understand its structure, placeholders, headings, and formatting.

**Reading approach:**

- **Binary formats** (docx, xlsx, pptx) or non-ASCII paths: try the `read_documents` skill first. If that skill isn't available, use a Python script (e.g. `python-docx`) to extract text content.
- **Plain text** (`.md`, `.txt`, `.csv`): use the Read tool directly.

**Common placeholder patterns to look for:**

- `{{placeholder}}` / `{{ placeholder }}` — double-brace (Jinja/Handlebars style)
- `[placeholder]` — single-bracket
- `<placeholder>` — angle-bracket
- `%placeholder%` — percent-wrapped
- `___` or `______` — underlined blanks
- Highlighted / colored text — formatting-based markers
- Comment text or yellow-highlighted runs in docx — annotation-style markers
- Table cells with a single default value or empty — form-style fill-in cells

> No placeholder found? The user likely wants to write content into each section directly — use the template's structure (headings, table rows, bullet lists) as the skeleton.

### Step 2 — Copy the template

Copy the template to a new file in the **same directory**:

```bash
cp -n "<TEMPLATE_PATH>" "<TEMPLATE_DIR>/<BASENAME>_<CONTENT_SUFFIX>.<EXT>"
```

> The original template must remain **untouched**. All edits go into the copy.
>
> Example: `~/docs/周报模板.docx` → `~/docs/周报模板_2025年3月周报.docx`

If the output file already exists (`cp -n` skips without overwriting), ask the user: overwrite, or use a different suffix?

### Step 3 — Write content into the copy

After reading the template structure, fill in the content according to format:

| Format | Tool | Notes |
| ------ | ---- | ----- |
| `.docx` | `python-docx` | Fill paragraphs and table cells by matching placeholders or structure. For merged/nested tables, access cells via `table.cell(row, col).text`. Only modify targeted text — python-docx preserves everything else automatically. |
| `.xlsx` | `openpyxl` | Fill by cell reference or named range. Handle merged cells — unmerge only when unavoidable. |
| `.pptx` | `python-pptx` | Fill text in slide placeholders (`slide.placeholders[idx]`) or specific shapes. |
| `.md` / `.txt` | Write tool | Use UTF-8 encoding. On Chinese Windows, if system locale is GBK, write UTF-8 with BOM explicitly to avoid garbled text. |
| `.csv` | Write tool | Use UTF-8. Respect locale-appropriate delimiter: `,` (most locales) or `;` (some European/Asian Excel). |

**Python execution:**

- **Command**: use `python` (works on both Windows and Linux). If `python` is not found, fall back to `python3`.
- **Simple fill**: inline script via `python -c "..."` is fine. Watch out for quoting — if the content contains `"""` or complex escapes, write a temp `.py` file instead.
- **Complex fill** (table manipulation, merged cells, multiple passes): write a `.py` file, review it, run it, then clean up.

```bash
python -c "
from docx import Document
doc = Document(r'<OUTPUT_PATH>')
for p in doc.paragraphs:
    if '{{name}}' in p.text:
        p.text = p.text.replace('{{name}}', 'Replacement Content')
doc.save(r'<OUTPUT_PATH>')
"
```

> If the required Python package is missing, install it first:
>
> ```bash
> pip install python-docx  # or openpyxl, python-pptx
> ```

### Step 4 — Verify

Read back the output file (or a summary) to confirm the content was written correctly. For binary formats, use `read_documents` skill if available; otherwise use a Python extraction script.

## Multiple templates in the same directory

If the user says "fill in a template" but doesn't specify which one, and there are multiple template files:

1. **List candidates** — find files with common template extensions (`.docx`, `.xlsx`, `.pptx`, `.md`, `.txt`, `.csv`)
2. **Ask the user** which one to use (or infer from name if one clearly matches the description)
3. **Proceed** with the selected template

## Important rules

1. **Never modify the original template.** Always work on a copy.
2. **Default output: same directory as the template.** The output filename appends a descriptive content suffix to the template's basename.
3. **Only modify the content being filled.** Elements you don't touch (headers, footers, images, TOC, etc.) are preserved automatically by the library — no extra protection needed.
4. **Match the template structure.** Identify placeholders, blanks, or marked sections and fill them according to the user's instructions.
