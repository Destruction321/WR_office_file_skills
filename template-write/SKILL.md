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

- **Binary formats** (docx, xlsx, pptx) or non-ASCII paths: use the `read_documents` skill first. If that skill isn't available, use `python -m extract_files` from `read_documents/`.
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

### Step 3 — Fill content with the reusable package (preferred)

Use the `fill_template` package (in `~/.claude/skills/template-write/fill_template/`) to fill placeholders.
It preserves formatting better than ad-hoc scripts, and handles all binary formats consistently.

```bash
cd ~/.claude/skills/template-write && python -m fill_template \
  --template "<OUTPUT_PATH>" \
  --output "<OUTPUT_PATH>" \
  --set name=张三 \
  --set date="2025年3月"
```

Or use a JSON data file for more placeholders:

```bash
python -m fill_template \
  --template template.docx \
  --output filled.docx \
  --data-file content.json
```

**Supported by the package:**

| Format | Tool | Notes |
| ------ | ---- | ----- |
| `.docx` | `fill_template/docx_filler.py` | Run-level replacement — preserves bold/italic/font |
| `.xlsx` | `fill_template/xlsx_filler.py` | Cell-by-cell replacement |
| `.pptx` | `fill_template/pptx_filler.py` | Slide shape + table replacement |
| `.md` / `.txt` | `fill_template/text_filler.py` | UTF-8 BOM on Chinese Windows |
| `.csv` | `fill_template/text_filler.py` | Same as text |

> If the package is missing or broken, fall back to the manual Python script approach below.

### Step 3 (fallback) — Manual Python script

If the `fill_template` package cannot be used (e.g., imported but no CLI), write an inline script:

```bash
python -c "
from docx import Document
doc = Document(r'<OUTPUT_PATH>')
for p in doc.paragraphs:
    if '{{name}}' in p.text:
        # Replace at run level to preserve formatting
        for run in p.runs:
            run.text = run.text.replace('{{name}}', 'Replacement Content')
doc.save(r'<OUTPUT_PATH>')
"
```

> For multi-run placeholders that span across runs, merge all runs into the first run first.
> If the required Python package is missing, install it: `pip install python-docx`

### Step 4 — Verify

Read back the output file (or a summary) to confirm the content was written correctly. For binary formats, use `read_documents` skill if available; otherwise use Python extraction.

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
