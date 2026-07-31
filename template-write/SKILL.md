---
name: template-write
description: >
    Write reports or documents based on a known file template. Reads the template,
    then fills in content using the fill_template Python package. Preserves the
    original template untouched.
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

## ⛔ Read this first — the one rule that must never be broken

**NEVER write inline `python-docx` scripts. ALWAYS use `python -m fill_template`.**

Every time someone bypassed the package and hand-wrote `python-docx` scripts, it went wrong. Concretely, on a real task this caused:

- **Reversed content** — `insert_paragraph_before()` + `reversed()` put every section in backwards; took a manual 64-paragraph cleanup to fix.
- **Wrong styles** — hardcoded `Heading 2`/`Title` while the template actually used custom `a4`/`a3` styles -> inserted content looks nothing like the rest of the document.
- **Missing images** — inline scripts wrote `图6.1.1` as plain text instead of embedding the PNG. The whole point of the task was the images.
- **Fragile indices** — manually probing `doc.paragraphs[288]` breaks the moment anything is inserted earlier.
- **6× the tool calls** — 29 calls vs 5 when the package is used.

The package already handles all of this: copying, style-aware boundary detection, image embedding, ordered insertion, and auto-verification. If you find yourself typing `from docx import Document` in a `python -c` string, **stop** — you are about to repeat those failures. Use the package.

The ONLY exception: the package itself fails to import (`ModuleNotFoundError` that `pip install` can't fix). Even then, do not write a section-filling script — tell the user the package is broken.

---

## Core principle

**opencode is the brain, the tool is the hands.**

- opencode reads the template, understands the structure, decides what to fill and where.
- The tool (`python -m fill_template`) does only mechanical operations: locate, clear, insert, verify.
- If the tool can't find a heading, opencode adjusts the parameters — the tool never guesses.

## Supported formats

- **`.docx`** — Mode A (placeholder) + Mode B (section injection)
- **`.doc`** — same as `.docx`: the tool auto-converts to a `.docx` working copy via Word COM (Windows + Microsoft Office required), then runs the normal pipeline. The original `.doc` is never modified. Output defaults to `.docx`; name the output `.doc` to round-trip back to the legacy format.
- **`.md` / `.txt`** — Mode A (placeholder) only
- `.xlsx`, `.pptx`, `.csv` — **not supported**

## Workflow (the only path)

### Step 1 — Scan the template

**Always start with `--scan`.** This shows the heading structure and custom styles — everything you need to fill correctly. Do NOT probe paragraph indices with inline scripts.

```bash
# workdir = <skill 目录>（~/.config/opencode/skills/template-write/）
python -m fill_template \
  --template "<TEMPLATE_PATH>" --scan
```

> **Shell compatibility:** run commands with the working directory set to the skill folder (the `workdir` parameter) — PowerShell 5.1 does not support `&&`, so do NOT chain with `cd ... && ...`. All examples below follow this form.

From the scan output, determine:

- **Custom heading style** (e.g. `a4`, `标题 1`) -> pass with `--heading-style`
- **Heading names** -> **substring match, case-insensitive**: `三、实习总结` already matches `三、实习总结（可另附纸）`; only body-direct paragraphs are searched — text inside table cells is NOT treated as headings
- **Duplicate heading names** -> use `"Parent / Child"` scoped syntax
- **Which sections are empty** `[EMPTY]` -> these need filling
- **Tables** `[空表格]` -> empty tables are candidate answer areas in Chinese form documents
- **Tables** `单元格内空段: N` -> the table cell contains N empty placeholder paragraphs — this is a **table-inside answer area** (e.g. `三、实习总结` heading followed by a signature table whose cell holds 41 blank paragraphs). Use `--section-mode cell` to fill INTO the cell (deletes the placeholders, keeps the signature)

To confirm a heading is a body paragraph (not inside a table cell): check that its text does NOT appear in the scan's `Tables` text list, and that `extract_files` output shows it on its own line (possibly with a `##` marker).

#### Interpreting scan output (decision table)

The scan output alone is enough to choose the fill mode. Do NOT probe the docx with inline scripts to "confirm" the structure — everything needed is already in the scan output.

| Scan output                        | Inference                                                     | Correct action                                                             |
| ---------------------------------- | ------------------------------------------------------------- | -------------------------------------------------------------------------- |
| Headings > 0, some `[EMPTY]`       | Standard template, clear boundaries                           | Mode B `replace`                                                           |
| Headings > 0, no `[EMPTY]`         | Sections already have content                                 | Mode B `append`, or fill elsewhere                                         |
| Headings: 0, no custom styles      | All paragraphs `Normal`, no boundaries                        | Mode B `append` only — `replace` deletes everything after the heading      |
| Headings: 0 + `[空表格]` in Tables | Answer area may be a table (Chinese forms)                    | Check body order via `extract_files`, then decide                          |
| Tables show `单元格内空段: N`      | Table-inside answer area (placeholder blanks)                 | Mode B `cell` — fills INTO the cell, deletes placeholders, keeps signature |
| Headings: 0 + no tables            | Flat form; use Mode A with `--pattern` matching existing text | Match real text, not `{{}}` placeholders                                   |

> **Info completeness**: structural info (headings, styles, empty sections, tables, body order) comes from exactly two commands — `--scan` (this skill) and `extract_files` (read_documents skill). Together their output covers paragraphs, tables, styles, empty sections and boundaries. If you feel the information is insufficient, re-read these two outputs and the mechanism notes below — do NOT write inline python-docx probing scripts.

#### Combined reading with `extract_files`

`--scan` answers "what headings / styles / tables exist"; `extract_files` answers "what is the body order" (e.g. `三、实习总结` -> 学生签字 table -> `四、实习评语`). Join the two to locate answer areas:

- Heading directly followed by a signature / 评语 table -> check the scan's `单元格内空段` hint: if the cell holds blank placeholder paragraphs, the answer area is **inside the cell** (use `cell` mode); otherwise it is the **paragraph space between heading and table** (where `append` inserts)
- Heading followed by an `[空表格]` -> candidate table answer area; use paragraph injection before it, or edit the template directly

#### How section boundaries work (why append vs replace vs cell)

- **`replace`**: clears the section's old content (from the heading to the next same-style heading = the boundary), then inserts new content.
- **No heading styles** (all `Normal`): boundary detection fails, `end=None` — `replace` would delete **everything after the heading, including later sections**. The tool refuses with the safety-guard error (see Troubleshooting).
- **`append`**: skips boundary detection — locates by heading text only, inserts right after the heading, clears nothing. Use it whenever the template has no styled headings.
- **`cell`**: table-inside answer area — locates the heading, finds the first following table cell containing empty paragraphs, deletes the placeholders, inserts content before the cell's terminal paragraph (e.g. the signature line). Keeps the cell structurally valid.
- **`--heading-style <style>`**: tells the tool which style defines section boundaries, restoring `replace` on documents with custom styles.

### Step 2 — Read the template for content requirements (if needed)

The scan shows structure but not what each section should contain. To understand requirements:

- **docx**: use the `read_documents` skill with `--section "<KEYWORD>"` to read only the relevant part (saves tokens). Do NOT read the full document unless necessary.

> **Location**: `read_documents` lives in the **same skills directory as this skill** — i.e. `~/.config/opencode/skills/read_documents/` (or `~/.claude/skills/read_documents/` in a Claude Code install). If your install path for this skill is X, then `read_documents` is at `X/../read_documents/`.

- **Plain text** (`.md`, `.txt`): use the Read tool directly.

> If the scan already shows enough (empty sections with obvious names), skip this step.
>
> **If the model has vision capability**: after reading the template text, also open each image file to examine its content. This lets you decide placement based on what the image actually shows rather than guessing from filenames. Images (paragraphs + pictures) cover ~95% of real-world section-injection use cases; more complex elements like tables or charts are best handled by editing the template directly.

### Step 3 — Fill

Both modes use: `--template <ORIGINAL> --output <NEW_FILE>`. The tool **copies automatically** — do NOT manually copy.

#### Mode A: Placeholder filling

When the template has `{{name}}` style placeholders:

```bash
python -m fill_template \
  --template "<TEMPLATE_PATH>" \
  --output "<OUTPUT_PATH>" \
  --set name=张三 --set date="2025年3月"
```

Or with a JSON file:

```bash
python -m fill_template \
  --template "<TEMPLATE_PATH>" \
  --output "<OUTPUT_PATH>" \
  --data-file content.json
```

**Placeholder patterns**: `{{placeholder}}`, `[placeholder]`, `<placeholder>`, `%placeholder%`, `___`. Custom via `--pattern`.

#### Mode B: Section injection (docx only)

When the template has empty sections under headings needing full content (paragraphs + images).

**3a.** Write content as a Markdown file — this is where opencode's intelligence goes:

> **Formatting rules:**
>
> - **Don't insert spaces** between Chinese and Latin/digit characters.
> - **Prefer paragraph breaks.** Separate logical points with blank lines (paragraph break, Enter in Word), rather than piling everything into one paragraph with only soft line breaks (Shift+Enter in Word).

```markdown
## 实验八 / 实验过程及分析

1. 首先打开**记事本**，输入以下内容：

   ![](C:/path/to/screenshot1.png)

2. 然后配置安全策略。

## 实验八 / 实验结果总结

本次实验成功验证了基本原理。
```

**3b.** Dry-run first (verifies all headings can be located, no file changes):

```bash
python -m fill_template \
  --template "<TEMPLATE_PATH>" \
  --output "<OUTPUT_PATH>" \
  --section-data-file "<TEMPLATE_DIR>/temp/sections.md" \
  --heading-style a4 \
  --dry-run
```

**3c.** Fill for real (add `--force` if the dry-run already created the output):

```bash
python -m fill_template \
  --template "<TEMPLATE_PATH>" \
  --output "<OUTPUT_PATH>" \
  --section-data-file "<TEMPLATE_DIR>/temp/sections.md" \
  --heading-style a4 \
  --force
```

> **`--heading-style`**: pass this whenever the scan shows custom heading styles (non-`Heading N`). This tells the tool which style name defines section boundaries.

#### Combining modes (A then B)

If the template has both placeholders and empty sections:

1. Run Mode A first -> filled copy
2. Run Mode B on the **filled copy** as `--template` -> final document

### Step 4 — Verify and clean up

1. **Auto-verify is sufficient.** The tool prints `[OK]` for each filled section with paragraph/image counts. Trust it — it re-opens the file and checks the actual content.
2. **Do NOT re-read with `read_documents`.** It wastes ~2500 tokens and provides no extra information. Only do so if auto-verify output is clearly suspicious (missing content that should be there).
3. Clean up: `rm -rf "<TEMPLATE_DIR>/temp/"`

## Markdown syntax reference (Mode B)

| Syntax                      | Effect                                          |
| --------------------------- | ----------------------------------------------- |
| `# ~ ###### heading`        | Section delimiter (matches docx heading text)   |
| `## Parent / Child`         | Scoped: find parent, then child within it       |
| `![](path)`                 | Image (absolute path)                           |
| `![](path){width=5.0}`      | Image with custom width in inches (default 5.5) |
| `**bold**`                  | Bold run                                        |
| `*italic*`                  | Italic run                                      |
| `1.` or `-` prefix          | List item (each rendered as separate paragraph) |
| Blank line                  | Paragraph separator                             |
| Consecutive non-blank lines | Merged into one paragraph                       |

## Key rules

1. **Never write inline `python-docx` scripts.** Use `python -m fill_template`. (See the top of this file for why.)
2. **Never modify the original template.** The tool copies automatically.
3. **Always `--scan` first.** Don't guess heading styles or probe paragraph indices manually.
4. **Pass `--heading-style`** when the scan shows custom styles — this is how the tool knows what defines a section boundary.
5. **Use scoped syntax** (`Parent / Child`) when heading names are duplicated across sections.
6. **The tool does not infer.** If it can't find a heading, check the scan output and adjust your parameters.
7. **Don't insert spaces between Chinese and Latin/digit characters.** Chinese typography does not use spaces between CJK and Latin script or digits.
8. **Prefer paragraph breaks over soft line breaks.** Separate logical points with blank lines (paragraph break, Enter), not soft line breaks (Shift+Enter) that pile multiple ideas into one paragraph.

## Troubleshooting

| Symptom                             | Cause                       | Fix                                                                      |
| ----------------------------------- | --------------------------- | ------------------------------------------------------------------------ |
| "replace 模式将清除其后 N 个元素"   | No style boundary detected  | Use `--section-mode append`, or pass `--heading-style` to set a boundary |
| ".doc 转换失败"                     | Not Windows / Word missing  | `.doc` needs Windows + Office; convert to `.docx` manually first         |
| "0 sections filled"                 | Heading text doesn't match  | Check scan output, use exact text                                        |
| Section fills into wrong location   | Duplicate heading names     | Use `Parent / Child` scoped syntax                                       |
| Content overflows into next section | Boundary not detected       | Add `--heading-style` for custom styles                                  |
| Style not recognized                | Custom style not passed     | Use `--heading-style <style>` from scan output                           |
| Tool crashes on fill                | Missing dependency          | `pip install python-docx`                                                |
| "文件已存在" error                  | Dry-run left an output file | Add `--force`                                                            |

## If the package is broken

If `python -m fill_template` fails with an import error that `pip install python-docx` cannot fix:

1. **Do NOT** write an inline section-filling script — it will reverse content, mismatch styles, and drop images (see top of file).
2. Tell the user the `fill_template` package is broken and needs repair.
3. Only for trivial single-placeholder replacement (Mode A, one `{{name}}`), and only with the user's explicit agreement, may you run a minimal inline script on an **already-copied** output file. Never use this for section filling or images.
