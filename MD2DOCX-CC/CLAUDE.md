# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

A single-file Python script (`md2docx.py`) that converts Markdown legal documents into formatted Word (.docx) files. It uses a template docx for page layout/styles and creates all Word numbering definitions from scratch at runtime.

## Running

```bash
python3 md2docx.py <input.md> <output.docx> [template.docx]
```

Template defaults to `Template.docx` in the same directory as the script. Only dependency is `python-docx` (`pip install python-docx`).

## Architecture

The script has one main pipeline: **parse MD line-by-line → emit docx paragraphs with formatting**.

**Numbering system** (`_setup_numbering`): Creates 4 `abstractNum` definitions in the docx's `numbering.xml` at runtime, avoiding any dependency on template-specific numIds:
- Chinese counting (一、二、三…) for `##` section headings
- Decimal multilevel (1. / 1.1) for `###` subsections and ordered sub-items
- Bullet (● ○ ■) for unordered lists
- Decimal bold (1. / 1.1) for standalone ordered lists — cloned per list group via `NumConfig.create_ordered_num()` to reset counters

**Heading text extraction**: Uses `_HEADING_RE` regex (not `lstrip('#')`) to safely parse ATX headings. `_strip_heading_prefix` removes manual numbering prefixes (中文 or Arabic) since Word auto-numbering replaces them.

**List item classification** (the trickiest part): A `*` bullet starting with `**label：**` pattern becomes an ordered sub-item (subsection numbering ilvl=1); otherwise it becomes an unordered bullet. Indented sub-bullets (`  *` or `\t*`) under numbered items share the parent's numId at ilvl=1.

**Ordered list continuity**: The parser tracks `in_ordered` state. Blank lines between numbered items don't break the list — only headings or body paragraphs do. Each new list group gets a fresh numId via `create_ordered_num()`.

## Key Design Decisions

- Template provides only page setup, margins, headers/footers, and base paragraph styles (Normal, List Paragraph). All numbering XML is script-generated.
- `<span>` and other HTML tags are kept as plain text (not converted).
- `<br>` is only meaningful inside table cells (splits into multiple cell paragraphs).
- Tables get explicit border XML rather than relying on a named table style existing in the template.
