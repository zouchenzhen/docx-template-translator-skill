---
name: docx-template-translator
description: Adaptive conversion of LaTeX, PDF, or Markdown sources into a complete Word .docx that follows a user-supplied .docx template. Use when pandoc --reference-doc alone is not enough — for thesis, dissertation, report, or institutional Word formatting that needs cover pages, declarations, TOC, heading numbering, captions, three-line tables, equations, citations, and visual verification.
---

# DOCX Template Translator

## Core Idea

Treat the input file as the content source and the Word template as the formatting source. Do not expect pandoc or PDF import to infer template semantics. Build a project-specific Python postprocessor after inspecting the template and the converted body document.

## Workflow

1. Identify inputs:
   - Source: `.tex` project, `.pdf`, `.md`, or an existing rough `.docx`.
   - Template: required `.docx`.
   - Output location and document metadata.
2. Inspect the template with `scripts/inspect_docx_template.py`.
3. Create a rough body `.docx`:
   - LaTeX/Markdown: use pandoc when available.
   - PDF: try Word COM import or `pdf2docx`; prefer PDF only when the original source is unavailable.
   - Existing DOCX: use it as the rough body source.
4. Write or patch a project-specific Python pipeline:
   - Start from `scripts/adaptive_docx_pipeline.py`.
   - Copy template front matter if needed.
   - Append rough body content while remapping DOCX relationships.
   - Remap styles to the template's real body, heading, caption, reference, and TOC styles.
   - Add or repair figure/table captions, table borders, hyperlinks, bookmarks, citations, and page breaks.
5. Finalize with Microsoft Word when available:
   - Use `scripts/finalize_word_docx.py` to update fields/TOC and export a PDF preview.
6. Visually verify:
   - Use `scripts/render_pdf_preview.py` to inspect cover pages, abstracts, TOC, representative tables, figures, formulas, and references.

## Required Engineering Rules

- Prefer deterministic Python and OOXML operations over manual Word edits.
- Preserve content first; only change formatting unless the user explicitly asks to edit text.
- Keep a generated PDF preview beside the final DOCX.
- For LaTeX, extract structured information from `.tex`, `.aux`, `.bbl`, `.toc`, and source captions when pandoc loses numbering or labels.
- For equations, preserve pandoc-generated OMML when possible; avoid touching paragraphs containing `m:oMath` unless necessary.
- For hyperlinks, explicitly set black/no-underline styling if the target template requires print-style links.
- For references, add bookmarks at bibliography entries before converting in-text numeric citations into internal hyperlinks.
- For institutional templates, avoid generic style names like `Body Text`; inspect the template because those names may be repurposed.
- For finalization, always run with macros disabled. The bundled `finalize_word_docx.py` sets `Word.Application.AutomationSecurity = msoAutomationSecurityForceDisable` before opening the document; do not loosen this for inputs of unknown provenance.
- Three-line table coercion is opt-in (`--three-line-tables` or `enable_three_line_tables` in config). Don't enable it unless the target template actually requires that layout.

## Script Guide

- `scripts/inspect_docx_template.py`: dumps template styles (including those defined but unused in the body), paragraphs, tables, section settings, numbering hints, and hyperlink colors.
- `scripts/adaptive_docx_pipeline.py`: reusable starter pipeline for template-based reconstruction. Behavior is config-driven (`--config`). Three-line tables and other Chinese-thesis-specific tweaks are opt-in.
- `scripts/finalize_word_docx.py`: updates Word fields/TOC and optionally exports PDF through Word COM. Disables macros for safety.
- `scripts/render_pdf_preview.py`: renders selected PDF pages into contact sheets for visual QA.
- `presets/zhengzhou_thesis.json`: example config that reproduces the Zhengzhou-University thesis behaviour described in the case study.

## References

- Read `references/pandoc-limitations.md` when explaining why this workflow differs from pandoc defaults.
- Read `references/zhengzhou-case-study.md` when building or adapting a university-thesis pipeline.
