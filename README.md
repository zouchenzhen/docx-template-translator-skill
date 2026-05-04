# DOCX Template Translator Skill

An open-source Codex skill for converting **LaTeX, PDF, Markdown, or rough DOCX**
sources into complete Word documents that follow a user-supplied `.docx`
template.

This project is designed for thesis, dissertation, institutional report, and
standards-document workflows where **pandoc's default DOCX output is not enough**.

中文版: [README.zh-CN.md](README.zh-CN.md)

## Why Not Just Pandoc?

Pandoc is excellent for first-pass conversion, especially from Markdown and
LaTeX-like sources. Its `--reference-doc` option can reuse Word styles, but it
does not understand the semantic meaning of a school's or institution's Word
template.

Common problems in strict template workflows:

- cover pages are not filled correctly;
- declaration/signature pages are missing or visually wrong;
- template style names are misleading;
- body paragraphs accidentally inherit cover-page styles;
- figure/table captions lose numbering;
- tables are not formatted as required three-line tables;
- citations are not superscripted or linked to references;
- hyperlinks appear blue/underlined in print-style documents;
- TOC fields and page numbers are not updated;
- no visual PDF verification is produced.

This skill uses pandoc, Word, and Python as building blocks, then lets the AI
write or patch a **project-specific postprocessing script** for the exact
template and source files.

## Related Projects and Prior Art

- [pandoc](https://pandoc.org/MANUAL.html) supports `--reference-doc` for DOCX
  style customization. Its own manual notes that the reference DOCX contents are
  ignored while styles and document properties are used.
- [Quarto DOCX](https://quarto.org/docs/reference/formats/docx.html) also
  supports `reference-doc`, TOC, numbering, citations, and bibliography linking.
- [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) converts native PDFs
  to DOCX; its documentation explains that it uses PyMuPDF for PDF extraction,
  rule-based layout parsing, and `python-docx` for DOCX generation.
- [python-docx-template](https://docxtpl.readthedocs.io/) is excellent for
  Jinja-style variable replacement inside a prepared Word template.
- [MDDoc](https://www.mddoc.app/) is a commercial Markdown-to-Word product that
  maps Markdown elements to uploaded Word templates.
- [wmvanvliet/pandoc-tutorial](https://github.com/wmvanvliet/pandoc-tutorial)
  demonstrates the practical reality that plain pandoc often gets most of the
  way for complex LaTeX-to-DOCX conversion, but the last part needs custom work.
- [openclaw/pdf-to-docx skill](https://playbooks.com/skills/openclaw/skills/pdf-to-docx)
  is a PDF-to-DOCX skill based on pdf2docx.

This project is different because it is an **open agent workflow** for strict
template reconstruction across LaTeX/PDF/Markdown inputs. It asks the AI to
inspect the uploaded template and produce a dedicated Python postprocessor,
instead of pretending one static converter can satisfy every institutional
template.

## Core Approach

1. Use the source file as the **content source**.
2. Use the `.docx` template as the **formatting source**.
3. Generate a rough body `.docx` with pandoc, Word import, or another converter.
4. Inspect the template styles, paragraphs, numbering, tables, and XML.
5. Let the AI write a Python pipeline adapted to that template.
6. Rebuild the final `.docx` with `python-docx` and raw OOXML operations.
7. Use Word COM to update fields/TOC and export a PDF preview.
8. Render preview contact sheets for quick visual QA.

## Supported Inputs

- LaTeX projects (`main.tex`, chapters, figures, BibTeX);
- Markdown files;
- born-digital PDFs;
- existing rough DOCX files.

PDF input is inherently less semantic than LaTeX or Markdown. Use source files
when available.

## Included Skill

The Codex skill lives at:

```text
skills/docx-template-translator/
```

It includes:

- `SKILL.md`: the agent workflow;
- `scripts/inspect_docx_template.py`: template inspection;
- `scripts/adaptive_docx_pipeline.py`: starter reconstruction pipeline;
- `scripts/finalize_word_docx.py`: Word field/TOC update and PDF export;
- `scripts/render_pdf_preview.py`: PDF preview contact sheet rendering;
- `references/pandoc-limitations.md`: comparison with pandoc defaults;
- `references/zhengzhou-case-study.md`: university-thesis case study.

## Installation

Copy the skill folder into your Codex skills directory:

```bash
mkdir -p ~/.codex/skills
cp -r skills/docx-template-translator ~/.codex/skills/
```

Then ask Codex:

```text
Use $docx-template-translator to convert my LaTeX thesis into Word using this .docx template.
```

## Python Dependencies

Recommended:

```bash
pip install python-docx pywin32 pymupdf pillow
```

Optional but useful:

```bash
pip install pdf2docx
```

Install pandoc separately if converting LaTeX or Markdown.

For best finalization, use Windows with Microsoft Word installed. Word COM is
used to update TOC/page fields and export the final PDF preview.

## Typical Usage

1. Provide the source project and the Word template.
2. Let Codex run:

```bash
python scripts/inspect_docx_template.py template.docx --out template_report.json
```

3. Let Codex create a rough body document:

```bash
pandoc main.tex --citeproc --reference-doc template.docx -o body.docx
```

4. Let Codex adapt `scripts/adaptive_docx_pipeline.py` for the concrete
template.
5. Finalize:

```bash
python scripts/finalize_word_docx.py final.docx --pdf
python scripts/render_pdf_preview.py final.pdf --pages 1-8
```

## Project Status

This is a workflow skill, not a one-click universal converter. The central idea
is that every strict Word template has local rules, so the AI should inspect the
template and generate a dedicated postprocessor.

## Community

[LINUX DO — 中文开发者社区](https://linux.do/)

This project recognizes and appreciates LINUX DO as a Chinese developer
community for open-source sharing and technical discussion. This acknowledgment
is not a claim of official endorsement unless separately stated by the
community.

## License

MIT
