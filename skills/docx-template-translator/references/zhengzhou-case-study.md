# Zhengzhou University Thesis Case Study

This reference summarizes the proven pattern from a Zhengzhou University
LaTeX-thesis-to-Word conversion.

## Input

- Source: a multi-file LaTeX thesis project with `main.tex`, `data/*.tex`,
  `figures/*`, and BibTeX references.
- Template: `郑大毕业论文（设计）模板-V2.docx`.
- Goal: editable Word document matching the university template.

## What Failed

- Pure pandoc with `--reference-doc` produced structurally complete output, but
  many body paragraphs inherited the template's `Body Text` style.
- In that template, `Body Text` was actually used by the English cover page and
  rendered as large bold text, so normal thesis paragraphs looked wrong.
- LaTeX-generated captions and reference behavior required repair.

## What Worked

1. Convert LaTeX to rough DOCX with pandoc to preserve:
   - paragraphs/headings,
   - images,
   - tables,
   - OMML formulas,
   - bibliography entries.
2. Open the Word template with `python-docx`.
3. Keep/fill the template's native cover, English cover, originality statement,
   authorization statement, abstract headings, and TOC style.
4. Append the rough pandoc body while remapping image and hyperlink
   relationships in OOXML.
5. Remap body paragraphs from `Normal`/`Body Text` to the template's real
   thesis body style (`论文正文` in the case study).
6. Suppress automatic Heading 1 numbering for front/back matter headings.
7. Parse LaTeX `\caption{}` text and repair visible captions such as
   `图 3.1  本文主线系统架构图`.
8. Center actual image paragraphs and actual captions only; avoid centering
   in-text references like `图 4.4 与表 4.3 分别给出...`.
9. Format all tables as three-line tables by writing `w:tcBorders` directly.
10. Add bookmarks at reference entries and convert numeric citations into
    superscript internal hyperlinks.
11. Force hyperlinks to black/no underline for print-style thesis output.
12. Use Word COM to update TOC/fields and export a PDF preview.

## Libraries Used

- `pandoc`: rough LaTeX/Markdown to DOCX conversion.
- `python-docx`: high-level DOCX object model plus raw OOXML access.
- `pywin32`: Microsoft Word COM automation for TOC fields and PDF export.
- `PyMuPDF` (`fitz`) and Pillow: PDF preview contact sheets.

## Key Lesson

The robust unit of reuse is not a universal converter. The robust unit is a
repeatable workflow plus a starter Python pipeline that the AI patches for the
specific template and source project.
