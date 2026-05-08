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
- Running the starter `adaptive_docx_pipeline.py` unchanged is also wrong for
  this template: it keeps the template's sample cover/abstract/chapters and
  appends the real thesis body after the sample back matter. That output must be
  treated as a failed smoke test, not as a partial final document.
- Copying raw pandoc body XML into this template can trigger style ID collisions:
  visible `Heading 1/2/3` paragraphs may be interpreted as the template's body
  style if style IDs are not remapped by style name before insertion. This breaks
  Word TOC generation even when the visible text looks correct.
- After deleting sample sections, the remaining `sectPr` can still reference the
  last sample header, producing a stale body header such as `致谢`. Header and
  footer references must be inspected and cleared or rebuilt.
- A pipeline can still start from `Document(template)` and corrupt the template:
  deleting/recreating runs in cover placeholders or running a whole-document
  body-style pass can turn native cover/declaration paragraphs into `论文正文`.
  Content checks may pass while the two covers and originality/signature pages
  lose their formatting.

## What Worked

1. Convert LaTeX to rough DOCX with pandoc to preserve:
   - paragraphs/headings,
   - images,
   - tables,
   - OMML formulas,
   - bibliography entries.
2. Open the Word template with `python-docx`.
3. Build a project-specific pipeline from the starter script, saved in the run
   or project output directory.
4. Keep/fill the template's native cover, English cover, originality statement,
   authorization statement, abstract headings, and TOC style. Treat the cover,
   declaration, authorization, and signature pages as protected regions.
5. Delete template-only sample content, including sample abstracts, sample
   chapters, sample references, sample appendices, and sample acknowledgements.
6. Insert the real rough pandoc body at the body start, or rebuild the document
   around selected template front matter and section settings. Do not append the
   real body to the end of the complete template.
7. Remap image and hyperlink relationships in OOXML when moving body content.
8. Remap copied style IDs (`w:pStyle`, `w:rStyle`, `w:tblStyle`) by visible
   style name before appending, so source `Heading 1/2/3` keep their heading
   semantics inside the template.
9. Remap body paragraphs from `Normal`/`Body Text` to the template's real
   thesis body style (`论文正文` in the case study), but only inside the generated
   abstract/body/back-matter scope. Do not apply this pass to the native cover
   or declaration pages.
10. Clear or rebuild inherited section header/footer references after deleting
    template sample pages; verify the PDF does not show `致谢` or another
    back-matter header on abstract/body pages.
11. Suppress automatic Heading 1 numbering for front/back matter headings.
12. Parse LaTeX `\caption{}` text and repair visible captions such as
   `图 3.1  本文主线系统架构图`.
13. Center actual image paragraphs and actual captions only; avoid centering
   in-text references like `图 4.4 与表 4.3 分别给出...`.
14. Format all tables as three-line tables by writing `w:tcBorders` directly.
15. Add bookmarks at reference entries and convert numeric citations into
    superscript internal hyperlinks.
16. Force hyperlinks to black/no underline for print-style thesis output.
17. Use Word COM to update TOC/fields and export a PDF preview.
18. Validate protected front matter against the original template, using
    `validate_docx_conversion.py --template ... --protected-until "中 文 摘 要"`.
    Metadata text may differ, but paragraph styles, run-level fonts/sizes/bold,
    alignment, spacing, and page-break structure must not drift.

## Required QA Checks

Treat any of the following as FAIL and iterate before reporting completion:

- The generated document still contains `李四`, `王五`, `张三`, red template
  instructions, or sample chapter titles from the template.
- The real thesis text appears after `致谢`, sample appendices, or sample
  references.
- Body pages inherit a back-matter header such as `致谢`.
- Source chapter headings are not real Word heading styles after reconstruction;
  this usually means style IDs were copied without remapping.
- The TOC points to the template sample chapters instead of the real source
  chapters.
- Chinese/English abstracts or keywords come from the template instead of the
  source `.tex` files.
- Cover pages, English cover pages, originality/declaration pages, authorization
  pages, or signature/date lines change from their template styles to `论文正文`
  or otherwise lose run-level font/size/bold formatting.
- A project pipeline uses a function like `set_paragraph_text()` that removes
  all runs in protected front matter, or an `apply_final_styles()` pass that
  iterates over every paragraph/table without a generated-content scope marker.

## Render-level QA (added after the v4 → v5 retrofit)

The structural validator returned `STATUS: PASS` while the rendered document
still had five visible defects: chapters had no `第N章` prefix, sections
collapsed to `[1]`, references started at `[47]…[79]`, the body running
header said `致谢`, and the 目录 page contained only the heading with no
TOC body. Add these render-level checks (all available in
`validate_docx_render.py`) to the iteration loop:

- **TOC field present.** Walk `word/document.xml` for `<w:fldChar
  fldCharType="begin">` followed by `<w:instrText> TOC `. If absent, run
  `scripts/inject_toc_field.py final.docx --in-place` before finalization.
- **`numId=1` is bound to the multilevel `abstractNum`** whose level 0 is
  `第%1章`, level 1 is `%1.%2`, level 2 is `%1.%2.%3`. The previous
  `repair_reference_numbering_links.py` step in v4 silently re-pointed
  `numId=1` at a single-level `[%1]` abstract while building the reference
  list — the heading level format strings then degraded to `[1]` for both
  H2 and H3 paragraphs.
- **Heading 1 actually carries `numPr`**, either at the style level or via
  inline paragraph properties on every body chapter heading. Without it the
  chapter prefix never renders.
- **References use a dedicated `numId`** (e.g. `numId=4`) bound to the
  `[%1]` `abstractNum`. Do not let the references share `numId=1` with
  Heading 2/3 — the counter accumulates through every preceding heading,
  which is why 33 references rendered as `[47]…[79]`.
- **Body header is dynamic.** When the body+back-matter live in a single
  section, replace the inherited static `致谢` text with a `STYLEREF 1
  \* MERGEFORMAT` field via `scripts/set_styleref_header.py final.docx
  --header headerN.xml --style-id 1 --in-place`. Use the **numeric
  styleId**, not the localized display name `Heading 1`, otherwise Word
  prints `错误!使用'开始'选项卡…` in the header.
- **Body section restarts page numbering at 1.** Add `<w:pgNumType
  w:fmt="decimal" w:start="1"/>` to the body `sectPr`; otherwise the body
  silently inherits the abstract's Roman numerals and you see
  "目录 page footer = 5 / 第1章 page footer = 8".

## Libraries Used

- `pandoc`: rough LaTeX/Markdown to DOCX conversion.
- `python-docx`: high-level DOCX object model plus raw OOXML access.
- `pywin32`: Microsoft Word COM automation for TOC fields and PDF export.
- `PyMuPDF` (`fitz`) and Pillow: PDF preview contact sheets.

## Key Lesson

The robust unit of reuse is not a universal converter. The robust unit is a
repeatable workflow plus a starter Python pipeline that the AI patches for the
specific template and source project.
