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

## Content-audit gaps found after the v5 → v6 round

The v5 retrofit fixed the five render-level bugs above, and the new render
validator declared `PASS`. The user then pointed out two more failures that
neither validator was looking for:

- **3 figures missing.** `chap01.tex`'s `\includegraphics{thesis_structure}`,
  `chap02.tex`'s `\includegraphics{q_learning_algorithm}`, and `chap03.tex`'s
  `\includegraphics{simulation_script_framework}` all reference the basename
  without `.png`. Pandoc resolved each basename to the vector sibling
  (`.pdf` or `.vsdx`) and could not embed it, so all three figures dropped
  silently. The structural validator counted what *was* embedded (20 unique
  image rels, 24 drawings via reuse) and declared the count "fine".
  *Mitigations:* run `validate_docx_render.py --source-latex-dir
  zzuthesis/data` to assert source `\includegraphics` count ≤ rendered
  `<w:drawing>` count; or pre-process the LaTeX to add `.png` to every
  bare-basename `\includegraphics`.
- **Three-line tables never applied.** The skill ships
  `format_three_line_tables` in `adaptive_docx_pipeline.py` and the preset
  carries `enable_three_line_tables: true`, but the project pipeline forgot
  to call the function. All 20 tables ended up borderless. Even calling the
  function once isn't enough: it iterates `table.rows[-1].cells`, and
  python-docx's wrapper for a `vMerge="continue"` cell collapses onto the
  merge-start cell above, so the bottom heavy line gets a gap at the
  bottom of any column whose final cell is a vertical-merge continuation.
  *Mitigations:* run `validate_docx_render.py
  --expected-table-style three-line`; and when applying borders, fall back
  to writing `<w:tcBorders>` directly into the raw `<w:tc>` elements inside
  the last `<w:tr>` rather than relying on python-docx's cell wrappers.

The takeaway: a user's bug list is a **sample**, not a catalog. The render
validator must also model "what the source said should be in the output" —
figure count, caption count, table count, table border style — not only "is
the structure internally consistent". Both kinds of checks must run before
declaring a conversion complete.

## Citation + caption gaps found after the v6 → v7 round

After v6 fixed the figure / three-line problems, the user surfaced two more
defects that *all* validators (structural + the v6 render validator) had
declared PASS on:

- **Citations rendered as `(Author Year)`, not GB/T-7714 `[N]` superscript.**
  Pandoc's default cite-rendering produces tokens like `(Zhu et al. 2025)` /
  `(Han et al. 2024; Zhu et al. 2025)`. They are **visible**, so a paragraph
  count or reference-bookmark count cannot tell anything is wrong. The thesis
  shipped with 33 ref bookmarks (`ref_1`…`ref_33`) but only 10 in-text
  `<w:hyperlink w:anchor="ref_*">` and **0** `<w:vertAlign w:val="superscript">`
  runs. *Mitigations:* (1) parse `main.bbl` for the `\bibitem{key}` order
  (== bibliography number) and the bracketed opt's `(first_surname, year)`;
  (2) walk every body paragraph in `document.xml` for `(...)` cite tokens;
  (3) split multi-cite tokens by `;`, resolve each to a cite_key (handle
  collisions like two `lozano-cuadra 2024` entries by matching the bbl opt's
  secondary author list **plus the first author's initial from the bbl
  author line** — pandoc emits `(W. Lu et al. 2025)` vs `(J. Lu et al. 2025)`
  when two cite_keys share the same surname-year, and `(Lozano-Cuadra and
  Soret 2024)` vs `(Lozano-Cuadra et al. 2024)` is distinguishable by
  whether the secondary surname `Soret` actually appears in the docx token.
  Use word-boundary matching `\bX\b` for single-letter initials so `'j'` does
  not accidentally match the letter inside `et al.`); (4) replace the
  original parenthetical text with `[N1, N2]` wrapped as
  `<w:hyperlink w:anchor="ref_N">…<w:vertAlign w:val="superscript"/>` runs.
  Detected by the new `citation-coverage` check (`--source-latex-dir`).
- **Figure / table captions missing chapter-relative numbers and not
  centered.** Pandoc's caption handling for Chinese theses dropped most
  captions onto a plain body style without the `图 X.Y` / `表 X.Y` prefix;
  some captions disappeared entirely along with their `\begin{table}` env
  when the table was too complex for pandoc to render as `<w:tbl>`. The v6
  thesis ended up with 5 caption paragraphs total (out of 30 source caption
  envs), 0 of them centered. *Mitigations:* walk the docx body chapter by
  chapter, group adjacent `<w:drawing>` paragraphs as a single figure
  (subfigure pattern), assign `图 chap.idx_in_chap` / `表 chap.idx_in_chap`
  numbers, pull caption text from `\begin{figure}…\caption{…}` /
  `\begin{table}…\caption{…}` in source `chapNN.tex` in document order, and
  insert a centered caption paragraph after each figure-group / before each
  table. Detected by the new `caption-count-vs-source` and `caption-format`
  checks. Captions whose source `\begin{table}` env was dropped by pandoc
  remain a `caption-count-vs-source` FAIL — those tables must be rebuilt
  from `\begin{tabular}` separately, not masked by inserting an empty
  caption.

The pattern repeating across v4 / v5 / v6 / v7 rounds: every iteration the
validator catches everything *the validator looks for*, the user spots
something *new*, the new failure becomes a new validator check, and the
checklist grows. Each round of "STATUS: PASS" was honest within its own
scope; what was missing was a piece of the model — what the source said the
output should contain. Citations and captions are the two largest categories
of "source declares it, validator never asked, pandoc dropped it" — fixing
the validator gap is what makes future runs catch the same defect without
another reviewer round.

## Libraries Used

- `pandoc`: rough LaTeX/Markdown to DOCX conversion.
- `python-docx`: high-level DOCX object model plus raw OOXML access.
- `pywin32`: Microsoft Word COM automation for TOC fields and PDF export.
- `PyMuPDF` (`fitz`) and Pillow: PDF preview contact sheets.

## Key Lesson

The robust unit of reuse is not a universal converter. The robust unit is a
repeatable workflow plus a starter Python pipeline that the AI patches for the
specific template and source project.
