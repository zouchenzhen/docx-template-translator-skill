# expected/

This directory holds the **reference outputs** for the minimal_markdown
example, so a reader can see the produced look-and-feel without having to set
up Word + pywin32 locally.

## preview.png

A 2×2 contact-sheet of `final.pdf` produced on a Windows machine with
Microsoft Word. The committed image is **deliberately** taken from real body
pages — not the front-matter — so it actually demonstrates what the skill
does to real content. Specifically the four pages are:

- **Page 7** — Table of contents: shows that headings, page numbers and dot
  leaders are correctly resolved by Word after `finalize_word_docx.py`
  updates fields/TOC.
- **Page 12** — A "figure & table format" sample chapter from the V2
  template's own bundled body, visually verifying that captions are
  centred + bold and three-line tables render correctly.
- **Page 13** — A "math equation format" sample, showing that an inline
  numbered display equation survived the pandoc → adaptive_docx_pipeline →
  Word round-trip.
- **Page 14** — References section: shows that bibliography entries keep
  their numbering and indent style under the template's `参考文献条目` style.

The screenshot is generated against a **real Chinese-thesis template** (the
Zhengzhou University undergraduate thesis template, V2) which itself ships
with a complete sample chapter inside `data/`. Running our pipeline against
it produces a 20-page `final.pdf`; we render those four specific body pages
because they exercise the structural features the skill cares about.

It was generated with:

```bash
python examples/minimal_markdown/run_example.py \
    --template /path/to/your-real-thesis-template.docx \
    --preview-pages "7,12,13,14"
cp examples/minimal_markdown/final.preview.png \
   examples/minimal_markdown/expected/preview.png
```

The committed PNG is also resized to ~70% of the original Word-rendered
resolution to keep the README load lightweight (~700 KB).

Without `--template`, `run_example.py` falls back to the lightweight
`build_template.py`-generated sample, which is intentionally minimal (one
cover style, one body style, one heading) and produces a much simpler
preview. The default `--preview-pages` is `"1-4"`, which on the demo
template covers cover + abstract + a single body chapter. The committed
`expected/preview.png` is **not** the output of that default path; it is
provided to demonstrate behaviour on a realistic institutional template
with realistic body content.

If you're not on Windows, you can still inspect the docx outputs:

- `examples/minimal_markdown/final.docx` — open in Word / LibreOffice
- `examples/minimal_markdown/sample_template.template-report.json` — JSON dump
  of the template's styles, paragraphs, and OOXML hints.
