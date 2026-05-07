# expected/

This directory holds the **reference visual output** for the
`minimal_markdown` example, so a reader can see the produced look-and-feel
without setting up Word + pywin32 locally.

## preview.png

`preview.png` is a cropped body-content screenshot from `final.pdf`, produced
on Windows with Microsoft Word and a real Zhengzhou University undergraduate
thesis template (V2).

It is deliberately **not** a cover-page screenshot and not a screenshot of
the template's bundled sample chapter. The visible content comes from this
repository's own `sample.md` after:

1. `pandoc` converts Markdown to rough DOCX,
2. `adaptive_docx_pipeline.py` maps the body into the thesis template styles,
3. `finalize_word_docx.py` asks Word to update/export the final PDF.

The committed crop focuses on PDF page 20 because that single generated page
shows the most important body features at once:

- an embedded PNG figure from `assets/pipeline_diagram.png`,
- a generated figure caption,
- Chinese/English body paragraphs,
- a display equation,
- a Markdown table converted into a three-line table,
- a second caption-like block for caption-rule regression coverage.

Generate the source PDF with:

```bash
python examples/minimal_markdown/run_example.py \
    --template /path/to/your-real-thesis-template.docx \
    --preview-pages "20"
```

The committed PNG is then cropped from page 20 to remove the unrelated
running header and page margins, keeping the body region readable in the
GitHub README. The image is ~313 KB.

## pandoc_baseline.png

`pandoc_baseline.png` is the direct baseline used in the README comparison.
It uses the same `sample.md` and the same real thesis template, but skips this
project's inspect/adaptive pipeline:

```bash
pandoc examples/minimal_markdown/sample.md \
    --resource-path examples/minimal_markdown \
    --reference-doc /path/to/your-real-thesis-template.docx \
    -o pandoc_baseline.docx
python skills/docx-template-translator/scripts/finalize_word_docx.py \
    pandoc_baseline.docx --pdf
```

The committed PNG stacks the relevant rendered PDF pages because the direct
pandoc output splits the figure/equation/table/caption region across pages.
This is intentional: it shows the limitation of `pandoc --reference-doc`
alone. It can reuse style definitions, but it does not understand the target
template's semantic layout rules, nor does it apply this project's adaptive
three-line-table and caption/body remapping logic.

The numbering differs for the same reason. `preview.png` preserves the real
template's existing document context, so headings from `sample.md` continue
after the chapters already present in the template. `pandoc_baseline.png`
starts a new document from `sample.md`, using the template only as
`--reference-doc`, so the same headings start under Chapter 1.

Both README images have a light border baked into the PNG itself. This is
done by drawing the Word/PDF crop onto a slightly larger light-gray canvas and
adding a thin gray rectangle, rather than by relying on GitHub Markdown/CSS.

Without `--template`, `run_example.py` falls back to the lightweight
`build_template.py`-generated sample, which is intentionally minimal and
fully reproducible. The committed `expected/preview.png` is provided to show
behaviour on a realistic institutional template with realistic Markdown body
content.

If you're not on Windows, you can still inspect the docx outputs:

- `examples/minimal_markdown/final.docx` — open in Word / LibreOffice
- `examples/minimal_markdown/sample_template.template-report.json` — JSON dump
  of the template's styles, paragraphs, and OOXML hints.
