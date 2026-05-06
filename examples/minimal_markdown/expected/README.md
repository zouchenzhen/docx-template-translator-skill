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
GitHub README. The image is ~280 KB.

Without `--template`, `run_example.py` falls back to the lightweight
`build_template.py`-generated sample, which is intentionally minimal and
fully reproducible. The committed `expected/preview.png` is provided to show
behaviour on a realistic institutional template with realistic Markdown body
content.

If you're not on Windows, you can still inspect the docx outputs:

- `examples/minimal_markdown/final.docx` — open in Word / LibreOffice
- `examples/minimal_markdown/sample_template.template-report.json` — JSON dump
  of the template's styles, paragraphs, and OOXML hints.
