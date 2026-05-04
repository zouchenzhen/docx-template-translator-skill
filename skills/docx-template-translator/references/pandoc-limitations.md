# Why This Skill Is Not Just `pandoc --reference-doc`

Pandoc is the best first-pass converter for Markdown and LaTeX-like sources, but
`--reference-doc` mainly supplies style definitions. It does not know the
semantic meaning of an institution's template pages.

Use this comparison when explaining the workflow to users:

| Need | Pandoc reference DOCX | This skill's workflow |
| --- | --- | --- |
| Convert Markdown/LaTeX body text | Good first pass | Uses pandoc as rough body converter |
| Fill cover/declaration pages | Not automatic | Copies/fills template front matter via Python |
| Detect real body style | Not guaranteed | Inspects template and remaps styles |
| Preserve images/equations | Often good | Preserves and verifies relationship/OMML XML |
| Three-line tables | Not template-aware | Applies table border XML explicitly |
| Caption numbering | Often incomplete from LaTeX | Can parse `.tex` captions/labels and repair text |
| Citation hyperlinks | Basic citations only | Adds bookmarks/internal hyperlinks if needed |
| Visual verification | Not included | Exports PDF and renders preview sheets |

Recommended framing: pandoc is a parser/rough translator; this skill is an
adaptive reconstruction workflow for strict Word-template deliverables.
