#!/usr/bin/env python
r"""Inject a Word ``{ TOC \o "1-3" \h \z \u }`` field into a DOCX.

Why this exists
===============
Pandoc + ``--reference-doc`` does not always emit a real Word TOC field.
A rough conversion that copies the template's TOC heading paragraph but loses
the surrounding ``<w:fldChar>`` / ``<w:instrText>`` produces a "TOC heading
with nothing under it". Word's "update fields" cannot populate a TOC that
isn't there.

This script idempotently inserts a TOC field paragraph immediately after the
TOC heading. Run it once on the rebuilt DOCX, then call ``finalize_word_docx.py``
(or any field-update step) so Word expands the TOC entries.

Anchor selection
================
The script tries several anchors in order and uses the first that matches:

  1.  ``--anchor-style <styleId>``: insert after the first paragraph styled
      with the given styleId (e.g. ``TOC1`` for the Chinese-thesis 目录 style).
  2.  ``--anchor-text <text>`` (default ``目录``, also accepts ``Contents``,
      ``Table of Contents``): insert after the first paragraph whose text,
      after whitespace normalization, equals the anchor.
  3.  Append the field at the very end of the body (last resort, rare).

Idempotency
===========
If the document already contains an ``<w:instrText>`` that begins with
``TOC``, this script exits without modification (status code 0).

Usage
=====

    python inject_toc_field.py final.docx
    python inject_toc_field.py final.docx --switches '\o "1-3" \h \z \u'
    python inject_toc_field.py final.docx --anchor-style TOC1
    python inject_toc_field.py final.docx --anchor-text "Table of Contents" --in-place
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path


def normalize(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def already_has_toc_field(document_xml: str) -> bool:
    return bool(re.search(r"<w:instrText[^>]*>\s*TOC\s", document_xml))


def build_field_paragraph(switches: str, body_style_id: str) -> str:
    instr = f" TOC {switches.strip()} ".replace("  ", " ")
    return (
        "<w:p>"
        f"<w:pPr><w:pStyle w:val=\"{body_style_id}\"/></w:pPr>"
        "<w:r><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>"
        f"<w:r><w:instrText xml:space=\"preserve\">{instr}</w:instrText></w:r>"
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
        "<w:r><w:t></w:t></w:r>"
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
        "</w:p>"
    )


def find_anchor_position(
    document_xml: str,
    anchor_style: str | None,
    anchor_texts: list[str],
) -> int | None:
    """Return char position immediately after the matched paragraph's </w:p>, or None."""
    para_iter = list(re.finditer(r"<w:p\b[^>]*>.*?</w:p>", document_xml, flags=re.S))
    norm_targets = {normalize(t) for t in anchor_texts}

    if anchor_style:
        for m in para_iter:
            if re.search(
                r"<w:pStyle w:val=\"" + re.escape(anchor_style) + r"\"\s*/>",
                m.group(0),
            ):
                return m.end()

    for m in para_iter:
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", m.group(0)))
        if normalize(text) in norm_targets:
            return m.end()
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("docx")
    parser.add_argument("--switches", default=r'\o "1-3" \h \z \u',
                        help="Field switches, default: %(default)s")
    parser.add_argument("--anchor-style", default=None,
                        help="StyleId to anchor on (e.g. TOC1 / 目录 style)")
    parser.add_argument("--anchor-text", action="append",
                        default=["目录", "目  录", "Contents", "Table of Contents", "Table of contents"],
                        help="Heading text to anchor on (repeatable)")
    parser.add_argument("--body-style-id", default="a",
                        help="Style id for the field paragraph (default 'a' = Normal)")
    parser.add_argument("--out", default=None,
                        help="Output path. If omitted, writes <docx>.with_toc.docx unless --in-place is set.")
    parser.add_argument("--in-place", action="store_true",
                        help="Overwrite the input file in-place")
    args = parser.parse_args()

    src = Path(args.docx)
    if not src.is_file():
        parser.error(f"docx not found: {src}")

    with zipfile.ZipFile(str(src), "r") as zin:
        document_xml = zin.read("word/document.xml").decode("utf-8", errors="replace")

    if already_has_toc_field(document_xml):
        print("TOC field already present; nothing to do.")
        return 0

    insert_at = find_anchor_position(document_xml, args.anchor_style, args.anchor_text)
    if insert_at is None:
        # last resort: insert before the final sectPr at end of body
        body_close = document_xml.rfind("</w:body>")
        last_sect = document_xml.rfind("<w:sectPr", 0, body_close)
        if last_sect == -1:
            print("ERROR: could not find anchor; document has no recognizable 目录/Contents/Table of Contents heading and no sectPr to fall back to.", file=sys.stderr)
            return 2
        # find the start of the paragraph wrapping that sectPr
        para_start = document_xml.rfind("<w:p ", 0, last_sect)
        if para_start == -1:
            para_start = last_sect
        insert_at = para_start
        print("WARN: no TOC heading anchor matched; inserting before the trailing sectPr as a last resort.", file=sys.stderr)

    field_para = build_field_paragraph(args.switches, args.body_style_id)
    new_doc_xml = document_xml[:insert_at] + field_para + document_xml[insert_at:]

    out_path = (
        src
        if args.in_place
        else Path(args.out)
        if args.out
        else src.with_name(src.stem + ".with_toc.docx")
    )

    # Write to a temp file then copy over to keep src readable while writing
    tmp_fd, tmp_name = tempfile.mkstemp(suffix=".docx")
    os.close(tmp_fd)
    Path(tmp_name).unlink(missing_ok=True)
    tmp_path = Path(tmp_name)
    with zipfile.ZipFile(str(src), "r") as zin:
        with zipfile.ZipFile(str(tmp_path), "w", zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)
                if name == "word/document.xml":
                    data = new_doc_xml.encode("utf-8")
                zout.writestr(name, data)
    shutil.move(str(tmp_path), str(out_path))
    print(f"TOC field injected -> {out_path}")
    return 0


if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        if hasattr(_s, "reconfigure"):
            _s.reconfigure(encoding="utf-8", errors="replace")

    raise SystemExit(main())
