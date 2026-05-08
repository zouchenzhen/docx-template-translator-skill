#!/usr/bin/env python
r"""Replace a section's running-header text with a STYLEREF field.

Why this exists
===============
When a thesis pipeline collapses the body and back-matter into a single
section, the body running header inherits whatever static text the last
template header part contained — frequently "致谢" / "Acknowledgements".
The cleanest portable fix is to replace the static text with a Word field
that reads the current chapter title:

    { STYLEREF <styleId> \* MERGEFORMAT }

This script does that on a single ``word/headerN.xml`` part of an existing
DOCX, preserving the run-level rPr (font/size/bold) of the first run so the
rendered header keeps the template's typography.

Usage
=====

    python set_styleref_header.py final.docx --header header5.xml
    python set_styleref_header.py final.docx --header header5.xml --style-id 1
    python set_styleref_header.py final.docx --header header5.xml --style-name "Heading 1"
    python set_styleref_header.py final.docx --header header5.xml --in-place

Tips
====
* On localized templates the heading **display name** ("Heading 1" / "标题 1")
  may not resolve the field — Word will print the localized error string
  ("错误!使用'开始'选项卡将 Heading 1 应用于要在此处显示的文字"). Use the
  numeric ``--style-id`` form (e.g. ``1`` for the first heading style) for
  maximum portability.
* Combine with ``inject_toc_field.py`` and ``finalize_word_docx.py``:
  inject TOC, set STYLEREF header, then update fields and export PDF.
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


def build_field_runs(rpr: str, style_ref_target: str) -> str:
    return (
        f"<w:r>{rpr}<w:fldChar w:fldCharType=\"begin\"/></w:r>"
        f"<w:r>{rpr}<w:instrText xml:space=\"preserve\"> STYLEREF {style_ref_target} \\* MERGEFORMAT </w:instrText></w:r>"
        f"<w:r>{rpr}<w:fldChar w:fldCharType=\"separate\"/></w:r>"
        f"<w:r>{rpr}<w:t></w:t></w:r>"
        f"<w:r>{rpr}<w:fldChar w:fldCharType=\"end\"/></w:r>"
    )


def rewrite_first_paragraph(header_xml: str, style_ref_target: str) -> tuple[str, str]:
    """Return (new_xml, status). status is 'ok', 'no-paragraph', or 'no-runs'."""
    m = re.search(r"(<w:p\b[^>]*>)(.*?)(</w:p>)", header_xml, flags=re.S)
    if not m:
        return header_xml, "no-paragraph"
    para_open = m.group(1)
    body = m.group(2)
    para_close = m.group(3)

    ppr_m = re.search(r"<w:pPr\b[^>]*>.*?</w:pPr>", body, flags=re.S)
    pPr = ppr_m.group(0) if ppr_m else ""

    first_run_m = re.search(r"<w:r\b[^>]*>.*?</w:r>", body, flags=re.S)
    rPr = ""
    if first_run_m:
        rpr_m = re.search(r"<w:rPr\b[^>]*>.*?</w:rPr>", first_run_m.group(0), flags=re.S)
        rPr = rpr_m.group(0) if rpr_m else ""

    new_body = pPr + build_field_runs(rPr, style_ref_target)
    return header_xml[: m.start()] + para_open + new_body + para_close + header_xml[m.end():], "ok"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("docx")
    parser.add_argument("--header", required=True,
                        help="Filename inside word/, e.g. header5.xml")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--style-id", default="1",
                       help="Numeric/string styleId, default 1 (first Heading 1 style)")
    group.add_argument("--style-name", default=None,
                       help="Style display name, e.g. 'Heading 1'. Less portable than --style-id.")
    parser.add_argument("--out", default=None)
    parser.add_argument("--in-place", action="store_true")
    args = parser.parse_args()

    src = Path(args.docx)
    if not src.is_file():
        parser.error(f"docx not found: {src}")

    target = args.style_name if args.style_name else args.style_id
    if args.style_name and " " in args.style_name:
        target = f"\"{args.style_name}\""

    member = f"word/{args.header.lstrip('/')}"
    with zipfile.ZipFile(str(src), "r") as zin:
        if member not in zin.namelist():
            print(f"ERROR: {member} not found in docx", file=sys.stderr)
            return 2
        header_xml = zin.read(member).decode("utf-8", errors="replace")

    new_xml, status = rewrite_first_paragraph(header_xml, target)
    if status != "ok":
        print(f"ERROR: cannot rewrite {member}: {status}", file=sys.stderr)
        return 2
    if new_xml == header_xml:
        print(f"WARN: rewrite produced no change in {member}", file=sys.stderr)

    out_path = (
        src
        if args.in_place
        else Path(args.out)
        if args.out
        else src.with_name(src.stem + ".styleref_header.docx")
    )

    tmp_fd, tmp_name = tempfile.mkstemp(suffix=".docx")
    os.close(tmp_fd)
    Path(tmp_name).unlink(missing_ok=True)
    tmp_path = Path(tmp_name)
    with zipfile.ZipFile(str(src), "r") as zin:
        with zipfile.ZipFile(str(tmp_path), "w", zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)
                if name == member:
                    data = new_xml.encode("utf-8")
                zout.writestr(name, data)
    shutil.move(str(tmp_path), str(out_path))
    print(f"STYLEREF header injected ({member} -> ref={target}) -> {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
