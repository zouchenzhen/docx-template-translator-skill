#!/usr/bin/env python
"""Update Word fields/TOC and optionally export a PDF preview.

Requires Microsoft Word on Windows and pywin32.

Usage:
  python finalize_word_docx.py thesis.docx --pdf
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pythoncom
import win32com.client as win32


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx")
    parser.add_argument("--pdf", action="store_true", help="export PDF beside DOCX")
    parser.add_argument("--pdf-out", default=None)
    args = parser.parse_args()

    docx_path = Path(args.docx).resolve()
    pdf_path = Path(args.pdf_out).resolve() if args.pdf_out else docx_path.with_suffix(".pdf")

    pythoncom.CoInitialize()
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(docx_path), ReadOnly=False, AddToRecentFiles=False)
        doc.Fields.Update()
        for toc in doc.TablesOfContents:
            toc.Update()
        doc.Repaginate()
        doc.Fields.Update()
        doc.Save()
        if args.pdf:
            doc.ExportAsFixedFormat(str(pdf_path), ExportFormat=17, OpenAfterExport=False)
            print(f"pdf: {pdf_path}")
        doc.Close(False)
        print(f"updated: {docx_path}")
    finally:
        word.Quit()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
