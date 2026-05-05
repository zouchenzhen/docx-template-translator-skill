#!/usr/bin/env python
"""Update Word fields/TOC and optionally export a PDF preview.

Requires Microsoft Word on Windows and pywin32.

Usage:
  python finalize_word_docx.py thesis.docx --pdf

Security note:
  This script disables Word AutoMacros for the input document by setting
  Application.AutomationSecurity = msoAutomationSecurityForceDisable (3).
  Do NOT weaken this unless you fully trust the input .docx, because the
  finalize step opens the document with the real local Microsoft Word.
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pythoncom
import win32com.client as win32

# msoAutomationSecurityForceDisable: 强制禁用所有 AutoMacros / VBA
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
# wdFormatPDF
WD_FORMAT_PDF = 17


def _dispatch_word():
    """Try gencache.EnsureDispatch first; fall back to late-bound Dispatch.

    A common pywin32 failure is a stale / corrupted gen_py cache, which makes
    EnsureDispatch raise AttributeError or ImportError. Falling back to
    Dispatch keeps the script usable without manual cache cleanup.
    """
    try:
        return win32.gencache.EnsureDispatch("Word.Application")
    except (AttributeError, ImportError):
        return win32.Dispatch("Word.Application")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx")
    parser.add_argument("--pdf", action="store_true", help="export PDF beside DOCX")
    parser.add_argument("--pdf-out", default=None)
    args = parser.parse_args()

    docx_path = Path(args.docx).resolve()
    pdf_path = (
        Path(args.pdf_out).resolve() if args.pdf_out else docx_path.with_suffix(".pdf")
    )

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = _dispatch_word()
        word.Visible = False
        word.DisplayAlerts = 0
        # Security: disable macros before opening any user-supplied .docx.
        try:
            word.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception:
            # Older Word versions may not expose AutomationSecurity; we still
            # want the rest of the pipeline to work, just with a softer guard.
            pass

        doc = word.Documents.Open(
            str(docx_path), ReadOnly=False, AddToRecentFiles=False
        )
        doc.Fields.Update()
        for toc in doc.TablesOfContents:
            toc.Update()
        doc.Repaginate()
        doc.Fields.Update()
        doc.Save()
        if args.pdf:
            doc.ExportAsFixedFormat(
                str(pdf_path), ExportFormat=WD_FORMAT_PDF, OpenAfterExport=False
            )
            print(f"pdf: {pdf_path}")
        print(f"updated: {docx_path}")
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
