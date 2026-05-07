#!/usr/bin/env python
"""Update Word fields/TOC and optionally export a PDF preview.

Requires Microsoft Word on Windows and pywin32.

Usage:
  python finalize_word_docx.py thesis.docx --pdf
  python finalize_word_docx.py thesis.docx --pdf --prefer-dispatch-ex

Security note:
  This script disables Word AutoMacros for the input document by setting
  Application.AutomationSecurity = msoAutomationSecurityForceDisable (3).
  Do NOT weaken this unless you fully trust the input .docx, because the
  finalize step opens the document with the real local Microsoft Word.
"""
from __future__ import annotations

import argparse
import os
from pathlib import Path

import pythoncom
import win32com.client as win32

# msoAutomationSecurityForceDisable: 强制禁用所有 AutoMacros / VBA
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
# wdFormatPDF
WD_FORMAT_PDF = 17


def bootstrap_windows_env() -> None:
    """Fill environment variables often missing in non-interactive shells.

    Word COM can fail with Server execution failed or Operation unavailable when
    SystemRoot/WINDIR or user profile paths are missing from a wrapped CLI
    environment. Keep this local to the current process and do not set proxies.
    """
    if os.name != "nt":
        return
    userprofile = os.environ.get("USERPROFILE") or str(Path.home())
    defaults = {
        "SystemRoot": r"C:\Windows",
        "WINDIR": r"C:\Windows",
        "USERPROFILE": userprofile,
        "HOME": userprofile,
        "APPDATA": str(Path(userprofile) / "AppData" / "Roaming"),
        "LOCALAPPDATA": str(Path(userprofile) / "AppData" / "Local"),
    }
    for key, value in defaults.items():
        os.environ.setdefault(key, value)


def _dispatch_word(*, prefer_dispatch_ex: bool = False):
    """Create a Word.Application COM object with robust fallbacks.

    EnsureDispatch is convenient but can fail with stale gen_py caches. Dispatch
    can attach to a broken existing instance. DispatchEx starts a fresh Word
    process and is often more reliable in Codex/CI-like Windows shells.
    """
    errors = []
    attempts = (
        ("DispatchEx", lambda: win32.DispatchEx("Word.Application")),
        ("EnsureDispatch", lambda: win32.gencache.EnsureDispatch("Word.Application")),
        ("Dispatch", lambda: win32.Dispatch("Word.Application")),
    ) if prefer_dispatch_ex else (
        ("EnsureDispatch", lambda: win32.gencache.EnsureDispatch("Word.Application")),
        ("DispatchEx", lambda: win32.DispatchEx("Word.Application")),
        ("Dispatch", lambda: win32.Dispatch("Word.Application")),
    )
    for name, factory in attempts:
        try:
            return factory()
        except Exception as exc:  # COM failures vary by Office installation.
            errors.append(f"{name}: {exc}")
    raise RuntimeError("Could not start Microsoft Word COM. Attempts: " + " | ".join(errors))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx")
    parser.add_argument("--pdf", action="store_true", help="export PDF beside DOCX")
    parser.add_argument("--pdf-out", default=None)
    parser.add_argument(
        "--prefer-dispatch-ex",
        action="store_true",
        help="start a fresh Word COM instance before trying EnsureDispatch",
    )
    args = parser.parse_args()

    docx_path = Path(args.docx).resolve()
    pdf_path = (
        Path(args.pdf_out).resolve() if args.pdf_out else docx_path.with_suffix(".pdf")
    )

    bootstrap_windows_env()
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = _dispatch_word(prefer_dispatch_ex=args.prefer_dispatch_ex)
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
