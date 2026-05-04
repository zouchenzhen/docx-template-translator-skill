#!/usr/bin/env python
"""Render selected PDF pages into a contact sheet for visual QA.

Usage:
  python render_pdf_preview.py file.pdf --pages 1-8 --out preview.png
"""
from __future__ import annotations

import argparse
from pathlib import Path

import fitz
from PIL import Image, ImageDraw


def parse_pages(spec: str, page_count: int) -> list[int]:
    pages: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start, end = part.split("-", 1)
            pages.extend(range(int(start), int(end) + 1))
        else:
            pages.append(int(part))
    return [p - 1 for p in pages if 1 <= p <= page_count]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf")
    parser.add_argument("--pages", default="1-8")
    parser.add_argument("--scale", type=float, default=0.35)
    parser.add_argument("--columns", type=int, default=4)
    parser.add_argument("--out", default=None)
    args = parser.parse_args()

    pdf = Path(args.pdf)
    out = Path(args.out) if args.out else pdf.with_suffix(".preview.png")
    doc = fitz.open(pdf)
    page_indices = parse_pages(args.pages, len(doc))

    thumbs = []
    for idx in page_indices:
        pix = doc[idx].get_pixmap(matrix=fitz.Matrix(args.scale, args.scale), alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        canvas = Image.new("RGB", (img.width, img.height + 26), "white")
        canvas.paste(img, (0, 26))
        ImageDraw.Draw(canvas).text((6, 6), f"p{idx + 1}", fill="black")
        thumbs.append(canvas)

    if not thumbs:
        raise SystemExit("no pages selected")

    cell_w = max(t.width for t in thumbs)
    cell_h = max(t.height for t in thumbs)
    rows = (len(thumbs) + args.columns - 1) // args.columns
    sheet = Image.new("RGB", (cell_w * args.columns, cell_h * rows), "white")
    for i, img in enumerate(thumbs):
        sheet.paste(img, ((i % args.columns) * cell_w, (i // args.columns) * cell_h))
    sheet.save(out)
    print(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
