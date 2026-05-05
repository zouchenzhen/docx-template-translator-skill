#!/usr/bin/env python
"""Render selected PDF pages into a contact sheet for visual QA.

Usage:
  python render_pdf_preview.py file.pdf --pages 1-8 --out preview.png
  python render_pdf_preview.py file.pdf --dpi 150               # higher quality
  python render_pdf_preview.py file.pdf --columns 2 --dpi 150

By default the script renders at 150 DPI and lays the thumbnails out in 2
columns, which keeps Chinese text legible on GitHub previews. Pass --scale to
fall back to the legacy multiplier-based behaviour.
"""
from __future__ import annotations

import argparse
from pathlib import Path

import fitz
from PIL import Image, ImageDraw, ImageFont


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


def label_font(size: int) -> ImageFont.ImageFont:
    """Pick a font that can render Chinese page labels if requested."""
    candidates = [
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simsun.ttc",
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/PingFang.ttc",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf")
    parser.add_argument("--pages", default="1-8")
    # Default DPI is 150 — A4 → 1240x1754, large enough that CJK text stays
    # readable when the contact sheet is viewed inline on GitHub.
    parser.add_argument("--dpi", type=int, default=150,
                        help="render DPI (default 150). Overrides --scale if both given.")
    parser.add_argument("--scale", type=float, default=None,
                        help="legacy scale multiplier; only used when --dpi is not provided")
    parser.add_argument("--columns", type=int, default=2)
    parser.add_argument("--label-size", type=int, default=22)
    parser.add_argument("--out", default=None)
    args = parser.parse_args()

    pdf = Path(args.pdf)
    out = Path(args.out) if args.out else pdf.with_suffix(".preview.png")

    if args.scale is not None and args.dpi == 150:
        # User explicitly opted into the legacy --scale mode (no --dpi given).
        zoom = args.scale
    else:
        zoom = args.dpi / 72.0

    label_h = max(28, args.label_size + 8)
    fnt = label_font(args.label_size)

    doc = fitz.open(pdf)
    try:
        page_indices = parse_pages(args.pages, len(doc))

        thumbs = []
        for idx in page_indices:
            pix = doc[idx].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            canvas = Image.new("RGB", (img.width, img.height + label_h), "white")
            canvas.paste(img, (0, label_h))
            ImageDraw.Draw(canvas).text((10, 4), f"p{idx + 1}", fill="black", font=fnt)
            thumbs.append(canvas)
    finally:
        doc.close()

    if not thumbs:
        raise SystemExit("no pages selected")

    cell_w = max(t.width for t in thumbs)
    cell_h = max(t.height for t in thumbs)
    rows = (len(thumbs) + args.columns - 1) // args.columns
    sheet = Image.new("RGB", (cell_w * args.columns, cell_h * rows), "white")
    for i, img in enumerate(thumbs):
        sheet.paste(img, ((i % args.columns) * cell_w, (i // args.columns) * cell_h))
    sheet.save(out, optimize=True)
    print(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
