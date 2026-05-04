#!/usr/bin/env python
"""Starter pipeline for adaptive DOCX template reconstruction.

This script is intentionally not a universal converter. It gives an AI agent a
safe base to copy and modify for a user's concrete template and source project.

Typical use:
  1. Convert LaTeX/Markdown/PDF into a rough body.docx.
  2. Inspect template.docx with inspect_docx_template.py.
  3. Copy this script into the work directory and patch mapping rules.
  4. Run it to produce final.docx, then finalize with Word COM.
"""
from __future__ import annotations

import argparse
import copy
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


def set_east_asia_font(run, font_name: str) -> None:
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font_name)


def suppress_heading_number(paragraph) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    num_pr = ppr.find(qn("w:numPr"))
    if num_pr is None:
        num_pr = OxmlElement("w:numPr")
        ppr.append(num_pr)
    ilvl = num_pr.find(qn("w:ilvl"))
    if ilvl is None:
        ilvl = OxmlElement("w:ilvl")
        num_pr.append(ilvl)
    ilvl.set(qn("w:val"), "0")
    num_id = num_pr.find(qn("w:numId"))
    if num_id is None:
        num_id = OxmlElement("w:numId")
        num_pr.append(num_id)
    num_id.set(qn("w:val"), "0")


def append_docx_body_preserve_relationships(target_doc: Document, source_doc: Document) -> None:
    """Append source body XML to target, remapping image/hyperlink relationships."""
    target_body = target_doc.element.body
    sect_pr = target_body.sectPr
    if sect_pr is not None:
        target_body.remove(sect_pr)

    relmap: dict[str, str] = {}
    rel_attrs = {qn("r:id"), qn("r:embed"), qn("r:link")}

    def remap_relationships(element) -> None:
        for node in element.iter():
            for attr_name in list(node.attrib):
                if attr_name not in rel_attrs:
                    continue
                old_rid = node.attrib[attr_name]
                if old_rid not in source_doc.part.rels:
                    continue
                if old_rid not in relmap:
                    rel = source_doc.part.rels[old_rid]
                    if rel.is_external:
                        relmap[old_rid] = target_doc.part.relate_to(
                            rel.target_ref, rel.reltype, is_external=True
                        )
                    else:
                        relmap[old_rid] = target_doc.part.relate_to(
                            rel.target_part, rel.reltype
                        )
                node.attrib[attr_name] = relmap[old_rid]

    for child in list(source_doc.element.body):
        if child.tag == qn("w:sectPr"):
            continue
        copied = copy.deepcopy(child)
        remap_relationships(copied)
        target_body.append(copied)

    if sect_pr is not None:
        target_body.append(sect_pr)


def set_cell_border(cell, edge: str, *, val: str = "nil", size: int = 0) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    border = tc_borders.find(qn(f"w:{edge}"))
    if border is None:
        border = OxmlElement(f"w:{edge}")
        tc_borders.append(border)
    border.set(qn("w:val"), val)
    if val != "nil":
        border.set(qn("w:sz"), str(size))
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")


def format_three_line_tables(doc: Document) -> None:
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
                    set_cell_border(cell, edge, val="nil")
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    paragraph.paragraph_format.first_line_indent = None
                    for run in paragraph.runs:
                        run.font.size = Pt(10.5)
                        run.font.name = "Times New Roman"
                        set_east_asia_font(run, "宋体")
        if table.rows:
            for cell in table.rows[0].cells:
                set_cell_border(cell, "top", val="single", size=12)
                set_cell_border(cell, "bottom", val="single", size=8)
            for cell in table.rows[-1].cells:
                set_cell_border(cell, "bottom", val="single", size=12)


def normalize_hyperlinks_black(doc: Document) -> None:
    for hyperlink in doc.element.body.iter(qn("w:hyperlink")):
        for run_el in hyperlink.findall(qn("w:r")):
            rpr = run_el.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                run_el.insert(0, rpr)
            color = rpr.find(qn("w:color"))
            if color is None:
                color = OxmlElement("w:color")
                rpr.append(color)
            color.set(qn("w:val"), "000000")
            u = rpr.find(qn("w:u"))
            if u is None:
                u = OxmlElement("w:u")
                rpr.append(u)
            u.set(qn("w:val"), "none")


def apply_basic_style_mapping(
    doc: Document,
    *,
    body_style: str,
    unnumbered_h1: set[str],
    caption_regex: str,
) -> None:
    caption_re = re.compile(caption_regex)
    for paragraph in doc.paragraphs:
        text = re.sub(r"\s+", "", paragraph.text)
        if paragraph.style.name == "Heading 1" and text in unnumbered_h1:
            suppress_heading_number(paragraph)
            continue
        if "<w:drawing" in paragraph._element.xml:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = None
            continue
        if caption_re.match(paragraph.text.strip()):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = None
            paragraph.paragraph_format.space_before = Pt(6)
            paragraph.paragraph_format.space_after = Pt(6)
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(10.5)
            continue
        if paragraph.style.name in {"Normal", "Body Text"} and paragraph.text.strip():
            if body_style in doc.styles:
                paragraph.style = doc.styles[body_style]
            for run in paragraph.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
                set_east_asia_font(run, "宋体")


def load_config(path: Path | None) -> dict:
    if path is None:
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--template", required=True)
    parser.add_argument("--body-docx", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--config", default=None, help="JSON config with style names/rules")
    args = parser.parse_args()

    cfg = load_config(Path(args.config) if args.config else None)
    template_doc = Document(args.template)
    body_doc = Document(args.body_docx)

    append_docx_body_preserve_relationships(template_doc, body_doc)
    apply_basic_style_mapping(
        template_doc,
        body_style=cfg.get("body_style", "Normal"),
        unnumbered_h1=set(cfg.get("unnumbered_h1", [])),
        caption_regex=cfg.get("caption_regex", r"^(图|表)\s*\d+\.\d+\s+"),
    )
    format_three_line_tables(template_doc)
    normalize_hyperlinks_black(template_doc)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    template_doc.save(out)
    print(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
