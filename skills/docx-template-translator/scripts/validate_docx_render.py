#!/usr/bin/env python
r"""Render-level QA for a reconstructed DOCX/PDF.

Why a second validator
======================
``validate_docx_conversion.py`` checks structure: counts, placeholders, ordered
terms, heading-style preservation, image/table totals, protected-front-matter
formatting drift. Those checks can return ``STATUS: PASS`` while the document
is visibly broken in Word/PDF — for example, an empty TOC paragraph, an
unbound STYLEREF field, a chapter heading style without ``<w:numPr>``, or a
shared numbering counter between body headings and reference list entries.

This validator targets the *rendering* layer: it inspects the OOXML packages
that Word actually consumes (``word/numbering.xml``, ``word/styles.xml``,
``word/document.xml``, ``word/headerN.xml``) and the resulting PDF text, and
fails when any of the following are detected:

1.  ``--require-toc-field`` (default on): the document must contain at least
    one ``<w:fldChar>``-bracketed ``TOC`` field. Word's "update fields" cannot
    populate a non-existent TOC.
2.  ``--require-numbering-consistency`` (default on): every ``(numId, ilvl)``
    pair used by a paragraph or style must resolve to a defined ``<w:lvl>``
    inside the bound abstract numbering. A missing level falls back to level 0
    and silently degrades multilevel headings to bracketed counters.
3.  ``--require-multilevel-headings`` (default on, when at least one Heading 1
    style numId is present): the bound abstract numbering's level 0/1/2
    ``lvlText`` must look like a chapter / section / subsection format string.
    The default expectation is ``第%1章``-style for level 0 and ``%1.%2`` /
    ``%1.%2.%3`` for level 1/2; configure with ``--chapter-prefix-pattern``
    and ``--multilevel-pattern`` for templates that use ``Chapter %1``,
    ``%1-%2``, etc.
4.  ``--require-ref-counter-independence`` (default on): the numId used by any
    heading style must not also be used by any non-heading paragraph that
    appears after the last "References / 参考文献" Heading 1. This is the bug
    where 33 references render as ``[47]…[79]`` because their counter was
    shared with H2/H3 paragraphs.
5.  ``--pdf`` (when given): the PDF text must not contain Word's localized
    field-error strings ("错误!", "Error!", "!Reference source not found").
    These typically come from broken STYLEREF / PAGEREF / REF fields.
6.  ``--require-body-header-non-back-matter`` (default on): for every body
    section that uses a ``<w:headerReference>``, the referenced ``headerN.xml``
    must either contain a Word field (``<w:fldChar>``) or its static text
    must not match a back-matter title ("致谢", "致 谢",
    "Acknowledgements", "参考文献", "References", "附录", "Appendix",
    "攻读学位期间…"). The body running header showing "致谢" was the
    first signal that the body+back-matter were collapsed into one section.
7.  ``--source-latex-dir <dir>`` or ``--min-figures <N>``: count
    ``<w:drawing>`` in ``document.xml`` and FAIL if it is below the source
    LaTeX project's ``\\includegraphics`` count. Pandoc silently drops
    figures whose path lacks an explicit extension (e.g.
    ``\\includegraphics{thesis_structure}``) when the basename resolves to a
    vector file (``.pdf`` / ``.vsdx``) it cannot embed. Without this check
    the structural validator counts what *was* embedded and never knows what
    *should have been* embedded.
8.  ``--expected-table-style three-line`` (default ``any``): every data
    table (>= ``--table-min-data-rows`` rows) must have a recognizable
    three-line layout (top heavy, header-row bottom thin, last-row bottom
    heavy, vertical edges nil) or a ``tblStyle`` that may carry borders.
    A docx that contains 20 borderless tables when the school template
    requires three-line tables is the failure mode this catches.

Each check is reported individually so a CI run can see exactly which
rendering invariant broke. Pass ``--allow X`` (e.g. ``--allow toc-field``) to
demote a check from FAIL to WARN when a project legitimately deviates.

Usage
=====

    python validate_docx_render.py final.docx \
        --pdf final.pdf \
        --out validation_render.json

    # turn off TOC field requirement (project intentionally has no TOC):
    python validate_docx_render.py final.docx --allow toc-field

    # supply project-specific chapter-prefix pattern:
    python validate_docx_render.py final.docx \
        --chapter-prefix-pattern '^(第%1[章节]|Chapter\s+%1)'
"""
from __future__ import annotations

import sys
import argparse
import json
import re
import zipfile
from pathlib import Path

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

DEFAULT_BACK_MATTER_TITLES = [
    "致谢",
    "致 谢",
    "致  谢",
    "Acknowledgements",
    "Acknowledgement",
    "参考文献",
    "参 考 文 献",
    "References",
    "Bibliography",
    "附录",
    "附 录",
    "Appendix",
    "Appendices",
    "攻读学位期间学术论文和科研成果目录",
    "攻读学位期间",
]

CHECK_NAMES = (
    "toc-field",
    "numbering-consistency",
    "multilevel-headings",
    "ref-counter-independence",
    "pdf-field-errors",
    "body-header-non-back-matter",
    "figure-count-vs-source",
    "table-border-style",
    "citation-coverage",
    "caption-count-vs-source",
    "caption-numbering",
    "caption-centering",
    "caption-untagged-near-figure-table",
)

DEFAULT_CHAPTER_PREFIX_PATTERN = r"(第\s*%1\s*章|Chapter\s+%1)"
DEFAULT_MULTILEVEL_PATTERN = r"%1.*%2"


def normalize(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def read_part(zf: zipfile.ZipFile, name: str) -> str:
    try:
        return zf.read(name).decode("utf-8", errors="replace")
    except KeyError:
        return ""


def find_style_block(styles_xml: str, style_id: str) -> str | None:
    pattern = re.compile(
        r"<w:style\b[^>]*\bw:styleId=\"" + re.escape(style_id) + r"\"[^>]*>(.*?)</w:style>",
        re.S,
    )
    m = pattern.search(styles_xml)
    return m.group(0) if m else None


def heading_style_ids_by_name(styles_xml: str) -> dict[str, str]:
    """Map a heading display name (e.g. "heading 1") to its styleId."""
    out: dict[str, str] = {}
    for block in re.findall(r"<w:style\b[^>]*?>.*?</w:style>", styles_xml, flags=re.S):
        sid_m = re.search(r"w:styleId=\"([^\"]+)\"", block)
        nm_m = re.search(r"<w:name w:val=\"([^\"]+)\"", block)
        if not (sid_m and nm_m):
            continue
        sid = sid_m.group(1)
        nm = nm_m.group(1).strip()
        nm_lower = nm.lower()
        if nm_lower in {"heading 1", "heading 2", "heading 3", "标题 1", "标题 2", "标题 3"}:
            out[nm_lower] = sid
    return out


def style_num_pr(styles_xml: str, style_id: str) -> tuple[str | None, str | None]:
    """Return (numId, ilvl) declared at style level for the given style."""
    block = find_style_block(styles_xml, style_id)
    if not block:
        return None, None
    num_pr_m = re.search(r"<w:numPr\b[^>]*>(.*?)</w:numPr>", block, flags=re.S)
    if not num_pr_m:
        return None, None
    inner = num_pr_m.group(1)
    nid_m = re.search(r"<w:numId w:val=\"(\d+)\"", inner)
    ilvl_m = re.search(r"<w:ilvl w:val=\"(\d+)\"", inner)
    return (nid_m.group(1) if nid_m else None, ilvl_m.group(1) if ilvl_m else None)


def parse_numbering(numbering_xml: str) -> tuple[dict[str, str], dict[str, dict[str, dict]]]:
    """Return (numId -> abstractNumId, abstractNumId -> {ilvl -> {fmt, lvlText, start}})."""
    num_to_abstract: dict[str, str] = {}
    abstract_levels: dict[str, dict[str, dict]] = {}
    if not numbering_xml:
        return num_to_abstract, abstract_levels

    for blk in re.findall(r"<w:num\b[^>]*>.*?</w:num>", numbering_xml, flags=re.S):
        nid_m = re.search(r"w:numId=\"(\d+)\"", blk)
        anid_m = re.search(r"<w:abstractNumId w:val=\"(\d+)\"", blk)
        if nid_m and anid_m:
            num_to_abstract[nid_m.group(1)] = anid_m.group(1)

    for blk in re.findall(r"<w:abstractNum\b[^>]*>.*?</w:abstractNum>", numbering_xml, flags=re.S):
        anid_m = re.search(r"w:abstractNumId=\"(\d+)\"", blk)
        if not anid_m:
            continue
        anid = anid_m.group(1)
        levels: dict[str, dict] = {}
        for lvl in re.findall(r"<w:lvl\b[^>]*>.*?</w:lvl>", blk, flags=re.S):
            ilvl_m = re.search(r"w:ilvl=\"(\d+)\"", lvl)
            if not ilvl_m:
                continue
            fmt_m = re.search(r"<w:numFmt w:val=\"([^\"]+)\"", lvl)
            txt_m = re.search(r"<w:lvlText w:val=\"([^\"]*)\"", lvl)
            start_m = re.search(r"<w:start w:val=\"(\d+)\"", lvl)
            levels[ilvl_m.group(1)] = {
                "fmt": fmt_m.group(1) if fmt_m else None,
                "lvlText": txt_m.group(1) if txt_m else None,
                "start": start_m.group(1) if start_m else None,
            }
        abstract_levels[anid] = levels
    return num_to_abstract, abstract_levels


def collect_paragraph_numpr_usages(document_xml: str) -> list[dict]:
    """For every paragraph in document.xml, collect its style id, numId, ilvl, text head."""
    usages: list[dict] = []
    for p in re.findall(r"<w:p\b[^>]*>.*?</w:p>", document_xml, flags=re.S):
        style_m = re.search(r"<w:pStyle w:val=\"([^\"]+)\"", p)
        nid_m = re.search(r"<w:numId w:val=\"(\d+)\"", p)
        ilvl_m = re.search(r"<w:ilvl w:val=\"(\d+)\"", p)
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p))
        usages.append({
            "style_id": style_m.group(1) if style_m else None,
            "numId": nid_m.group(1) if nid_m else None,
            "ilvl": ilvl_m.group(1) if ilvl_m else None,
            "text_head": text[:60],
        })
    return usages


def section_header_refs(document_xml: str) -> list[list[dict]]:
    """For each sectPr, return a list of {type, rId} headerReference dicts."""
    out: list[list[dict]] = []
    for sect in re.findall(r"<w:sectPr\b[^>]*>.*?</w:sectPr>", document_xml, flags=re.S):
        refs = []
        for m in re.finditer(
            r"<w:headerReference\s+[^/]*?w:type=\"([^\"]+)\"\s+r:id=\"([^\"]+)\"",
            sect,
        ):
            refs.append({"type": m.group(1), "rId": m.group(2)})
        out.append(refs)
    return out


def parse_doc_rels(rels_xml: str) -> dict[str, str]:
    out: dict[str, str] = {}
    for m in re.finditer(
        r"<Relationship\s+Id=\"([^\"]+)\"\s+Type=\"([^\"]+)\"\s+Target=\"([^\"]+)\"",
        rels_xml,
    ):
        out[m.group(1)] = m.group(3)
    return out


# ---- individual checks ----

def check_toc_field(document_xml: str) -> dict:
    has_toc_instr = bool(re.search(r"<w:instrText[^>]*>\s*TOC\s", document_xml))
    has_fld_begin = "fldCharType=\"begin\"" in document_xml
    return {
        "name": "toc-field",
        "passed": has_toc_instr and has_fld_begin,
        "evidence": {"has_toc_instrText": has_toc_instr, "has_fldChar_begin": has_fld_begin},
    }


def check_numbering_consistency(
    document_xml: str,
    styles_xml: str,
    num_to_abstract: dict[str, str],
    abstract_levels: dict[str, dict[str, dict]],
) -> dict:
    failures: list[dict] = []
    seen: set[tuple[str, str]] = set()

    # paragraph-level usages
    for u in collect_paragraph_numpr_usages(document_xml):
        nid = u["numId"]
        if nid is None or nid == "0":
            continue
        ilvl = u["ilvl"] or "0"
        seen.add((nid, ilvl))

    # style-level usages: for any style that defines numPr, capture (numId, ilvl).
    for block in re.findall(r"<w:style\b[^>]*?>.*?</w:style>", styles_xml, flags=re.S):
        sid_m = re.search(r"w:styleId=\"([^\"]+)\"", block)
        if not sid_m:
            continue
        nid, ilvl = style_num_pr(styles_xml, sid_m.group(1))
        if nid and nid != "0":
            seen.add((nid, ilvl or "0"))

    for nid, ilvl in sorted(seen):
        anid = num_to_abstract.get(nid)
        if anid is None:
            failures.append({"numId": nid, "ilvl": ilvl, "reason": "numId not bound to any abstractNumId"})
            continue
        levels = abstract_levels.get(anid, {})
        if ilvl not in levels:
            failures.append({
                "numId": nid,
                "ilvl": ilvl,
                "abstractNumId": anid,
                "reason": "abstractNum has no <w:lvl> for this ilvl; numbering will fall back to level 0",
            })

    return {"name": "numbering-consistency", "passed": not failures, "failures": failures, "checked_pairs": sorted(seen)}


def check_multilevel_headings(
    document_xml: str,
    styles_xml: str,
    num_to_abstract: dict[str, str],
    abstract_levels: dict[str, dict[str, dict]],
    chapter_prefix_pattern: str,
    multilevel_pattern: str,
    back_matter_titles: list[str],
) -> dict:
    """Inspect the abstractNum that Heading 1 actually consumes.

    Acceptable forms:
      - Style-level: ``Heading 1`` style declares ``<w:numPr><w:numId .../></w:numPr>``.
      - Paragraph-level: each body Heading 1 paragraph (i.e. NOT a back-matter
        title like 致谢/参考文献/附录) declares inline numPr.

    The check fails only when *neither* path leads to a multilevel abstract
    numbering whose level 0/1/2 lvlText looks like a chapter format.
    """
    name_to_styleid = heading_style_ids_by_name(styles_xml)
    h1_sid = name_to_styleid.get("heading 1") or name_to_styleid.get("标题 1")
    if not h1_sid:
        return {"name": "multilevel-headings", "passed": True, "evidence": "no Heading 1 style — skipped"}

    # 1. style-level binding
    style_nid, _ = style_num_pr(styles_xml, h1_sid)
    candidate_nids: set[str] = set()
    if style_nid and style_nid != "0":
        candidate_nids.add(style_nid)

    # 2. paragraph-level body H1 bindings (skip front/back matter titles)
    forbidden_norms = {normalize(t) for t in back_matter_titles}
    body_h1_nids: list[str] = []
    body_h1_count = 0
    for u in collect_paragraph_numpr_usages(document_xml):
        if u["style_id"] != h1_sid:
            continue
        if normalize(u["text_head"]) in forbidden_norms:
            continue
        body_h1_count += 1
        if u["numId"] and u["numId"] != "0":
            body_h1_nids.append(u["numId"])
            candidate_nids.add(u["numId"])

    if not candidate_nids:
        return {
            "name": "multilevel-headings",
            "passed": False,
            "failures": [{
                "reason": "Heading 1 has no numPr at style level OR on any body paragraph; chapters will not auto-number",
                "styleId": h1_sid,
                "body_h1_paragraph_count": body_h1_count,
            }],
        }

    # If we found candidates, verify at least one resolves to a multilevel format we like.
    detail: list[dict] = []
    any_ok = False
    for nid in sorted(candidate_nids):
        anid = num_to_abstract.get(nid)
        if anid is None:
            detail.append({"numId": nid, "reason": "numId not bound to any abstractNumId"})
            continue
        levels = abstract_levels.get(anid, {})
        lvl0 = (levels.get("0") or {}).get("lvlText") or ""
        lvl1 = (levels.get("1") or {}).get("lvlText") or ""
        lvl2 = (levels.get("2") or {}).get("lvlText") or ""
        problems: list[dict] = []
        if not re.search(chapter_prefix_pattern, lvl0):
            problems.append({"ilvl": "0", "lvlText": lvl0, "expected_pattern": chapter_prefix_pattern})
        if lvl1 and not re.search(multilevel_pattern, lvl1):
            problems.append({"ilvl": "1", "lvlText": lvl1, "expected_pattern": multilevel_pattern})
        if lvl2 and not re.search(multilevel_pattern, lvl2):
            problems.append({"ilvl": "2", "lvlText": lvl2, "expected_pattern": multilevel_pattern})
        detail.append({
            "numId": nid,
            "abstractNumId": anid,
            "levels": {"0": lvl0, "1": lvl1, "2": lvl2},
            "problems": problems,
        })
        if not problems:
            any_ok = True

    return {
        "name": "multilevel-headings",
        "passed": any_ok,
        "style_level_numId": style_nid,
        "paragraph_level_numIds_used": sorted(set(body_h1_nids)),
        "body_h1_paragraph_count": body_h1_count,
        "detail": detail,
    }


def check_ref_counter_independence(
    document_xml: str,
    styles_xml: str,
    reference_marker_terms: list[str],
) -> dict:
    """
    Heuristic:
      1. Collect every numId used by any paragraph that is styled as Heading 1/2/3
         (or whose style id matches a Heading 1/2/3 styleId).
      2. Find the *last* paragraph whose text matches a reference-section marker
         (e.g. '参考文献' / 'References') and is styled Heading 1.
      3. After that point, look at the next non-Heading paragraphs that carry
         numPr. If any of them uses a numId that is also a heading numId → FAIL.
    """
    name_to_styleid = heading_style_ids_by_name(styles_xml)
    heading_style_ids = set(name_to_styleid.values())
    heading_num_ids: set[str] = set()
    # collect heading numIds from style level
    for sid in heading_style_ids:
        nid, _ = style_num_pr(styles_xml, sid)
        if nid and nid != "0":
            heading_num_ids.add(nid)
    # also collect from any paragraph that has heading style + paragraph-level numPr
    paragraphs = []
    for p in re.findall(r"<w:p\b[^>]*>.*?</w:p>", document_xml, flags=re.S):
        style_m = re.search(r"<w:pStyle w:val=\"([^\"]+)\"", p)
        nid_m = re.search(r"<w:numId w:val=\"(\d+)\"", p)
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p))
        sid = style_m.group(1) if style_m else None
        nid = nid_m.group(1) if nid_m else None
        is_heading = sid in heading_style_ids
        paragraphs.append({"sid": sid, "nid": nid, "text": text, "is_heading": is_heading})
        if is_heading and nid and nid != "0":
            heading_num_ids.add(nid)

    # find ref section start
    ref_norms = {normalize(t) for t in reference_marker_terms}
    ref_start_idx = -1
    for idx, info in enumerate(paragraphs):
        if info["is_heading"] and normalize(info["text"]) in ref_norms:
            ref_start_idx = idx
            break
    if ref_start_idx < 0:
        return {
            "name": "ref-counter-independence",
            "passed": True,
            "evidence": "no reference-section heading found — skipped",
        }

    failures: list[dict] = []
    for info in paragraphs[ref_start_idx + 1:]:
        if info["is_heading"]:
            # leaving the ref section
            break
        if info["nid"] and info["nid"] != "0" and info["nid"] in heading_num_ids:
            failures.append({
                "reason": "reference paragraph reuses a heading numId — counter spillover risk",
                "numId": info["nid"],
                "text_head": info["text"][:60],
            })
            if len(failures) >= 8:
                break
    return {
        "name": "ref-counter-independence",
        "passed": not failures,
        "heading_num_ids": sorted(heading_num_ids),
        "failures": failures,
    }


def check_pdf_field_errors(pdf_path: Path) -> dict:
    try:
        import fitz
    except ImportError:
        return {"name": "pdf-field-errors", "passed": True, "evidence": "PyMuPDF not installed — skipped"}
    err_patterns = [
        ("zh", "错误!"),
        ("en", "Error!"),
        ("ref", "!Reference source not found"),
        ("ref-zh", "!未找到引用源"),
    ]
    failures: list[dict] = []
    with fitz.open(str(pdf_path)) as pdf:
        for pn, page in enumerate(pdf, start=1):
            txt = page.get_text("text")
            for tag, needle in err_patterns:
                if needle in txt:
                    failures.append({"page": pn, "pattern": tag, "needle": needle})
    return {"name": "pdf-field-errors", "passed": not failures, "failures": failures[:20]}


def check_body_header_non_back_matter(
    zf: zipfile.ZipFile,
    document_xml: str,
    rels_map: dict[str, str],
    back_matter_titles: list[str],
) -> dict:
    failures: list[dict] = []
    forbidden_norms = {normalize(t) for t in back_matter_titles}
    for sect_idx, refs in enumerate(section_header_refs(document_xml)):
        for ref in refs:
            target = rels_map.get(ref["rId"])
            if not target:
                continue
            header_path = "word/" + target.lstrip("/").replace("..", "").lstrip("/")
            if not header_path.endswith(".xml"):
                continue
            txt = read_part(zf, header_path)
            if not txt:
                continue
            visible = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", txt))
            has_field = "<w:fldChar" in txt
            if has_field:
                continue
            if normalize(visible) in forbidden_norms:
                failures.append({
                    "section_index": sect_idx,
                    "header_target": target,
                    "type": ref["type"],
                    "static_text": visible[:80],
                    "reason": "section header carries a back-matter title literal — likely inherited",
                })
    return {"name": "body-header-non-back-matter", "passed": not failures, "failures": failures}


def check_figure_count_vs_source(
    document_xml: str,
    source_latex_dir: Path | None,
    explicit_min_figures: int | None,
) -> dict:
    """Compare the count of <w:drawing> elements in document.xml against the
    number of \\includegraphics references in the LaTeX source directory."""
    drawings = document_xml.count("<w:drawing")
    expected = explicit_min_figures
    source_count = None
    if source_latex_dir is not None and source_latex_dir.is_dir():
        source_count = 0
        for tex in source_latex_dir.rglob("*.tex"):
            try:
                source_count += sum(
                    1
                    for line in tex.read_text(encoding="utf-8", errors="replace").splitlines()
                    if "\\includegraphics" in line and not line.lstrip().startswith("%")
                )
            except OSError:
                pass
        if expected is None:
            expected = source_count
    if expected is None:
        return {
            "name": "figure-count-vs-source",
            "passed": True,
            "evidence": "no --source-latex-dir or --min-figures supplied — skipped",
            "drawings": drawings,
        }
    return {
        "name": "figure-count-vs-source",
        "passed": drawings >= expected,
        "drawings": drawings,
        "source_includegraphics_count": source_count,
        "expected_min": expected,
        "failures": (
            []
            if drawings >= expected
            else [{
                "reason": "rendered drawing count is below the source includegraphics count; pandoc likely dropped figures whose path lacks an explicit extension or points to a vector file (.pdf/.vsdx) it cannot embed",
                "drawings": drawings,
                "expected_min": expected,
            }]
        ),
    }


def _table_border_kind(table_xml: str) -> str:
    """Classify a table's border configuration into one of:
       'none', 'three-line', 'full-grid', 'partial', 'inherit-style'.

       'inherit-style' means the table has a tblStyle that may carry borders,
       and we cannot determine purely from this XML what the rendered look is.
    """
    rows = re.findall(r"<w:tr\b[^>]*>.*?</w:tr>", table_xml, flags=re.S)
    if not rows:
        return "none"
    cells_in_row = lambda row: re.findall(r"<w:tc\b[^>]*>.*?</w:tc>", row, flags=re.S)
    first_cells = cells_in_row(rows[0])
    last_cells = cells_in_row(rows[-1])
    middle_cells: list[str] = []
    for r in rows[1:-1]:
        middle_cells.extend(cells_in_row(r))

    def edge_val(cell_xml: str, edge: str) -> str | None:
        bdr = re.search(r"<w:tcBorders\b[^>]*>.*?</w:tcBorders>", cell_xml, flags=re.S)
        if not bdr:
            return None
        m = re.search(rf'<w:{edge}\b[^>]*?w:val="([^"]+)"', bdr.group(0))
        return m.group(1) if m else None

    has_tbl_borders = "<w:tblBorders" in table_xml
    has_tbl_style = bool(re.search(r"<w:tblStyle\b", table_xml))
    has_any_tc_borders = "<w:tcBorders" in table_xml

    if not has_tbl_borders and not has_any_tc_borders and has_tbl_style:
        return "inherit-style"
    if not has_tbl_borders and not has_any_tc_borders:
        return "none"

    # try to detect three-line
    if first_cells and last_cells:
        first_top = edge_val(first_cells[0], "top")
        first_bottom = edge_val(first_cells[0], "bottom")
        last_bottom = edge_val(last_cells[0], "bottom")
        first_left = edge_val(first_cells[0], "left")
        first_right = edge_val(first_cells[0], "right")
        if (
            first_top in {"single", "thick"}
            and first_bottom in {"single", "thick"}
            and last_bottom in {"single", "thick"}
            and (first_left is None or first_left == "nil")
            and (first_right is None or first_right == "nil")
        ):
            # also check middle rows have no horizontal between (insideH or top/bottom = nil)
            middle_inside_ok = True
            for c in middle_cells[:6]:
                if edge_val(c, "top") == "single":
                    middle_inside_ok = False
                    break
            if middle_inside_ok:
                return "three-line"

    # full grid: every cell has top/bottom/left/right = single
    if first_cells:
        first = first_cells[0]
        if all(edge_val(first, e) in {"single", "thick"} for e in ("top", "bottom", "left", "right")):
            return "full-grid"

    return "partial"


def check_table_border_style(
    document_xml: str,
    expected_style: str,
    min_data_rows: int,
) -> dict:
    """Classify every table's border configuration. FAIL if expected_style is
    'three-line' and any data table (>= min_data_rows rows) is 'none' / 'partial'.
    Layout/wrapper tables (single-row tables with very few cells) are skipped."""
    if expected_style == "any":
        return {"name": "table-border-style", "passed": True, "evidence": "expected_style=any — skipped"}

    summary: dict[str, int] = {"three-line": 0, "full-grid": 0, "none": 0, "partial": 0, "inherit-style": 0, "skipped-layout": 0}
    failures: list[dict] = []
    tables = re.findall(r"<w:tbl\b[^>]*>.*?</w:tbl>", document_xml, flags=re.S)
    for idx, t in enumerate(tables):
        rows = re.findall(r"<w:tr\b[^>]*>.*?</w:tr>", t, flags=re.S)
        if len(rows) < min_data_rows:
            summary["skipped-layout"] += 1
            continue
        kind = _table_border_kind(t)
        summary[kind] = summary.get(kind, 0) + 1
        if expected_style == "three-line" and kind not in {"three-line", "inherit-style"}:
            failures.append({"index": idx, "kind": kind})
        elif expected_style == "full-grid" and kind not in {"full-grid", "inherit-style"}:
            failures.append({"index": idx, "kind": kind})
    return {"name": "table-border-style", "expected_style": expected_style, "summary": summary, "passed": not failures, "failures": failures}


def _count_cites_in_source(source_dir: Path) -> int:
    pat = re.compile(r"\\cite[pt]?\*?(?:\[[^\]]*\])?\{([^}]+)\}")
    n = 0
    for tex in source_dir.rglob("*.tex"):
        if tex.name in {"resume.tex", "review.tex"}:
            continue
        try:
            text = tex.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        for line in text.splitlines():
            if line.lstrip().startswith("%"):
                continue
            n += len(pat.findall(line))
    return n


def check_citation_coverage(
    document_xml: str,
    source_latex_dir: Path | None,
    explicit_min_cites: int | None,
) -> dict:
    """Compare source \\cite count to docx hyperlinks pointing at ref_* / bib_* /
    cite_* anchors. Pandoc emitted as inline (Author Year) WITHOUT hyperlink is
    the failure mode we want to catch — the rendered text looks like a
    citation but Word/PDF cannot navigate to the bibliography entry, and the
    GB/T 7714 numeric format is missing.
    """
    anchors = re.findall(r'<w:hyperlink[^>]*w:anchor="([^"]+)"', document_xml)
    cite_anchors = [a for a in anchors if re.match(r"(ref|bib|cite|biblio)[_\-]?\d+", a, re.I)]
    n_hl = len(cite_anchors)
    n_unique_anchors = len(set(cite_anchors))
    expected = explicit_min_cites
    source_count = None
    if source_latex_dir is not None and source_latex_dir.is_dir():
        source_count = _count_cites_in_source(source_latex_dir)
        if expected is None:
            expected = source_count
    if expected is None:
        return {
            "name": "citation-coverage",
            "passed": True,
            "evidence": "no --source-latex-dir or --min-citations supplied — skipped",
            "ref_hyperlinks": n_hl,
            "unique_ref_anchors": n_unique_anchors,
        }
    return {
        "name": "citation-coverage",
        "passed": n_hl >= expected,
        "ref_hyperlinks": n_hl,
        "unique_ref_anchors": n_unique_anchors,
        "source_cite_count": source_count,
        "expected_min": expected,
        "failures": (
            []
            if n_hl >= expected
            else [{
                "reason": "in-text citation hyperlinks fewer than source \\cite count; pandoc may have emitted (Author Year) tokens without internal hyperlinks. For GB/T 7714 thesis style, every \\cite must render as a numeric superscript [N] hyperlink to ref_N.",
                "ref_hyperlinks": n_hl,
                "expected_min": expected,
            }]
        ),
    }


def _count_caption_envs_in_source(source_dir: Path) -> tuple[int, int]:
    fig_count = 0
    tab_count = 0
    fig_pat = re.compile(r"\\begin\{figure\*?\}(.*?)\\end\{figure\*?\}", re.S)
    tab_pat = re.compile(r"\\begin\{table\*?\}(.*?)\\end\{table\*?\}", re.S)
    cap_pat = re.compile(r"\\caption\{")
    for tex in source_dir.rglob("*.tex"):
        if tex.name in {"resume.tex", "review.tex"}:
            continue
        try:
            text = tex.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        for m in fig_pat.finditer(text):
            if cap_pat.search(m.group(1)):
                fig_count += 1
        for m in tab_pat.finditer(text):
            if cap_pat.search(m.group(1)):
                tab_count += 1
    return fig_count, tab_count


_CAPTION_FIG_PAT = re.compile(r"^\s*(图|Figure|Fig\.?)\s*[\d一二三四五六七八九十]+[.\-][\d一二三四五六七八九十]+")
_CAPTION_TAB_PAT = re.compile(r"^\s*(表|Table|Tab\.?)\s*[\d一二三四五六七八九十]+[.\-][\d一二三四五六七八九十]+")
_CAPTION_INLINE_MENTION_PAT = re.compile(r"^\s*(图|表|Figure|Table|Fig\.?|Tab\.?)\s*[\d一二三四五六七八九十]+[.\-][\d一二三四五六七八九十]+\s*(与|和|及|或|对应|展示|显示|表明|表示|说明|描述|给出)")


def _scan_docx_caption_paragraphs(document_xml: str) -> list[dict]:
    paras = re.findall(r"<w:p\b[^>]*>.*?</w:p>", document_xml, flags=re.S)
    out = []
    for p in paras:
        txt = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)).strip()
        if not txt:
            continue
        is_fig = bool(_CAPTION_FIG_PAT.match(txt))
        is_tab = bool(_CAPTION_TAB_PAT.match(txt))
        if not (is_fig or is_tab):
            continue
        # Inline mention like '图 4.4 与图 4.3 分别给出...' should not count.
        if _CAPTION_INLINE_MENTION_PAT.match(txt):
            continue
        jc = re.search(r'<w:jc w:val="([^"]+)"', p)
        out.append({
            "kind": "figure" if is_fig else "table",
            "jc": jc.group(1) if jc else None,
            "head": txt[:60],
        })
    return out


def check_caption_count_vs_source(
    document_xml: str,
    source_latex_dir: Path | None,
    expected_min_figures: int | None,
    expected_min_tables: int | None,
) -> dict:
    """Compare source figure/table caption envs to docx caption paragraphs."""
    caps = _scan_docx_caption_paragraphs(document_xml)
    fig_caps = [c for c in caps if c["kind"] == "figure"]
    tab_caps = [c for c in caps if c["kind"] == "table"]
    if source_latex_dir is not None and source_latex_dir.is_dir():
        sf, st = _count_caption_envs_in_source(source_latex_dir)
        if expected_min_figures is None:
            expected_min_figures = sf
        if expected_min_tables is None:
            expected_min_tables = st
    if expected_min_figures is None and expected_min_tables is None:
        return {
            "name": "caption-count-vs-source",
            "passed": True,
            "evidence": "no --source-latex-dir / --min-figure-captions / --min-table-captions — skipped",
            "docx_figure_captions": len(fig_caps),
            "docx_table_captions": len(tab_caps),
        }
    failures = []
    if expected_min_figures is not None and len(fig_caps) < expected_min_figures:
        failures.append({
            "reason": "fewer figure caption paragraphs than source \\begin{figure} envs with \\caption",
            "docx_count": len(fig_caps),
            "expected_min": expected_min_figures,
        })
    if expected_min_tables is not None and len(tab_caps) < expected_min_tables:
        failures.append({
            "reason": "fewer table caption paragraphs than source \\begin{table} envs with \\caption",
            "docx_count": len(tab_caps),
            "expected_min": expected_min_tables,
        })
    return {
        "name": "caption-count-vs-source",
        "passed": not failures,
        "docx_figure_captions": len(fig_caps),
        "docx_table_captions": len(tab_caps),
        "expected_figure_min": expected_min_figures,
        "expected_table_min": expected_min_tables,
        "failures": failures,
    }


def check_caption_numbering(document_xml: str) -> dict:
    """Every caption-bearing paragraph must match the strict numbering form
    '图 X.Y  说明文字' / '表 X.Y  说明文字' (or 'Figure N.M' / 'Table N.M' for
    English templates) with non-empty trailing description text.

    A paragraph that *looks* like a caption attempt but lacks the numeric
    N.M pair, or whose number prefix is followed by an empty description,
    fails this check. We only consider a paragraph a candidate when it begins
    with '图/表/Figure/Table' AND the next non-space char is a digit — body
    sentences like '表格型 Q-learning ...' or '表中的实验配置 ...' must NOT
    be treated as captions just because they happen to start with 表/图.
    """
    paras = re.findall(r"<w:p\b[^>]*>.*?</w:p>", document_xml, flags=re.S)
    # Candidate: '图/表/Figure/Table' + optional whitespace + a digit (any digit).
    # This excludes body words '表格', '表中', '图示', '图谱', etc.
    weak_pat = re.compile(r"^\s*(图|表|Figure|Fig\.?|Table|Tab\.?)\s*[\d一二三四五六七八九十]")
    strict_pat = re.compile(r"^\s*(图|表|Figure|Fig\.?|Table|Tab\.?)\s*\d+[.\-]\d+\s*\S+")
    failures = []
    total_attempted = 0
    for p in paras:
        txt = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)).strip()
        if not txt:
            continue
        if not weak_pat.match(txt):
            continue
        # Skip inline body mentions like '图 4.4 与图 4.3 分别给出...'
        if _CAPTION_INLINE_MENTION_PAT.match(txt):
            continue
        total_attempted += 1
        if not strict_pat.match(txt):
            failures.append({"reason": "caption paragraph missing 'X.Y  说明文字' format", "head": txt[:60]})
    return {
        "name": "caption-numbering",
        "passed": not failures,
        "candidate_caption_paragraphs": total_attempted,
        "failures": failures,
        "failures_sample": failures[:5],
    }


def check_caption_centering(document_xml: str) -> dict:
    """Every paragraph already recognized as a caption must have w:jc=center."""
    caps = _scan_docx_caption_paragraphs(document_xml)
    not_centered = [c for c in caps if (c.get("jc") or "").lower() != "center"]
    return {
        "name": "caption-centering",
        "passed": not not_centered,
        "total_caption_paragraphs": len(caps),
        "not_centered_count": len(not_centered),
        "not_centered_sample": not_centered[:5],
        "failures": (
            []
            if not not_centered
            else [{"reason": "caption paragraph(s) without w:jc=center", "count": len(not_centered)}]
        ),
    }


def check_caption_untagged_near_figure_table(document_xml: str) -> dict:
    """Every body <w:drawing> paragraph should be followed by a caption-pattern
    paragraph (within 1 step), and every data <w:tbl> should be preceded by
    one. A paragraph in that slot whose text does not start with the
    figure/table caption pattern is a missing or untagged caption — common
    when pandoc maps \\caption{} text to a plain body style without the
    'X.Y' prefix and the project pipeline forgets to repair it.

    Front-matter scope: cover-page logos, school-crest images, and other
    drawings that appear *before* the first body Heading 1 are not
    figures-with-captions and are excluded from this check.
    """
    paras_and_tbls = re.findall(r"<w:p\b[^>]*>.*?</w:p>|<w:tbl>.*?</w:tbl>", document_xml, flags=re.S)
    is_drawing = lambda x: x.startswith("<w:p") and "<w:drawing" in x
    is_tbl = lambda x: x.startswith("<w:tbl>")
    is_caption = lambda x: x.startswith("<w:p") and bool(_CAPTION_FIG_PAT.match("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", x)).strip()) or _CAPTION_TAB_PAT.match("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", x)).strip())) and not _CAPTION_INLINE_MENTION_PAT.match("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", x)).strip())
    # Locate the first body Heading-1-style paragraph; treat everything before
    # it as front-matter (cover/abstract/declaration). We accept either the
    # localized pStyle name (Heading1 / 1 / 标题 1) or a numPr binding.
    body_start = 0
    for idx, x in enumerate(paras_and_tbls):
        if not x.startswith("<w:p"):
            continue
        sid_m = re.search(r"<w:pStyle w:val=\"([^\"]+)\"", x)
        sid = sid_m.group(1) if sid_m else ""
        # Heuristic: chapter-level heading style ids in the school template.
        if sid in {"1", "Heading1", "a3", "Heading10", "1Char"} or "标题1" in sid or "Heading1" in sid:
            body_start = idx
            break
    failures = []
    for i, x in enumerate(paras_and_tbls):
        if i < body_start:
            continue
        if is_drawing(x):
            # Look ahead: skip over any subsequent drawing paragraphs (subfigure
            # group). After the last drawing in the group, the next non-drawing
            # paragraph must be a caption.
            j = i + 1
            while j < len(paras_and_tbls) and is_drawing(paras_and_tbls[j]):
                j += 1
            if j >= len(paras_and_tbls) or not is_caption(paras_and_tbls[j]):
                failures.append({
                    "kind": "drawing-without-following-caption",
                    "drawing_index": i,
                })
        if is_tbl(x):
            # Skip layout wrappers (single-row, ≤2 cells)
            rows = re.findall(r"<w:tr\b[^>]*>.*?</w:tr>", x, flags=re.S)
            if len(rows) < 2:
                continue
            first_cells = re.findall(r"<w:tc\b[^>]*>", rows[0])
            if len(first_cells) < 2:
                continue
            # Scan backwards for the nearest caption paragraph
            k = i - 1
            while k >= 0 and paras_and_tbls[k].startswith("<w:p") and not "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", paras_and_tbls[k])).strip():
                k -= 1
            if k < 0 or not is_caption(paras_and_tbls[k]):
                failures.append({
                    "kind": "table-without-preceding-caption",
                    "table_index": i,
                })
    return {
        "name": "caption-untagged-near-figure-table",
        "passed": not failures,
        "front_matter_skipped_until_index": body_start,
        "failures_count": len(failures),
        "failures_sample": failures[:5],
        "failures": failures,
    }


# ---- main ----

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("docx")
    parser.add_argument("--pdf", default=None)
    parser.add_argument("--out", default=None)
    parser.add_argument("--allow", action="append", choices=CHECK_NAMES, default=[],
                        help="Demote a check from FAIL to WARN.")
    parser.add_argument("--reference-marker", action="append", default=["参考文献", "References", "Bibliography", "参 考 文 献"])
    parser.add_argument("--back-matter-title", action="append", default=DEFAULT_BACK_MATTER_TITLES)
    parser.add_argument("--chapter-prefix-pattern", default=DEFAULT_CHAPTER_PREFIX_PATTERN)
    parser.add_argument("--multilevel-pattern", default=DEFAULT_MULTILEVEL_PATTERN)
    parser.add_argument("--source-latex-dir", default=None,
                        help="Directory of LaTeX sources (.tex files); enables figure-count-vs-source check.")
    parser.add_argument("--min-figures", type=int, default=None,
                        help="Explicit lower bound for <w:drawing> count. Used when --source-latex-dir not given.")
    parser.add_argument("--expected-table-style", default="any",
                        choices=("any", "three-line", "full-grid"),
                        help="Required table border style. Default 'any' = no enforcement.")
    parser.add_argument("--table-min-data-rows", type=int, default=2,
                        help="Tables with fewer rows are skipped as layout/wrapper.")
    parser.add_argument("--min-citations", type=int, default=None,
                        help="Explicit lower bound for cite hyperlinks. Used when --source-latex-dir not given.")
    parser.add_argument("--min-figure-captions", type=int, default=None,
                        help="Explicit lower bound for figure caption paragraphs. Used when --source-latex-dir not given.")
    parser.add_argument("--min-table-captions", type=int, default=None,
                        help="Explicit lower bound for table caption paragraphs. Used when --source-latex-dir not given.")
    args = parser.parse_args()

    docx_path = Path(args.docx)
    if not docx_path.is_file():
        parser.error(f"docx not found: {docx_path}")

    with zipfile.ZipFile(str(docx_path)) as zf:
        document_xml = read_part(zf, "word/document.xml")
        styles_xml = read_part(zf, "word/styles.xml")
        numbering_xml = read_part(zf, "word/numbering.xml")
        rels_xml = read_part(zf, "word/_rels/document.xml.rels")
        rels_map = parse_doc_rels(rels_xml)
        num_to_abstract, abstract_levels = parse_numbering(numbering_xml)

        results = [
            check_toc_field(document_xml),
            check_numbering_consistency(document_xml, styles_xml, num_to_abstract, abstract_levels),
            check_multilevel_headings(
                document_xml, styles_xml, num_to_abstract, abstract_levels,
                chapter_prefix_pattern=args.chapter_prefix_pattern,
                multilevel_pattern=args.multilevel_pattern,
                back_matter_titles=args.back_matter_title,
            ),
            check_ref_counter_independence(document_xml, styles_xml, args.reference_marker),
            check_body_header_non_back_matter(zf, document_xml, rels_map, args.back_matter_title),
            check_figure_count_vs_source(
                document_xml,
                Path(args.source_latex_dir) if args.source_latex_dir else None,
                args.min_figures,
            ),
            check_table_border_style(
                document_xml,
                expected_style=args.expected_table_style,
                min_data_rows=args.table_min_data_rows,
            ),
            check_citation_coverage(
                document_xml,
                Path(args.source_latex_dir) if args.source_latex_dir else None,
                args.min_citations,
            ),
            check_caption_count_vs_source(
                document_xml,
                Path(args.source_latex_dir) if args.source_latex_dir else None,
                args.min_figure_captions,
                args.min_table_captions,
            ),
            check_caption_numbering(document_xml),
            check_caption_centering(document_xml),
            check_caption_untagged_near_figure_table(document_xml),
        ]
        if args.pdf:
            results.append(check_pdf_field_errors(Path(args.pdf)))

    failures: list[str] = []
    warnings: list[str] = []
    for r in results:
        if r.get("passed"):
            continue
        msg = f"{r['name']}: {len(r.get('failures', [])) or 1} issue(s)"
        if r["name"] in args.allow:
            warnings.append(msg)
        else:
            failures.append(msg)

    status = "PASS" if not failures else "FAIL"
    out = {
        "docx": str(docx_path),
        "pdf": args.pdf,
        "checks": results,
        "status": status,
        "failures": failures,
        "warnings": warnings,
    }
    text = json.dumps(out, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
    print(f"STATUS: {status}")
    for r in results:
        flag = "PASS" if r.get("passed") else ("WARN" if r["name"] in args.allow else "FAIL")
        print(f"- [{flag}] {r['name']}")
    if args.out:
        print(args.out)
    return 0 if status == "PASS" else 1


if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        if hasattr(_s, "reconfigure"):
            _s.reconfigure(encoding="utf-8", errors="replace")

    raise SystemExit(main())
