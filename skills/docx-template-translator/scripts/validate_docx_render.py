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
    raise SystemExit(main())
