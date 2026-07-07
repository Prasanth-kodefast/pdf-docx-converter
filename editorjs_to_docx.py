"""
editorjs_to_docx
================

Render a CLM contract (per-page EditorJS blocks + per-page headers /
footers) into a real `.docx` file where the sender's redline marks
become native Word Track Changes:

  <ins class="clm-suggestion-insert">TEXT</ins>  →  <w:ins ...><w:r><w:t>TEXT
  <mark class="clm-redline">TEXT</mark>          →  <w:del ...><w:r><w:delText>TEXT

Once the counterparty opens the file in Microsoft Word (or LibreOffice
or Google Docs), the Review ribbon lights up and they can accept /
reject / add further changes exactly the same way they would with any
tracked-changes document from a peer.

This is the enterprise-standard flow. Ironclad + DocuSign CLM Negotiate
do the same round-trip; the counterparty never needs to log into a web
app or install a Word plugin — they just work in Word.

Public entry point:

    render_contract_docx(pages, meta) -> bytes

`pages` is a list shaped `[{ pageNo, blocks }]` where `blocks` are raw
EditorJS block objects.  `meta` carries the title, sender identity for
`w:author`, an optional cover message, and optional per-page header /
footer maps keyed by page index.

Returns the assembled DOCX as an in-memory bytestring so the FastAPI
handler can stream it back to the caller.
"""

from __future__ import annotations

import io
import re
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, List, Optional

from bs4 import BeautifulSoup, NavigableString
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


# ─── OOXML namespaces ───────────────────────────────────────────────────────
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ─── Revision id counter ────────────────────────────────────────────────────
class _RevisionCounter:
    """Word requires every `w:ins` / `w:del` to carry a unique `w:id`.
    A shared counter across a whole document keeps ids monotonic, which
    is what every serializer we've inspected (Word, LibreOffice, Google
    Docs) produces."""

    def __init__(self, start: int = 1) -> None:
        self._n = start - 1

    def next(self) -> str:
        self._n += 1
        return str(self._n)


# ─── Inline-run parser ──────────────────────────────────────────────────────
#
# Given a paragraph's inline HTML (as produced by EditorJS with the
# sender's redline marks + mentions + clause chips baked in), walk it
# and emit a stream of typed segments the paragraph renderer knows how
# to attach to a `w:p` element:
#
#   { "kind": "text", "text": str, "bold": bool, "italic": bool,
#     "underline": bool, "strike": bool, "color": Optional[str] }
#   { "kind": "ins",   "runs": [text-run-dicts] }
#   { "kind": "del",   "runs": [text-run-dicts] }
#   { "kind": "linebreak" }
#
# The parser is deliberately depth-first + accumulates the current
# formatting state as it descends, so nested `<strong><ins>foo</ins></strong>`
# still lands in Word as bold + tracked insertion.


_TRACKED_INS_CLASS = "clm-suggestion-insert"
_TRACKED_DEL_CLASS = "clm-redline"

# Elements we treat as "not visible content" — mentions and chip carriers.
# Their visible text is what the counterparty should see; the HTML plumbing
# around them is stripped.
_CHIP_CLASSES = {
    "mention",
    "clause-var-chip",
    "clause-tag-chip",  # hidden carrier — never emit any text for these
}


def _class_list(tag) -> List[str]:
    cls = tag.get("class") if hasattr(tag, "get") else None
    if not cls:
        return []
    if isinstance(cls, str):
        return cls.split()
    return list(cls)


def _has_class(tag, name: str) -> bool:
    return name in _class_list(tag)


def _style_dict(tag) -> Dict[str, str]:
    raw = tag.get("style") if hasattr(tag, "get") else None
    if not raw:
        return {}
    out: Dict[str, str] = {}
    for chunk in raw.split(";"):
        if ":" not in chunk:
            continue
        k, v = chunk.split(":", 1)
        out[k.strip().lower()] = v.strip()
    return out


def _hex_from_color(css: str) -> Optional[str]:
    """Coerce a CSS color value (`#rrggbb`, `rgb(r,g,b)`, name) into a
    6-char hex string usable by python-docx's `RGBColor`. Returns
    None when the color can't be parsed — the run just falls back to
    the paragraph default."""
    if not css:
        return None
    css = css.strip()
    if css.startswith("#") and (len(css) == 7 or len(css) == 4):
        return css[1:].upper() if len(css) == 7 else "".join(c * 2 for c in css[1:]).upper()
    m = re.match(r"rgb\((\d+)\s*,\s*(\d+)\s*,\s*(\d+)\)", css)
    if m:
        r, g, b = (int(x) for x in m.groups())
        return f"{r:02X}{g:02X}{b:02X}"
    return None


def _parse_inline(
    html: str,
) -> List[Dict[str, Any]]:
    """Walk the inline HTML of a block and return a flat sequence of
    typed segments (text runs, ins groups, del groups, line breaks).

    The returned list is what `_apply_runs_to_paragraph` consumes; the
    two functions are the sole boundary between "HTML I don't want
    inside my Word renderer" and "runs I know how to attach."
    """
    if not html:
        return []
    soup = BeautifulSoup(html, "html.parser")
    segments: List[Dict[str, Any]] = []

    def walk(node, fmt: Dict[str, Any]):
        # Text node — emit a run.
        if isinstance(node, NavigableString):
            text = str(node)
            if text:
                segments.append({"kind": "text", "text": text, **fmt})
            return

        # Line break.
        if getattr(node, "name", None) == "br":
            segments.append({"kind": "linebreak"})
            return

        # Tracked insertion / deletion — recurse with a marker so the
        # parent emits `w:ins` / `w:del` around the collected runs.
        cls = _class_list(node)
        if node.name == "ins" and _TRACKED_INS_CLASS in cls:
            child_runs: List[Dict[str, Any]] = []
            _walk_into(node, fmt, into=child_runs)
            segments.append({"kind": "ins", "runs": child_runs})
            return
        if node.name in ("mark", "del") and _TRACKED_DEL_CLASS in cls:
            child_runs: List[Dict[str, Any]] = []
            _walk_into(node, fmt, into=child_runs)
            segments.append({"kind": "del", "runs": child_runs})
            return

        # Chip carriers — hidden cite tag we always drop.
        if node.name == "cite" and _has_class(node, "clause-tag-chip"):
            return

        # Mention / clause-var-chip — pass through the visible text,
        # ignore the wrapping element. The value the counterparty sees
        # is whatever text-content the chip renders.
        if _has_class(node, "mention") or _has_class(node, "clause-var-chip"):
            for child in node.children:
                walk(child, fmt)
            return

        # Inline formatting — accumulate onto `fmt` and recurse.
        next_fmt = dict(fmt)
        if node.name in ("strong", "b"):
            next_fmt["bold"] = True
        if node.name in ("em", "i"):
            next_fmt["italic"] = True
        if node.name == "u":
            next_fmt["underline"] = True
        if node.name == "s":
            next_fmt["strike"] = True
        if node.name == "sup":
            next_fmt["superscript"] = True
        if node.name == "sub":
            next_fmt["subscript"] = True
        if node.name == "a":
            next_fmt["hyperlink"] = node.get("href")

        style = _style_dict(node)
        if "color" in style:
            hex_c = _hex_from_color(style["color"])
            if hex_c:
                next_fmt["color"] = hex_c
        if style.get("font-weight") in ("bold", "700", "800", "900"):
            next_fmt["bold"] = True
        if style.get("font-style") == "italic":
            next_fmt["italic"] = True
        if style.get("text-decoration") == "underline":
            next_fmt["underline"] = True
        if style.get("text-decoration") == "line-through":
            next_fmt["strike"] = True

        for child in node.children:
            walk(child, next_fmt)

    def _walk_into(container, fmt: Dict[str, Any], into: List[Dict[str, Any]]):
        nonlocal segments
        saved = segments
        segments = into
        try:
            for child in container.children:
                walk(child, fmt)
        finally:
            segments = saved

    for child in soup.children:
        walk(child, {})
    return segments


# ─── OOXML helpers ──────────────────────────────────────────────────────────
def _make_rPr(run_props: Dict[str, Any]) -> Optional[OxmlElement]:
    """Build a `<w:rPr>` element carrying the formatting for one text
    run. Returns None when the run has no formatting overrides — the
    caller can then just omit rPr entirely, matching what Word does."""
    if not any(run_props.get(k) for k in ("bold", "italic", "underline", "strike", "color", "superscript", "subscript")):
        return None
    rpr = OxmlElement("w:rPr")
    if run_props.get("bold"):
        b = OxmlElement("w:b")
        rpr.append(b)
    if run_props.get("italic"):
        i = OxmlElement("w:i")
        rpr.append(i)
    if run_props.get("underline"):
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rpr.append(u)
    if run_props.get("strike"):
        s = OxmlElement("w:strike")
        rpr.append(s)
    if run_props.get("superscript"):
        v = OxmlElement("w:vertAlign")
        v.set(qn("w:val"), "superscript")
        rpr.append(v)
    if run_props.get("subscript"):
        v = OxmlElement("w:vertAlign")
        v.set(qn("w:val"), "subscript")
        rpr.append(v)
    if run_props.get("color"):
        c = OxmlElement("w:color")
        c.set(qn("w:val"), run_props["color"])
        rpr.append(c)
    return rpr


def _make_run(seg: Dict[str, Any], *, deleted: bool = False) -> OxmlElement:
    """Build a `<w:r>` (run) element. When `deleted=True` the text
    child is `<w:delText>` instead of `<w:t>` — Word requires this
    inside `<w:del>` revisions."""
    r = OxmlElement("w:r")
    rpr = _make_rPr(seg)
    if rpr is not None:
        r.append(rpr)
    text_tag = "w:delText" if deleted else "w:t"
    t = OxmlElement(text_tag)
    t.text = seg.get("text", "")
    # Preserve leading / trailing whitespace exactly as authored —
    # Word collapses it otherwise. This mirrors how Word itself
    # serialises runs it emits from the Review ribbon.
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    return r


def _make_linebreak_run() -> OxmlElement:
    r = OxmlElement("w:r")
    r.append(OxmlElement("w:br"))
    return r


def _make_ins(runs: List[Dict[str, Any]], author: str, when: str, rev: _RevisionCounter) -> OxmlElement:
    """Build a `<w:ins>` (tracked insertion) element wrapping the given
    runs. Word's Review ribbon reads `w:author` + `w:date` to attribute
    the change and the timestamp shown on hover."""
    ins = OxmlElement("w:ins")
    ins.set(qn("w:id"), rev.next())
    ins.set(qn("w:author"), author)
    ins.set(qn("w:date"), when)
    for seg in runs:
        if seg.get("kind") == "linebreak":
            ins.append(_make_linebreak_run())
        else:
            ins.append(_make_run(seg))
    return ins


def _make_del(runs: List[Dict[str, Any]], author: str, when: str, rev: _RevisionCounter) -> OxmlElement:
    """Build a `<w:del>` (tracked deletion) element. Deletions use
    `<w:delText>` for the text child; if you accidentally emit `<w:t>`
    inside `w:del` Word will still open the file but won't recognise
    the deletion as a real revision."""
    d = OxmlElement("w:del")
    d.set(qn("w:id"), rev.next())
    d.set(qn("w:author"), author)
    d.set(qn("w:date"), when)
    for seg in runs:
        if seg.get("kind") == "linebreak":
            d.append(_make_linebreak_run())
        else:
            d.append(_make_run(seg, deleted=True))
    return d


def _apply_segments_to_paragraph(
    paragraph, segments: List[Dict[str, Any]], author: str, when: str, rev: _RevisionCounter
) -> None:
    """Attach a stream of typed segments to a python-docx paragraph.
    Text runs go through python-docx's ordinary API; `ins` / `del`
    segments hand-craft the OOXML because python-docx doesn't expose a
    high-level API for revisions."""
    p_element = paragraph._p
    for seg in segments:
        kind = seg.get("kind")
        if kind == "text":
            p_element.append(_make_run(seg))
        elif kind == "linebreak":
            p_element.append(_make_linebreak_run())
        elif kind == "ins":
            p_element.append(_make_ins(seg["runs"], author, when, rev))
        elif kind == "del":
            p_element.append(_make_del(seg["runs"], author, when, rev))


# ─── Block renderer ─────────────────────────────────────────────────────────
_ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _apply_alignment(paragraph, block: Dict[str, Any]) -> None:
    tunes = block.get("tunes") or {}
    align = ((tunes.get("alignment") or {}).get("alignment")) or "left"
    paragraph.alignment = _ALIGN_MAP.get(align, WD_ALIGN_PARAGRAPH.LEFT)


def _render_block(doc, block: Dict[str, Any], author: str, when: str, rev: _RevisionCounter) -> None:
    """Attach one EditorJS block to the given docx Document. Unknown
    block types fall back to a plain paragraph carrying the block's
    `text` field, so unsupported blocks never silently drop content."""
    if not isinstance(block, dict):
        return
    btype = (block.get("type") or "paragraph").lower()
    data = block.get("data") or {}
    text = data.get("text") if isinstance(data.get("text"), str) else ""

    if btype == "header":
        level = max(1, min(6, int(data.get("level") or 2)))
        p = doc.add_paragraph()
        _apply_alignment(p, block)
        # Word maps Heading N styles onto its outline levels; python-docx
        # exposes them under those names.
        p.style = doc.styles[f"Heading {level}"]
        _apply_segments_to_paragraph(p, _parse_inline(text), author, when, rev)
        return

    if btype in ("paragraph", "quote"):
        p = doc.add_paragraph()
        _apply_alignment(p, block)
        if btype == "quote":
            p.style = doc.styles["Intense Quote"]
        _apply_segments_to_paragraph(p, _parse_inline(text), author, when, rev)
        return

    if btype == "list":
        style_name = "List Number" if (data.get("style") == "ordered") else "List Bullet"
        items = data.get("items") or []
        for item in items:
            item_text = item if isinstance(item, str) else (item.get("content") or item.get("text") or "")
            p = doc.add_paragraph(style=style_name)
            _apply_alignment(p, block)
            _apply_segments_to_paragraph(p, _parse_inline(item_text), author, when, rev)
        return

    if btype == "checklist":
        items = data.get("items") or []
        for item in items:
            checked = bool(item.get("checked")) if isinstance(item, dict) else False
            marker = "☑ " if checked else "☐ "
            item_text = item.get("text") if isinstance(item, dict) else str(item or "")
            p = doc.add_paragraph()
            _apply_alignment(p, block)
            p.add_run(marker)
            _apply_segments_to_paragraph(p, _parse_inline(item_text or ""), author, when, rev)
        return

    if btype == "delimiter":
        # A centered dot row — Word doesn't have a first-class "delimiter"
        # so we replicate what most authoring tools do for section
        # breaks in narrative documents.
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("• • •")
        r.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        return

    if btype == "table":
        content = data.get("content") or []
        if not content or not isinstance(content, list):
            return
        rows = len(content)
        cols = max((len(r) for r in content if isinstance(r, list)), default=0)
        if cols == 0:
            return
        tbl = doc.add_table(rows=rows, cols=cols)
        tbl.style = "Table Grid"
        with_headings = bool(data.get("withHeadings"))
        for ri, row in enumerate(content):
            if not isinstance(row, list):
                continue
            for ci, cell in enumerate(row):
                if ci >= cols:
                    break
                cell_text = cell if isinstance(cell, str) else ""
                tcell = tbl.cell(ri, ci)
                # First paragraph created by add_table is empty; overwrite it.
                tcell.text = ""
                p = tcell.paragraphs[0]
                if with_headings and ri == 0:
                    p.style = doc.styles["Heading 4"]
                _apply_segments_to_paragraph(p, _parse_inline(cell_text), author, when, rev)
        return

    if btype == "image":
        # We don't embed remote images (would require fetching), but we
        # preserve the URL as a caption paragraph so nothing is silently
        # dropped. The caller can swap in an inline image later if we
        # decide to support that.
        src = ((data.get("file") or {}).get("url")) or data.get("url") or ""
        caption = data.get("caption") or ""
        if src:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(f"[image: {src}]")
            r.italic = True
            r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        if caption:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _apply_segments_to_paragraph(p, _parse_inline(caption), author, when, rev)
        return

    # Fallback for unknown block types — never silently drop content.
    if text:
        p = doc.add_paragraph()
        _apply_alignment(p, block)
        _apply_segments_to_paragraph(p, _parse_inline(text), author, when, rev)


# ─── Headers / footers ──────────────────────────────────────────────────────
def _normalise_hf(raw: Any) -> Dict[int, List[Dict[str, Any]]]:
    """Flatten the response's header / footer container (which can
    arrive as a dict, a list of {pageNo, blocks}, or Mongoose's Map
    with a `.entries()` method) into a plain `{page_index: blocks}`
    mapping keyed by 0-based page index."""
    if raw is None:
        return {}
    out: Dict[int, List[Dict[str, Any]]] = {}
    if isinstance(raw, list):
        for entry in raw:
            if not isinstance(entry, dict):
                continue
            idx = entry.get("pageNo") if "pageNo" in entry else entry.get("page_no")
            blocks = entry.get("blocks") or []
            try:
                out[int(idx)] = list(blocks) if isinstance(blocks, list) else []
            except (TypeError, ValueError):
                continue
        return out
    if isinstance(raw, dict):
        for k, v in raw.items():
            try:
                idx = int(k)
            except (TypeError, ValueError):
                continue
            if isinstance(v, dict) and isinstance(v.get("blocks"), list):
                out[idx] = list(v["blocks"])
            elif isinstance(v, list):
                out[idx] = list(v)
        return out
    return out


def _fill_hf_from_blocks(container, blocks: List[Dict[str, Any]], author: str, when: str, rev: _RevisionCounter) -> None:
    """Populate a header or footer container (both look the same to
    python-docx) with our block renderer. python-docx creates a single
    empty paragraph in the header/footer by default; we clear it first
    so nothing overlaps."""
    if not blocks:
        return
    for p in list(container.paragraphs):
        p._element.getparent().remove(p._element)
    # Route header/footer blocks through the same renderer as the body so
    # `<ins>` / `<del>` marks survive there too. If the contract has
    # different headers on different pages, that's handled by adding one
    # section per page with `is_linked_to_previous = False`.
    dummy_doc = container  # for the renderer we just need `.add_paragraph`
    for block in blocks:
        _render_hf_block(container, block, author, when, rev)


def _render_hf_block(container, block: Dict[str, Any], author: str, when: str, rev: _RevisionCounter) -> None:
    """Cut-down variant of `_render_block` that writes into a header /
    footer container (not the main body). Headers/footers can hold
    paragraphs, images, and tables — everything else falls back to
    paragraph text."""
    if not isinstance(block, dict):
        return
    btype = (block.get("type") or "paragraph").lower()
    data = block.get("data") or {}
    text = data.get("text") if isinstance(data.get("text"), str) else ""
    if btype == "header":
        level = max(1, min(6, int(data.get("level") or 2)))
        p = container.add_paragraph()
        p.style = container.part.document.styles[f"Heading {level}"]
        _apply_alignment(p, block)
        _apply_segments_to_paragraph(p, _parse_inline(text), author, when, rev)
        return
    # Everything else lands as a paragraph — headers/footers rarely need
    # tables/lists, and if we hit an edge case a plain paragraph is
    # graceful degradation.
    p = container.add_paragraph()
    _apply_alignment(p, block)
    _apply_segments_to_paragraph(p, _parse_inline(text), author, when, rev)


# ─── Public entry point ─────────────────────────────────────────────────────
def render_contract_docx(
    pages: List[Dict[str, Any]],
    meta: Dict[str, Any],
) -> bytes:
    """Assemble a Track-Changes-enabled DOCX for the given contract and
    return the file bytes.

    Arguments
    ---------
    pages : list of dicts shaped `{ "pageNo": int, "blocks": [...] }`
        The per-page EditorJS block dumps. Sorted by `pageNo` on entry
        so the caller doesn't have to.
    meta : dict
        Required keys:
          - "title" (str)  — used for the Word document title property
                             and rendered as an H1 at the top.
          - "author" (str) — populates `w:author` on every `w:ins` /
                             `w:del`; typically the sender's full name
                             (or their email if no name).
        Optional keys:
          - "message" (str)  — cover blurb rendered above the body.
          - "date" (str, ISO)— used for `w:date` on revisions; defaults
                               to now in UTC.
          - "headers" / "footers" — `{ pageIndex: {blocks} }` maps for
                               per-page header / footer content.
          - "enable_track_changes" (bool, default True) — whether to
                                open the file with Track Changes ON so
                                the counterparty's further edits get
                                recorded too.
    """
    title = str(meta.get("title") or "Contract")
    author = str(meta.get("author") or "Sender")
    message = meta.get("message") or ""
    when = meta.get("date") or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    enable_tc = bool(meta.get("enable_track_changes", True))
    headers = _normalise_hf(meta.get("headers"))
    footers = _normalise_hf(meta.get("footers"))

    rev = _RevisionCounter()
    doc = Document()

    # Metadata — title + author show up in Word's document properties
    # dialog and on the recent-files hover cards.
    doc.core_properties.title = title
    doc.core_properties.author = author

    # Sort pages so page 1 lands first even when the caller passes an
    # unordered list.
    sorted_pages = sorted(pages or [], key=lambda p: int(p.get("pageNo") or 0))

    # Force Word to open the document with Track Changes turned ON so
    # any FURTHER edits the counterparty makes (beyond accepting/rejecting
    # what the sender sent) also get tracked. Without this Word opens
    # in normal editing mode and the counterparty's straight-typed
    # changes disappear into the file without a revision marker.
    if enable_tc:
        settings = doc.settings.element
        track = OxmlElement("w:trackChanges")
        settings.append(track)

    # Cover — title + optional intro paragraph. Kept small so the actual
    # contract sits at the top of page 1.
    title_p = doc.add_paragraph()
    title_p.style = doc.styles["Title"]
    title_p.add_run(title)
    if message:
        p = doc.add_paragraph()
        p.style = doc.styles["Intense Quote"]
        p.add_run(str(message))

    # Attach per-page header + footer via the section headers/footers.
    # Word associates one hdr/ftr per section, so when we detect
    # per-page differences we insert a section break at the top of the
    # affected page and unlink its header from the previous section.
    section = doc.sections[0]

    # Common case: all pages share page-0's header/footer (or the doc
    # has a single set that isn't keyed by page index at all). Attach
    # them to section 0 and let Word repeat them across every page.
    default_header_blocks = headers.get(0) or (list(headers.values())[0] if headers else [])
    default_footer_blocks = footers.get(0) or (list(footers.values())[0] if footers else [])
    if default_header_blocks:
        _fill_hf_from_blocks(section.header, default_header_blocks, author, when, rev)
    if default_footer_blocks:
        _fill_hf_from_blocks(section.footer, default_footer_blocks, author, when, rev)

    # Emit the body. Page N (1-indexed for the user, 0-indexed in our
    # data) gets an explicit page-break-before once we're past page 1.
    for idx, page in enumerate(sorted_pages):
        page_no = int(page.get("pageNo") or idx)
        blocks = page.get("blocks") or []
        if idx > 0:
            # A `w:br` with `w:type="page"` inside a run is Word's
            # explicit page break marker.
            p = doc.add_paragraph()
            r = p.add_run()
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            r._r.append(br)

        # Per-page header/footer overrides — only if a specific per-page
        # override exists AND it differs from the default. Adding a new
        # section for every page bloats the file, so we only pay the
        # cost when it's actually different content.
        override_header = headers.get(page_no)
        override_footer = footers.get(page_no)
        if idx > 0 and (
            (override_header and override_header is not default_header_blocks)
            or (override_footer and override_footer is not default_footer_blocks)
        ):
            new_section = doc.add_section()
            new_section.header.is_linked_to_previous = False
            new_section.footer.is_linked_to_previous = False
            if override_header:
                _fill_hf_from_blocks(new_section.header, override_header, author, when, rev)
            if override_footer:
                _fill_hf_from_blocks(new_section.footer, override_footer, author, when, rev)

        for block in blocks:
            _render_block(doc, block, author, when, rev)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


__all__ = ["render_contract_docx"]
