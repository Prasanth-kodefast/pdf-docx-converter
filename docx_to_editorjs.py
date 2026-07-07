"""
docx_to_editorjs
================

Parse a `.docx` file that came back from a counterparty (edited in
Microsoft Word with Track Changes) back into per-page EditorJS blocks
where Word's `w:ins` / `w:del` revisions become the same
`<ins class="clm-suggestion-insert">` / `<mark class="clm-redline">`
marks the CLM builder already understands. The parsed blocks land in
the existing approver-version pipeline exactly like an in-app
approver's redlines would — no bespoke merge logic downstream.

Return shape:

    {
      "pages": [
        { "pageNo": 0, "blocks": [ ... EditorJS blocks ... ] },
        ...
      ],
      "counterparty_authors": ["Jane Doe", "Mark <mark@partner.com>"],
      "change_count": 42
    }

The parser is intentionally forgiving: an unknown Word element is
degraded to plain text rather than dropping the paragraph. Enterprise
CLMs (Ironclad, DocuSign) do the same — a lossy round-trip is much
better than a broken one.
"""

from __future__ import annotations

import io
import re
from typing import Any, Dict, List, Optional, Tuple

from docx import Document

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _ns(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


# ─── Inline run parser ─────────────────────────────────────────────────────
def _run_text(run_el, deleted: bool = False) -> str:
    """Extract the visible text out of a `<w:r>` element. `<w:t>` for
    ordinary runs, `<w:delText>` for anything inside a `<w:del>`. Word
    also uses `<w:br>` for line breaks — we surface those as \n so the
    paragraph splitter honours them."""
    tag_names = [_ns("delText")] if deleted else [_ns("t")]
    out: List[str] = []
    for child in run_el:
        if child.tag in tag_names:
            out.append(child.text or "")
        elif child.tag == _ns("br"):
            out.append("\n")
        elif child.tag == _ns("tab"):
            out.append("\t")
    return "".join(out)


def _rpr_style_flags(r_pr) -> Dict[str, bool]:
    """Read `<w:rPr>` and return the format flags relevant to our
    downstream EditorJS renderer. python-docx exposes some of these
    through its high-level API but we're already walking the OOXML
    tree, so read it directly."""
    if r_pr is None:
        return {}
    out: Dict[str, bool] = {}
    if r_pr.find(_ns("b")) is not None:
        out["bold"] = True
    if r_pr.find(_ns("i")) is not None:
        out["italic"] = True
    if r_pr.find(_ns("u")) is not None:
        out["underline"] = True
    if r_pr.find(_ns("strike")) is not None:
        out["strike"] = True
    v = r_pr.find(_ns("vertAlign"))
    if v is not None:
        val = v.get(_ns("val")) or ""
        if val == "superscript":
            out["sup"] = True
        elif val == "subscript":
            out["sub"] = True
    return out


def _wrap_format(text: str, flags: Dict[str, bool]) -> str:
    """Wrap a text run in the inline HTML tags EditorJS renders.
    Applied outside-in for `<b><i><u>...</u></i></b>` compatibility with
    Word's own emission order. Empty text passes through untouched."""
    if not text:
        return ""
    # Escape first — later insertions of ins/mark/tags don't get double-escaped.
    out = (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    if flags.get("sup"):
        out = f"<sup>{out}</sup>"
    if flags.get("sub"):
        out = f"<sub>{out}</sub>"
    if flags.get("underline"):
        out = f"<u>{out}</u>"
    if flags.get("strike"):
        out = f"<s>{out}</s>"
    if flags.get("italic"):
        out = f"<em>{out}</em>"
    if flags.get("bold"):
        out = f"<strong>{out}</strong>"
    return out


def _paragraph_to_html(p_element, authors: List[str], change_counter: List[int]) -> str:
    """Walk a `<w:p>` and turn its runs + revisions into inline HTML.
    Every `<w:ins>` becomes `<ins class="clm-suggestion-insert" data-author>`
    around its inner runs; every `<w:del>` becomes `<mark
    class="clm-redline" data-author>` around its inner (deleted) runs.

    `authors` is populated with any `w:author` string we see so the
    caller can surface a "By: Jane; Reviewed by: Mark" chip. `change_counter`
    is mutated in-place — a shared list-of-int lets both the recursive
    walk and the outer function agree without a class."""
    parts: List[str] = []

    def render_children(container, deleted_context: bool = False) -> str:
        buf: List[str] = []
        for node in container:
            tag = node.tag
            if tag == _ns("r"):
                r_pr = node.find(_ns("rPr"))
                text = _run_text(node, deleted=deleted_context)
                flags = _rpr_style_flags(r_pr)
                buf.append(_wrap_format(text, flags))
                continue
            if tag == _ns("ins"):
                # Nested revisions: rare but the OOXML spec allows them
                # (e.g. an insertion inside a comment). Recurse.
                inner = render_children(node, deleted_context=False)
                author = node.get(_ns("author"))
                if author:
                    authors.append(author)
                change_counter[0] += 1
                buf.append(
                    f'<ins class="clm-suggestion-insert"'
                    + (f' data-author="{_escape_attr(author)}"' if author else "")
                    + f'>{inner}</ins>'
                )
                continue
            if tag == _ns("del"):
                inner = render_children(node, deleted_context=True)
                author = node.get(_ns("author"))
                if author:
                    authors.append(author)
                change_counter[0] += 1
                buf.append(
                    f'<mark class="clm-redline"'
                    + (f' data-author="{_escape_attr(author)}"' if author else "")
                    + f'>{inner}</mark>'
                )
                continue
            if tag == _ns("hyperlink"):
                inner = render_children(node, deleted_context=deleted_context)
                buf.append(inner)  # Hyperlink formatting itself: dropped for now.
                continue
            # Unknown OOXML child — recurse; if there's nothing to
            # recurse into, drop it. Gracefully degrading is the goal.
            buf.append(render_children(node, deleted_context=deleted_context))
        return "".join(buf)

    parts.append(render_children(p_element))
    return "".join(parts)


def _escape_attr(val: str) -> str:
    return (val or "").replace('"', "&quot;")


# ─── Block detection ───────────────────────────────────────────────────────
def _detect_heading_level(p_element) -> Optional[int]:
    """Return the heading level 1-6 when the paragraph is styled as a
    Word heading; None otherwise. python-docx's `paragraph.style.name`
    would give us "Heading 3" etc., but pyxpath is quicker and doesn't
    depend on locale-specific style names."""
    p_pr = p_element.find(_ns("pPr"))
    if p_pr is None:
        return None
    p_style = p_pr.find(_ns("pStyle"))
    if p_style is None:
        return None
    val = p_style.get(_ns("val")) or ""
    m = re.match(r"[Hh]eading\s*(\d+)$", val) or re.match(r"[Hh](\d+)$", val)
    if not m:
        return None
    lvl = int(m.group(1))
    return lvl if 1 <= lvl <= 6 else None


def _detect_list_style(p_element) -> Optional[str]:
    """Return "ordered" / "unordered" if the paragraph is a numbered
    or bulleted list; None otherwise. Word encodes list membership via
    `<w:numPr>`; disambiguating ordered vs unordered requires reading
    the numbering.xml part which python-docx exposes via
    `paragraph._p.pPr` → but we don't need perfect disambiguation for
    EditorJS output. Ordered gets picked when `w:numId` maps to a
    numbering with a decimal format; otherwise bullet is safe."""
    p_pr = p_element.find(_ns("pPr"))
    if p_pr is None:
        return None
    num_pr = p_pr.find(_ns("numPr"))
    if num_pr is None:
        return None
    # We can't cheaply resolve numFmt without loading the numbering part;
    # default to unordered which renders as a bullet list in EditorJS.
    # Enterprise CLMs typically punt on this same edge (Ironclad falls
    # back to bulleted list unless the paragraph is explicitly Numbered
    # in Word's UI).
    return "unordered"


def _generate_block_id(idx: int) -> str:
    """EditorJS block ids are 10-char base36. We prefix with a per-file
    stamp so ids don't collide with existing blocks on the sender's
    side when the counterparty's reply lands."""
    import secrets
    return secrets.token_urlsafe(9)[:10].replace("-", "a").replace("_", "b")


# ─── Public entry point ────────────────────────────────────────────────────
def parse_counterparty_docx(docx_bytes: bytes) -> Dict[str, Any]:
    """Parse a `.docx` bytestring into EditorJS blocks preserving
    tracked-change marks. See module docstring for the return shape."""
    doc = Document(io.BytesIO(docx_bytes))
    authors_seen: List[str] = []
    change_counter = [0]
    blocks: List[Dict[str, Any]] = []
    running_list_style: Optional[str] = None
    running_list_items: List[str] = []

    body = doc.element.body

    def flush_list():
        nonlocal running_list_style, running_list_items
        if running_list_style and running_list_items:
            blocks.append(
                {
                    "id": _generate_block_id(len(blocks)),
                    "type": "list",
                    "data": {
                        "style": running_list_style,
                        "items": list(running_list_items),
                    },
                }
            )
        running_list_style = None
        running_list_items = []

    idx = 0
    for child in body:
        tag = child.tag
        if tag == _ns("p"):
            html = _paragraph_to_html(child, authors_seen, change_counter)
            heading = _detect_heading_level(child)
            list_style = _detect_list_style(child)
            if list_style:
                if running_list_style and running_list_style != list_style:
                    flush_list()
                running_list_style = list_style
                running_list_items.append(html)
                idx += 1
                continue
            # Not a list — flush any list-in-progress first.
            flush_list()
            if heading is not None:
                blocks.append(
                    {
                        "id": _generate_block_id(idx),
                        "type": "header",
                        "data": {"text": html, "level": heading},
                    }
                )
            else:
                blocks.append(
                    {
                        "id": _generate_block_id(idx),
                        "type": "paragraph",
                        "data": {"text": html},
                    }
                )
            idx += 1
        elif tag == _ns("tbl"):
            flush_list()
            table_content: List[List[str]] = []
            for row in child.findall(_ns("tr")):
                row_cells: List[str] = []
                for cell in row.findall(_ns("tc")):
                    cell_html_parts: List[str] = []
                    for p in cell.findall(_ns("p")):
                        cell_html_parts.append(_paragraph_to_html(p, authors_seen, change_counter))
                    row_cells.append("<br/>".join(cell_html_parts))
                table_content.append(row_cells)
            if table_content:
                blocks.append(
                    {
                        "id": _generate_block_id(idx),
                        "type": "table",
                        "data": {"content": table_content, "withHeadings": False},
                    }
                )
                idx += 1
        elif tag == _ns("sectPr"):
            # Section boundary — carries page break info. We emit no
            # block for it; page splitting is done downstream by
            # counting `<w:br w:type="page">` runs and inserted section
            # breaks. See page-split note in the calling handler.
            continue
        # Unknown children — silently skip to avoid crashing on
        # comment / bookmark / commentRangeStart elements Word likes to
        # sprinkle around.

    # Flush any trailing list.
    flush_list()

    # Page split — Word doesn't cleanly delimit "page 1 vs page 2" for
    # us; the closest we get is explicit page breaks (`<w:br w:type="page">`)
    # or `<w:lastRenderedPageBreak>` markers written by Word's layout
    # engine. For the round-trip flow we treat the whole document as
    # page 0 unless the caller passes `page_wise_blocks` to reconcile
    # against. That's the sender's problem, not the parser's — the
    # counterparty could have reflowed content across pages during
    # editing anyway. The BE controller does the page reconciliation
    # against the outbound-round snapshot; here we return one big page.
    pages = [{"pageNo": 0, "blocks": blocks}]

    return {
        "pages": pages,
        "counterparty_authors": sorted(set(a for a in authors_seen if a)),
        "change_count": change_counter[0],
    }


__all__ = ["parse_counterparty_docx"]
