"""Extract review comments from a DOCX file.

Parses comments.xml, commentsExtended.xml, and the document.xml to extract:
- Comment text, author, date, initials
- Anchor text (the text the comment is attached to)
- Parent comment ID (for threaded replies)
- Paragraph context around the anchor
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

from config import NAMESPACES

W = NAMESPACES["w"]
W14 = NAMESPACES["w14"]
W15 = NAMESPACES["w15"]


@dataclass
class ExtractedComment:
    """A comment extracted from the Early Rev document."""
    comment_id: int
    author: str
    date: str | None
    initials: str | None
    text: str
    # The text that was highlighted/anchored for this comment
    anchor_text: str
    # The full paragraph text containing the anchor
    anchor_paragraph_text: str
    # Paragraph index in document order where the anchor starts
    anchor_paragraph_index: int | None = None
    # Style of the anchor paragraph (e.g., "Heading1")
    anchor_paragraph_style: str | None = None
    # Whether the anchor is in a table
    anchor_in_table: bool = False
    # Table cell coordinates if applicable
    anchor_table_cell: tuple[int, int, int] | None = None
    # Parent comment ID for replies
    parent_id: int | None = None
    # Para ID from w14:paraId (for extended comment linking)
    para_id: str | None = None


def _get_text_recursive(element: etree._Element) -> str:
    """Get all text content from an element and its descendants."""
    texts = []
    for t in element.iter(f"{{{W}}}t"):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


def _get_comment_text(comment_elem: etree._Element) -> str:
    """Extract the text content from a <w:comment> element."""
    texts = []
    for p in comment_elem.iter(f"{{{W}}}p"):
        p_texts = []
        for r in p.iter(f"{{{W}}}r"):
            # Skip the annotationRef run
            ann_ref = r.find(f"{{{W}}}annotationRef")
            if ann_ref is not None:
                continue
            for t in r.iter(f"{{{W}}}t"):
                if t.text:
                    p_texts.append(t.text)
        if p_texts:
            texts.append("".join(p_texts))
    return "\n".join(texts)


def _extract_anchor_ranges(doc_root: etree._Element) -> dict[int, tuple[str, str, str | None]]:
    """Extract the anchor text for each comment from document.xml.

    Walks the XML to find commentRangeStart/commentRangeEnd pairs and
    collects the text between them.

    Returns:
        Dict mapping comment_id -> (anchor_text, paragraph_text, paragraph_style)
    """
    anchors: dict[int, tuple[str, str, str | None]] = {}

    # Build a map of comment range starts and ends
    range_starts: dict[int, etree._Element] = {}
    range_ends: dict[int, etree._Element] = {}

    for elem in doc_root.iter():
        tag = etree.QName(elem.tag).localname if isinstance(elem.tag, str) else ""
        if tag == "commentRangeStart":
            cid = elem.get(f"{{{W}}}id")
            if cid is not None:
                range_starts[int(cid)] = elem
        elif tag == "commentRangeEnd":
            cid = elem.get(f"{{{W}}}id")
            if cid is not None:
                range_ends[int(cid)] = elem

    # For each comment, collect text between start and end markers
    body = doc_root.find(f"{{{W}}}body")
    if body is None:
        return anchors

    # Flatten all elements in document order for range extraction
    all_paragraphs = list(body.iter(f"{{{W}}}p"))

    for cid, start_elem in range_starts.items():
        if cid not in range_ends:
            continue

        end_elem = range_ends[cid]

        # Find the paragraphs containing start and end
        start_para = _find_ancestor_paragraph(start_elem)
        end_para = _find_ancestor_paragraph(end_elem)

        if start_para is None:
            continue

        # Collect text between the range markers
        anchor_text = _collect_text_in_range(start_elem, end_elem, body)

        # Get full paragraph text and style
        para_text = _get_paragraph_full_text(start_para)
        para_style = _get_paragraph_style(start_para)

        anchors[cid] = (anchor_text, para_text, para_style)

    return anchors


def _find_ancestor_paragraph(elem: etree._Element) -> etree._Element | None:
    """Walk up the tree to find the containing <w:p> element."""
    current = elem.getparent()
    while current is not None:
        if etree.QName(current.tag).localname == "p":
            return current
        current = current.getparent()
    return None


def _get_paragraph_full_text(p_elem: etree._Element) -> str:
    """Get the full text of a paragraph."""
    texts = []
    for r in p_elem.iter(f"{{{W}}}r"):
        for t in r.iter(f"{{{W}}}t"):
            if t.text:
                texts.append(t.text)
    return "".join(texts)


def _get_paragraph_style(p_elem: etree._Element) -> str | None:
    """Get the style ID of a paragraph."""
    ppr = p_elem.find(f"{{{W}}}pPr")
    if ppr is not None:
        pstyle = ppr.find(f"{{{W}}}pStyle")
        if pstyle is not None:
            return pstyle.get(f"{{{W}}}val")
    return None


def _collect_text_in_range(
    start_elem: etree._Element,
    end_elem: etree._Element,
    body: etree._Element,
) -> str:
    """Collect all text content between commentRangeStart and commentRangeEnd.

    Uses a tree walk approach: iterate all elements in document order,
    collecting text from <w:t> elements between the two markers.
    """
    collecting = False
    texts = []

    for elem in body.iter():
        if elem is start_elem:
            collecting = True
            continue
        if elem is end_elem:
            break

        if collecting:
            tag = etree.QName(elem.tag).localname if isinstance(elem.tag, str) else ""
            if tag == "t" and elem.text:
                texts.append(elem.text)

    return "".join(texts)


def _get_para_ids(comments_extended_xml: bytes | None) -> dict[str, int | None]:
    """Parse commentsExtended.xml to get parent relationships.

    Returns:
        Dict mapping paraId -> parent comment's paraId (or None if top-level)
    """
    if not comments_extended_xml:
        return {}

    root = etree.fromstring(comments_extended_xml)
    para_map: dict[str, str | None] = {}

    for cex in root.iter(f"{{{W15}}}commentEx"):
        para_id = cex.get(f"{{{W15}}}paraId")
        parent_para_id = cex.get(f"{{{W15}}}paraIdParent")
        if para_id:
            para_map[para_id] = parent_para_id

    return para_map


def extract_comments(docx_path: str | Path) -> list[ExtractedComment]:
    """Extract all comments from a .docx file.

    Args:
        docx_path: Path to the .docx file.

    Returns:
        List of ExtractedComment objects.
    """
    docx_path = Path(docx_path)
    comments_list: list[ExtractedComment] = []

    with zipfile.ZipFile(docx_path, "r") as zf:
        names = zf.namelist()

        # Check if document has comments
        if "word/comments.xml" not in names:
            return comments_list

        # Parse comments.xml
        comments_xml = zf.read("word/comments.xml")
        comments_root = etree.fromstring(comments_xml)

        # Parse document.xml for anchor ranges
        doc_xml = zf.read("word/document.xml")
        doc_root = etree.fromstring(doc_xml)

        # Parse commentsExtended.xml for reply threading
        comments_ext_xml = None
        if "word/commentsExtended.xml" in names:
            comments_ext_xml = zf.read("word/commentsExtended.xml")

    # Extract anchor text for each comment
    anchor_map = _extract_anchor_ranges(doc_root)

    # Extract parent relationships from extended comments
    para_parent_map = _get_para_ids(comments_ext_xml)

    # Build paraId -> comment_id mapping
    para_to_comment: dict[str, int] = {}

    # Parse each comment
    for comment_elem in comments_root.iter(f"{{{W}}}comment"):
        cid_str = comment_elem.get(f"{{{W}}}id")
        if cid_str is None:
            continue

        cid = int(cid_str)
        author = comment_elem.get(f"{{{W}}}author", "Unknown")
        date = comment_elem.get(f"{{{W}}}date")
        initials = comment_elem.get(f"{{{W}}}initials")
        text = _get_comment_text(comment_elem)

        # Get para ID from first paragraph in comment
        para_id = None
        first_p = comment_elem.find(f"{{{W}}}p")
        if first_p is not None:
            para_id = first_p.get(f"{{{W14}}}paraId")
            if para_id:
                para_to_comment[para_id] = cid

        # Get anchor info
        anchor_text = ""
        anchor_para_text = ""
        anchor_style = None
        if cid in anchor_map:
            anchor_text, anchor_para_text, anchor_style = anchor_map[cid]

        comments_list.append(ExtractedComment(
            comment_id=cid,
            author=author,
            date=date,
            initials=initials,
            text=text,
            anchor_text=anchor_text,
            anchor_paragraph_text=anchor_para_text,
            anchor_paragraph_style=anchor_style,
            para_id=para_id,
        ))

    # Resolve parent comment IDs from paraId relationships
    comment_para_ids = {c.para_id: c.comment_id for c in comments_list if c.para_id}
    for comment in comments_list:
        if comment.para_id and comment.para_id in para_parent_map:
            parent_para_id = para_parent_map[comment.para_id]
            if parent_para_id and parent_para_id in comment_para_ids:
                comment.parent_id = comment_para_ids[parent_para_id]

    return comments_list
