"""Insert mapped comments into the output DOCX file.

This module takes the comparison document (which has tracked changes) and
injects comments from the Early Rev at their mapped locations. It works
directly on the Open XML inside the .docx zip archive.

Key operations:
1. Add comment definitions to word/comments.xml (and extended variants)
2. Add commentRangeStart/End markers to word/document.xml
3. Update relationships and content types
4. Handle unmapped comments by appending an "Unmapped Comments" section
"""

from __future__ import annotations

import logging
import random
import re
import zipfile
from copy import deepcopy
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

from lxml import etree

from comment_extractor import ExtractedComment
from comment_mapper import MappingResult, MappingStrategy
from config import NAMESPACES
from font_preserver import DocumentFonts, apply_fonts_to_rpr, extract_fonts
from text_extractor import StructuredParagraph, extract_from_xml

logger = logging.getLogger(__name__)

W = NAMESPACES["w"]
W14 = NAMESPACES["w14"]
W15 = NAMESPACES["w15"]
W16CID = NAMESPACES["w16cid"]
W16CEX = NAMESPACES["w16cex"]
R_NS = NAMESPACES["r"]

# Relationship types
REL_COMMENTS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
REL_COMMENTS_EXT = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)
REL_COMMENTS_IDS = (
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
)
REL_COMMENTS_EXTENSIBLE = (
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
)

# Content types
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
CT_COMMENTS_EXT = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"


def _gen_hex_id() -> str:
    return f"{random.randint(0, 0x7FFFFFFE):08X}"


def _get_max_id(doc_root: etree._Element) -> int:
    """Find the maximum w:id value used in the document for tracked changes, bookmarks, etc."""
    max_id = 0
    for elem in doc_root.iter():
        id_val = elem.get(f"{{{W}}}id")
        if id_val is not None:
            try:
                max_id = max(max_id, int(id_val))
            except ValueError:
                pass
    return max_id


def _get_max_rid(rels_root: etree._Element) -> int:
    """Find the max relationship ID number."""
    max_rid = 0
    for rel in rels_root.iter("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
        rid = rel.get("Id", "")
        if rid.startswith("rId"):
            try:
                max_rid = max(max_rid, int(rid[3:]))
            except ValueError:
                pass
    return max_rid


def _has_relationship(rels_root: etree._Element, target: str) -> bool:
    """Check if a relationship to the given target already exists."""
    for rel in rels_root.iter("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
        if rel.get("Target") == target:
            return True
    return False


def _ensure_comments_infrastructure(zip_contents: dict[str, bytes]) -> dict[str, bytes]:
    """Ensure the docx has all necessary comment-related parts, rels, and content types."""
    RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

    # --- Relationships ---
    rels_path = "word/_rels/document.xml.rels"
    if rels_path in zip_contents:
        rels_root = etree.fromstring(zip_contents[rels_path])
    else:
        rels_root = etree.Element(f"{{{RELS_NS}}}Relationships")

    next_rid = _get_max_rid(rels_root) + 1
    rels_to_add = [
        (REL_COMMENTS, "comments.xml"),
        (REL_COMMENTS_EXT, "commentsExtended.xml"),
        (REL_COMMENTS_IDS, "commentsIds.xml"),
        (REL_COMMENTS_EXTENSIBLE, "commentsExtensible.xml"),
    ]

    for rel_type, target in rels_to_add:
        if not _has_relationship(rels_root, target):
            rel = etree.SubElement(rels_root, f"{{{RELS_NS}}}Relationship")
            rel.set("Id", f"rId{next_rid}")
            rel.set("Type", rel_type)
            rel.set("Target", target)
            next_rid += 1

    zip_contents[rels_path] = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8")

    # --- Content Types ---
    ct_path = "[Content_Types].xml"
    if ct_path in zip_contents:
        ct_root = etree.fromstring(zip_contents[ct_path])
    else:
        return zip_contents  # Can't proceed without content types

    overrides_to_add = [
        ("/word/comments.xml", CT_COMMENTS),
        ("/word/commentsExtended.xml", CT_COMMENTS_EXT),
    ]

    existing_parts = {
        override.get("PartName")
        for override in ct_root.iter(f"{{{CT_NS}}}Override")
    }

    for part_name, content_type in overrides_to_add:
        if part_name not in existing_parts:
            override = etree.SubElement(ct_root, f"{{{CT_NS}}}Override")
            override.set("PartName", part_name)
            override.set("ContentType", content_type)

    zip_contents[ct_path] = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8")

    # --- Empty comment files if they don't exist ---
    if "word/comments.xml" not in zip_contents:
        comments_root = etree.Element(
            f"{{{W}}}comments",
            nsmap={
                "w": W,
                "w14": W14,
                "r": R_NS,
                "mc": NAMESPACES["mc"],
            },
        )
        zip_contents["word/comments.xml"] = etree.tostring(
            comments_root, xml_declaration=True, encoding="UTF-8"
        )

    if "word/commentsExtended.xml" not in zip_contents:
        ext_root = etree.Element(
            f"{{{W15}}}commentsEx",
            nsmap={"w15": W15},
        )
        zip_contents["word/commentsExtended.xml"] = etree.tostring(
            ext_root, xml_declaration=True, encoding="UTF-8"
        )

    return zip_contents


def _build_comment_element(
    comment: ExtractedComment,
    new_id: int,
    para_id: str,
) -> etree._Element:
    """Build a <w:comment> XML element from an ExtractedComment."""
    nsmap = {"w": W, "w14": W14, "r": R_NS}

    comment_elem = etree.Element(f"{{{W}}}comment", nsmap=nsmap)
    comment_elem.set(f"{{{W}}}id", str(new_id))
    comment_elem.set(f"{{{W}}}author", comment.author)
    if comment.date:
        comment_elem.set(f"{{{W}}}date", comment.date)
    if comment.initials:
        comment_elem.set(f"{{{W}}}initials", comment.initials)

    # Build comment content paragraphs
    for i, line in enumerate(comment.text.split("\n")):
        p = etree.SubElement(comment_elem, f"{{{W}}}p")
        p.set(f"{{{W14}}}paraId", para_id if i == 0 else _gen_hex_id())
        p.set(f"{{{W14}}}textId", "77777777")

        if i == 0:
            # First paragraph gets the annotation reference
            ref_run = etree.SubElement(p, f"{{{W}}}r")
            ref_rpr = etree.SubElement(ref_run, f"{{{W}}}rPr")
            ref_style = etree.SubElement(ref_rpr, f"{{{W}}}rStyle")
            ref_style.set(f"{{{W}}}val", "CommentReference")
            etree.SubElement(ref_run, f"{{{W}}}annotationRef")

        # Text run
        text_run = etree.SubElement(p, f"{{{W}}}r")
        text_rpr = etree.SubElement(text_run, f"{{{W}}}rPr")
        color = etree.SubElement(text_rpr, f"{{{W}}}color")
        color.set(f"{{{W}}}val", "000000")
        sz = etree.SubElement(text_rpr, f"{{{W}}}sz")
        sz.set(f"{{{W}}}val", "20")
        szcs = etree.SubElement(text_rpr, f"{{{W}}}szCs")
        szcs.set(f"{{{W}}}val", "20")

        t = etree.SubElement(text_run, f"{{{W}}}t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = line

    return comment_elem


def _find_paragraph_in_doc(
    doc_root: etree._Element,
    target: StructuredParagraph,
) -> etree._Element | None:
    """Find the <w:p> element in doc_root that matches the target paragraph.

    Uses text content matching since XML element references from extract_from_docx
    won't be valid against the comparison document's DOM.
    """
    body = doc_root.find(f"{{{W}}}body")
    if body is None:
        return None

    target_text = target.text.strip()
    if not target_text:
        # For empty paragraphs, match by index
        para_idx = 0
        for p in body.iter(f"{{{W}}}p"):
            if para_idx == target.index:
                return p
            para_idx += 1
        return None

    # Try matching against both "all text" and "accepted text" (no deletions)
    # The comparison document has tracked changes, so the "accepted" text
    # (what you'd see after accepting changes) should match the latest rev text.
    norm_target = re.sub(r"\s+", " ", target_text).strip().lower()

    # Pass 1: exact match against accepted text (most reliable for comparison docs)
    for p in body.iter(f"{{{W}}}p"):
        p_text_accepted = re.sub(r"\s+", " ", _get_para_text(p, include_deleted=False)).strip().lower()
        if p_text_accepted == norm_target:
            return p

    # Pass 2: exact match against all text (for docs without tracked changes)
    for p in body.iter(f"{{{W}}}p"):
        p_text_all = re.sub(r"\s+", " ", _get_para_text(p, include_deleted=True)).strip().lower()
        if p_text_all == norm_target:
            return p

    # Pass 3: containment / fuzzy match against accepted text
    import difflib
    best_match = None
    best_score = 0.0

    for p in body.iter(f"{{{W}}}p"):
        p_text = re.sub(r"\s+", " ", _get_para_text(p, include_deleted=False)).strip().lower()
        if not p_text:
            continue
        # Containment
        if norm_target in p_text or p_text in norm_target:
            score = min(len(norm_target), len(p_text)) / max(len(norm_target), len(p_text))
            if score > best_score:
                best_score = score
                best_match = p
        else:
            # SequenceMatcher ratio
            ratio = difflib.SequenceMatcher(None, norm_target, p_text).ratio()
            if ratio > best_score:
                best_score = ratio
                best_match = p

    if best_match is not None and best_score > 0.6:
        return best_match

    return None


def _get_para_text(p_elem: etree._Element, include_deleted: bool = True) -> str:
    """Get plain text from a <w:p>.

    Args:
        p_elem: The paragraph XML element.
        include_deleted: If True, include text from <w:delText> elements.
            If False, only include "current" text (what you'd see after accepting
            all changes) — i.e., <w:t> from regular runs and <w:ins> runs,
            but NOT <w:delText> from <w:del> runs.
    """
    if include_deleted:
        texts = []
        for elem in p_elem.iter():
            tag = etree.QName(elem.tag).localname if isinstance(elem.tag, str) else ""
            if tag in ("t", "delText") and elem.text:
                texts.append(elem.text)
        return "".join(texts)
    else:
        # Only collect <w:t> text that is NOT inside a <w:del>
        texts = []
        _collect_non_deleted_text(p_elem, texts, in_del=False)
        return "".join(texts)


def _collect_non_deleted_text(
    elem: etree._Element, texts: list[str], in_del: bool
) -> None:
    """Recursively collect text from <w:t> elements, skipping <w:del> contents."""
    tag = etree.QName(elem.tag).localname if isinstance(elem.tag, str) else ""

    if tag == "del":
        # Skip everything inside a deletion
        return

    if tag == "t" and elem.text:
        texts.append(elem.text)

    for child in elem:
        _collect_non_deleted_text(child, texts, in_del)


def _insert_comment_markers(
    p_elem: etree._Element,
    comment_id: int,
    anchor_text: str | None,
    char_offset: int | None,
    anchor_length: int | None,
) -> bool:
    """Insert commentRangeStart/End markers into a paragraph.

    If anchor_text is provided, tries to wrap exactly that text.
    Otherwise, wraps the entire paragraph content.

    Returns True if markers were successfully inserted.
    """
    # Create marker elements
    range_start = etree.Element(f"{{{W}}}commentRangeStart")
    range_start.set(f"{{{W}}}id", str(comment_id))

    range_end = etree.Element(f"{{{W}}}commentRangeEnd")
    range_end.set(f"{{{W}}}id", str(comment_id))

    ref_run = etree.Element(f"{{{W}}}r")
    ref_rpr = etree.SubElement(ref_run, f"{{{W}}}rPr")
    ref_style = etree.SubElement(ref_rpr, f"{{{W}}}rStyle")
    ref_style.set(f"{{{W}}}val", "CommentReference")
    ref_elem = etree.SubElement(ref_run, f"{{{W}}}commentReference")
    ref_elem.set(f"{{{W}}}id", str(comment_id))

    # Get direct children that are runs, ins, del, etc.
    children = list(p_elem)

    if not children:
        # Empty paragraph - just add markers
        p_elem.append(range_start)
        p_elem.append(range_end)
        p_elem.append(ref_run)
        return True

    if anchor_text and char_offset is not None:
        # Try to find the specific anchor location
        # For now, use a simplified approach: find the run containing the offset
        current_offset = 0
        start_inserted = False
        end_inserted = False
        end_offset = char_offset + (anchor_length or len(anchor_text))

        for i, child in enumerate(children):
            tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""

            # Calculate text length of this element
            elem_text = ""
            if tag == "r":
                for t in child.iter(f"{{{W}}}t"):
                    if t.text:
                        elem_text += t.text
            elif tag in ("ins", "del"):
                for t in child.iter(f"{{{W}}}t", f"{{{W}}}delText"):
                    if t.text:
                        elem_text += t.text

            elem_len = len(elem_text)

            if not start_inserted and current_offset + elem_len > char_offset:
                # Insert range start before this element
                child.addprevious(range_start)
                start_inserted = True

            if start_inserted and not end_inserted and current_offset + elem_len >= end_offset:
                # Insert range end after this element
                child.addnext(range_end)
                end_inserted = True
                break

            current_offset += elem_len

        if start_inserted and not end_inserted:
            # End marker at end of paragraph
            p_elem.append(range_end)
            end_inserted = True

        if not start_inserted:
            # Fallback: wrap entire paragraph
            if len(children) > 0:
                first_content = None
                for child in children:
                    tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
                    if tag != "pPr":
                        first_content = child
                        break
                if first_content is not None:
                    first_content.addprevious(range_start)
                else:
                    p_elem.append(range_start)
            p_elem.append(range_end)

        # Always add the comment reference at the end
        p_elem.append(ref_run)
        return True

    else:
        # Wrap entire paragraph content
        first_content = None
        for child in children:
            tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
            if tag != "pPr":
                first_content = child
                break

        if first_content is not None:
            first_content.addprevious(range_start)
        else:
            p_elem.append(range_start)

        p_elem.append(range_end)
        p_elem.append(ref_run)
        return True


def _add_unmapped_comments_section(
    doc_root: etree._Element,
    unmapped: list[MappingResult],
    fonts: DocumentFonts | None = None,
) -> None:
    """Add an 'Unmapped Comments' section at the end of the document body.

    Uses fonts from the Latest Rev so the appendix matches the document's typography.
    """
    body = doc_root.find(f"{{{W}}}body")
    if body is None:
        return

    # Find the sectPr (section properties) - usually the last child
    sect_pr = body.find(f"{{{W}}}sectPr")

    def _styled_rpr(parent_run, *, bold=False, italic=False, color=None, size_hp=None):
        """Helper: build a <w:rPr> with font settings from the Latest Rev."""
        rpr = etree.SubElement(parent_run, f"{{{W}}}rPr")
        if fonts:
            apply_fonts_to_rpr(rpr, fonts, is_heading=False)
        if bold:
            etree.SubElement(rpr, f"{{{W}}}b")
        if italic:
            etree.SubElement(rpr, f"{{{W}}}i")
        if color:
            c = etree.SubElement(rpr, f"{{{W}}}color")
            c.set(f"{{{W}}}val", color)
        if size_hp:
            sz = rpr.find(f"{{{W}}}sz")
            if sz is None:
                sz = etree.SubElement(rpr, f"{{{W}}}sz")
            sz.set(f"{{{W}}}val", size_hp)
            szcs = rpr.find(f"{{{W}}}szCs")
            if szcs is None:
                szcs = etree.SubElement(rpr, f"{{{W}}}szCs")
            szcs.set(f"{{{W}}}val", size_hp)
        return rpr

    # Add a page break before the appendix
    break_para = etree.SubElement(body, f"{{{W}}}p")
    break_run = etree.SubElement(break_para, f"{{{W}}}r")
    br = etree.SubElement(break_run, f"{{{W}}}br")
    br.set(f"{{{W}}}type", "page")

    # Add heading (uses the document's Heading1 style, which already has the right font)
    heading_para = etree.SubElement(body, f"{{{W}}}p")
    heading_ppr = etree.SubElement(heading_para, f"{{{W}}}pPr")
    heading_style = etree.SubElement(heading_ppr, f"{{{W}}}pStyle")
    heading_style.set(f"{{{W}}}val", "Heading1")
    heading_run = etree.SubElement(heading_para, f"{{{W}}}r")
    if fonts:
        heading_rpr = etree.SubElement(heading_run, f"{{{W}}}rPr")
        apply_fonts_to_rpr(heading_rpr, fonts, is_heading=True)
    heading_t = etree.SubElement(heading_run, f"{{{W}}}t")
    heading_t.text = "Unmapped Comments from Early Revision"

    # Add description
    desc_para = etree.SubElement(body, f"{{{W}}}p")
    desc_run = etree.SubElement(desc_para, f"{{{W}}}r")
    _styled_rpr(desc_run, italic=True, size_hp="20")
    desc_t = etree.SubElement(desc_run, f"{{{W}}}t")
    desc_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    desc_t.text = (
        "The following comments from the early revision could not be reliably "
        "mapped to locations in the latest revision."
    )

    # Add each unmapped comment
    for mapping in unmapped:
        comment = mapping.comment

        # Separator line
        sep_para = etree.SubElement(body, f"{{{W}}}p")
        sep_ppr = etree.SubElement(sep_para, f"{{{W}}}pPr")
        sep_border = etree.SubElement(sep_ppr, f"{{{W}}}pBdr")
        sep_bottom = etree.SubElement(sep_border, f"{{{W}}}bottom")
        sep_bottom.set(f"{{{W}}}val", "single")
        sep_bottom.set(f"{{{W}}}sz", "4")
        sep_bottom.set(f"{{{W}}}space", "1")
        sep_bottom.set(f"{{{W}}}color", "auto")

        # Author and date
        meta_para = etree.SubElement(body, f"{{{W}}}p")
        meta_run = etree.SubElement(meta_para, f"{{{W}}}r")
        _styled_rpr(meta_run, bold=True, size_hp="20")
        meta_t = etree.SubElement(meta_run, f"{{{W}}}t")
        meta_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        date_str = f" ({comment.date})" if comment.date else ""
        meta_t.text = f"Comment by {comment.author}{date_str}:"

        # Comment text
        text_para = etree.SubElement(body, f"{{{W}}}p")
        text_run = etree.SubElement(text_para, f"{{{W}}}r")
        _styled_rpr(text_run, size_hp="20")
        text_t = etree.SubElement(text_run, f"{{{W}}}t")
        text_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text_t.text = comment.text

        # Original anchor info
        if comment.anchor_text:
            anchor_para = etree.SubElement(body, f"{{{W}}}p")
            anchor_run = etree.SubElement(anchor_para, f"{{{W}}}r")
            _styled_rpr(anchor_run, italic=True, color="808080", size_hp="18")
            anchor_t = etree.SubElement(anchor_run, f"{{{W}}}t")
            anchor_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            anchor_t.text = f'Original anchor text: "{comment.anchor_text[:200]}"'

        # Mapping note
        if mapping.note:
            note_para = etree.SubElement(body, f"{{{W}}}p")
            note_run = etree.SubElement(note_para, f"{{{W}}}r")
            _styled_rpr(note_run, italic=True, color="808080", size_hp="18")
            note_t = etree.SubElement(note_run, f"{{{W}}}t")
            note_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            note_t.text = f"Reason: {mapping.note[:300]}"

    # Move sectPr back to end if it exists
    if sect_pr is not None:
        body.remove(sect_pr)
        body.append(sect_pr)


def insert_comments(
    docx_path: str | Path,
    output_path: str | Path,
    mapping_results: list[MappingResult],
    latest_rev_path: str | Path | None = None,
) -> tuple[bool, str, list[str]]:
    """Insert mapped comments into the output DOCX file.

    Args:
        docx_path: Path to the comparison document (with tracked changes).
        output_path: Where to save the final output.
        mapping_results: Comment mapping results from comment_mapper.
        latest_rev_path: Path to the Latest Rev for font extraction.

    Returns:
        Tuple of (success, message, list of log entries).
    """
    docx_path = Path(docx_path)
    output_path = Path(output_path)
    log_entries: list[str] = []

    # Read all contents from the docx
    zip_contents: dict[str, bytes] = {}
    with zipfile.ZipFile(docx_path, "r") as zf:
        for name in zf.namelist():
            zip_contents[name] = zf.read(name)

    # Extract fonts from Latest Rev for style-consistent appendix text
    fonts: DocumentFonts | None = None
    if latest_rev_path:
        try:
            fonts = extract_fonts(latest_rev_path)
        except Exception as e:
            logger.warning(f"Could not extract fonts from Latest Rev: {e}")

    # Ensure comment infrastructure exists
    zip_contents = _ensure_comments_infrastructure(zip_contents)

    # Parse the XML files we need to modify
    doc_root = etree.fromstring(zip_contents["word/document.xml"])
    comments_root = etree.fromstring(zip_contents["word/comments.xml"])
    comments_ext_root = etree.fromstring(zip_contents["word/commentsExtended.xml"])

    # Find the next available comment ID
    next_id = _get_max_id(doc_root) + 1

    # Build a set of existing comments (author, normalized text) for deduplication.
    # Word COM's Compare() may have already carried forward comments from the early rev.
    existing_comments: set[tuple[str, str]] = set()
    for c in comments_root.iter(f"{{{W}}}comment"):
        c_author = c.get(f"{{{W}}}author", "")
        c_texts = []
        for t_elem in c.iter(f"{{{W}}}t"):
            if t_elem.text:
                c_texts.append(t_elem.text)
        c_text = re.sub(r"\s+", " ", " ".join(c_texts)).strip().lower()
        existing_comments.add((c_author.lower(), c_text))

    # Separate mapped vs unmapped comments
    mapped = [r for r in mapping_results if r.strategy != MappingStrategy.UNMAPPED]
    unmapped = [r for r in mapping_results if r.strategy == MappingStrategy.UNMAPPED]

    # Track comment ID remapping (old_id -> new_id) for reply threading
    id_remap: dict[int, int] = {}
    para_id_map: dict[int, str] = {}  # new_id -> para_id
    skipped_count = 0

    # Process mapped comments
    for result in mapped:
        comment = result.comment

        # Deduplication: skip if this comment already exists in the comparison doc
        norm_text = re.sub(r"\s+", " ", comment.text).strip().lower()
        key = (comment.author.lower(), norm_text)
        if key in existing_comments:
            log_entries.append(
                f"SKIPPED comment {comment.comment_id} by {comment.author} "
                f"(already present from Word comparison)."
            )
            skipped_count += 1
            continue

        new_id = next_id
        next_id += 1
        id_remap[comment.comment_id] = new_id
        para_id = _gen_hex_id()
        para_id_map[new_id] = para_id

        # Find the target paragraph in the document XML
        target_para = _find_paragraph_in_doc(doc_root, result.target_paragraph)
        if target_para is None:
            log_entries.append(
                f"WARNING: Could not find target paragraph for comment {comment.comment_id} "
                f"by {comment.author}. Moving to unmapped."
            )
            unmapped.append(result)
            continue

        # Add comment definition to comments.xml
        comment_elem = _build_comment_element(comment, new_id, para_id)
        comments_root.append(comment_elem)

        # Add to commentsExtended.xml
        cex = etree.SubElement(comments_ext_root, f"{{{W15}}}commentEx")
        cex.set(f"{{{W15}}}paraId", para_id)
        cex.set(f"{{{W15}}}done", "0")

        # Insert comment range markers into the document
        success = _insert_comment_markers(
            target_para,
            new_id,
            result.target_anchor_text,
            result.target_char_offset,
            result.target_anchor_length,
        )

        if success:
            log_entries.append(
                f"Inserted comment {comment.comment_id} -> {new_id} "
                f"by {comment.author} [{result.strategy.value}] "
                f"(confidence={result.confidence:.2f})"
            )
        else:
            log_entries.append(
                f"WARNING: Failed to insert markers for comment {comment.comment_id}"
            )

    # Process reply threading
    for result in mapped:
        comment = result.comment
        if comment.parent_id is not None and comment.parent_id in id_remap:
            new_id = id_remap.get(comment.comment_id)
            parent_new_id = id_remap[comment.parent_id]
            if new_id and new_id in para_id_map and parent_new_id in para_id_map:
                # Update the commentsExtended entry with parent reference
                for cex in comments_ext_root.iter(f"{{{W15}}}commentEx"):
                    if cex.get(f"{{{W15}}}paraId") == para_id_map[new_id]:
                        cex.set(f"{{{W15}}}paraIdParent", para_id_map[parent_new_id])
                        break

    # Handle unmapped comments
    if unmapped:
        _add_unmapped_comments_section(doc_root, unmapped, fonts=fonts)
        log_entries.append(
            f"Added {len(unmapped)} unmapped comment(s) to appendix section."
        )

    # Serialize modified XML back
    zip_contents["word/document.xml"] = etree.tostring(
        doc_root, xml_declaration=True, encoding="UTF-8"
    )
    zip_contents["word/comments.xml"] = etree.tostring(
        comments_root, xml_declaration=True, encoding="UTF-8"
    )
    zip_contents["word/commentsExtended.xml"] = etree.tostring(
        comments_ext_root, xml_declaration=True, encoding="UTF-8"
    )

    # Write output docx
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in zip_contents.items():
            zf.writestr(name, content)

    inserted_count = len(mapped) - len([r for r in unmapped if r in mapped]) - skipped_count
    total = len(mapping_results)

    parts = [f"Output saved to {output_path}"]
    if skipped_count:
        parts.insert(0, f"{skipped_count} already present from Word comparison (skipped)")
    if inserted_count:
        parts.insert(0, f"{inserted_count} newly inserted")
    if unmapped:
        parts.insert(0, f"{len(unmapped)} unmapped (appended to document)")
    parts.insert(0, f"{total} comments total")

    message = ". ".join(parts) + "."

    return True, message, log_entries
