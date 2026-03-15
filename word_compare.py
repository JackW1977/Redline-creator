"""Generate tracked changes between two Word documents using Word COM automation.

This module uses Microsoft Word's built-in Document.Compare() method via
win32com to produce the highest-fidelity tracked changes possible. Word's
native comparison engine handles:
- Character/word-level insertions and deletions
- Formatting changes
- Table structure changes
- Header/footer changes
- Move detection

Fallback: If COM automation is unavailable, provides a pure-Python Open XML
diff approach (lower fidelity but cross-platform).
"""

from __future__ import annotations

import logging
import shutil
import time
from pathlib import Path

from config import COM_TIMEOUT, COMPARE_GRANULARITY

logger = logging.getLogger(__name__)


def _is_com_available() -> bool:
    """Check if Word COM automation is available."""
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def compare_with_word_com(
    early_rev_path: str | Path,
    latest_rev_path: str | Path,
    output_path: str | Path,
    author: str = "Revision Compare",
    granularity: int = COMPARE_GRANULARITY,
) -> tuple[bool, str]:
    """Use Word COM to compare two documents and save the result with tracked changes.

    The output document will be based on the Latest Rev with tracked changes
    showing what changed from Early Rev.

    Args:
        early_rev_path: Path to the earlier revision.
        latest_rev_path: Path to the later revision.
        output_path: Where to save the comparison result.
        author: Author name for tracked changes.
        granularity: 0=character level, 1=word level.

    Returns:
        Tuple of (success, message).
    """
    import pythoncom
    import win32com.client

    early_rev_path = Path(early_rev_path).resolve()
    latest_rev_path = Path(latest_rev_path).resolve()
    output_path = Path(output_path).resolve()

    if not early_rev_path.exists():
        return False, f"Early revision not found: {early_rev_path}"
    if not latest_rev_path.exists():
        return False, f"Latest revision not found: {latest_rev_path}"

    word = None
    try:
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        # Open the EARLY revision as the "original" base document.
        # We then compare it against the LATEST revision.
        # This way Word's comparison engine treats:
        #   - Early Rev = original document
        #   - Latest Rev = revised document
        # Result:
        #   - Insertions (green/underline) = NEW content added in Latest Rev
        #   - Deletions (red/strikethrough) = content DELETED from Early Rev
        #   - Replacements = old text deleted + new text inserted (UPDATED)
        logger.info("Opening early revision as baseline in Word...")
        early_doc = word.Documents.Open(
            str(early_rev_path),
            ReadOnly=True,
            AddToRecentFiles=False,
        )

        logger.info("Running Word comparison engine (Early → Latest)...")
        # Document.Compare(Name):
        #   The document you call Compare on = original
        #   The Name parameter = revised (the document it is compared against)
        # CompareTarget=2 → wdCompareTargetNew (output to a new document)
        early_doc.Compare(
            Name=str(latest_rev_path),
            AuthorName=author,
            CompareTarget=2,  # wdCompareTargetNew - creates new comparison doc
            DetectFormatChanges=True,
            IgnoreAllComparisonWarnings=True,
            AddToRecentFiles=False,
        )

        active_doc = word.ActiveDocument

        logger.info(f"Saving comparison document to {output_path}...")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        active_doc.SaveAs2(
            str(output_path),
            FileFormat=12,  # wdFormatXMLDocument (.docx)
        )

        # Close documents
        active_doc.Close(SaveChanges=0)
        early_doc.Close(SaveChanges=0)

        logger.info("Word comparison complete.")
        return True, f"Comparison document saved to {output_path}"

    except Exception as e:
        logger.error(f"Word COM comparison failed: {e}")
        return False, f"Word COM comparison failed: {e}"

    finally:
        if word is not None:
            try:
                word.Quit(SaveChanges=0)
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def compare_documents(
    early_rev_path: str | Path,
    latest_rev_path: str | Path,
    output_path: str | Path,
    author: str = "Revision Compare",
    force_xml: bool = False,
) -> tuple[bool, str]:
    """Compare two documents, preferring COM automation when available.

    Args:
        early_rev_path: Path to the earlier revision.
        latest_rev_path: Path to the later revision.
        output_path: Where to save the output.
        author: Author name for tracked changes.
        force_xml: Force pure-XML approach even if COM is available.

    Returns:
        Tuple of (success, message).
    """
    if not force_xml and _is_com_available():
        logger.info("Using Word COM automation for document comparison.")
        return compare_with_word_com(
            early_rev_path, latest_rev_path, output_path, author
        )

    logger.info("Word COM not available. Using pure-XML comparison fallback.")
    return compare_with_xml(early_rev_path, latest_rev_path, output_path, author)


def compare_with_xml(
    early_rev_path: str | Path,
    latest_rev_path: str | Path,
    output_path: str | Path,
    author: str = "Revision Compare",
) -> tuple[bool, str]:
    """Pure-Python Open XML comparison fallback.

    Produces tracked changes showing:
      - Insertions  = content ADDED in Latest Rev (not in Early Rev)
      - Deletions   = content DELETED from Early Rev (not in Latest Rev)
      - Replacements = old text deleted + new text inserted (UPDATED)

    Limitations vs COM:
    - Paragraph-level granularity (not word/character level within paragraphs)
    - Cannot detect moves (only shows as delete + insert)
    - Cannot track formatting-only changes
    - May not handle complex table diffs well
    """
    import difflib
    import zipfile
    from copy import deepcopy
    from datetime import datetime, timezone

    from lxml import etree

    from text_extractor import extract_from_docx

    early_rev_path = Path(early_rev_path)
    latest_rev_path = Path(latest_rev_path)
    output_path = Path(output_path)

    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # Start with a copy of the latest revision (preserves styles/fonts/layout)
    shutil.copy2(latest_rev_path, output_path)

    # Extract paragraph text from both documents
    early_paras = extract_from_docx(early_rev_path)
    latest_paras = extract_from_docx(latest_rev_path)

    early_texts = [p.text for p in early_paras]
    latest_texts = [p.text for p in latest_paras]

    # Compute paragraph-level diff: early (original) → latest (revised)
    sm = difflib.SequenceMatcher(None, early_texts, latest_texts)
    opcodes = sm.get_opcodes()

    has_changes = any(op != "equal" for op, *_ in opcodes)
    if not has_changes:
        return True, "No differences found between documents."

    # Read the latest revision's document.xml
    with zipfile.ZipFile(output_path, "r") as zf:
        doc_xml = zf.read("word/document.xml")

    doc_root = etree.fromstring(doc_xml)
    body = doc_root.find(f"{{{W_NS}}}body")
    if body is None:
        return False, "Could not find document body in latest revision."

    # Get direct <w:p> children of body (skip paragraphs nested in tables etc.)
    body_paras = [child for child in body if child.tag == f"{{{W_NS}}}p"]

    ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    change_id = 1000  # Start high to avoid conflicts with existing IDs

    # Also read early doc XML so we can grab original paragraph formatting
    with zipfile.ZipFile(early_rev_path, "r") as zf:
        early_doc_xml = zf.read("word/document.xml")
    early_doc_root = etree.fromstring(early_doc_xml)
    early_body = early_doc_root.find(f"{{{W_NS}}}body")
    early_body_paras = [c for c in early_body if c.tag == f"{{{W_NS}}}p"] if early_body is not None else []

    def _next_id():
        nonlocal change_id
        cid = change_id
        change_id += 1
        return str(cid)

    def _make_ins_para(p_elem):
        """Wrap all runs in a paragraph with <w:ins> to mark as inserted (ADDED)."""
        # Collect all non-pPr children (runs, hyperlinks, etc.)
        children = [c for c in list(p_elem) if c.tag != f"{{{W_NS}}}pPr"]
        if not children:
            return

        # Create <w:ins> wrapper
        ins = etree.Element(f"{{{W_NS}}}ins")
        ins.set(f"{{{W_NS}}}id", _next_id())
        ins.set(f"{{{W_NS}}}author", author)
        ins.set(f"{{{W_NS}}}date", ts)

        # Move each child into the <w:ins> element
        for child in children:
            p_elem.remove(child)
            ins.append(child)

        # Append <w:ins> after pPr (or as first child)
        ppr = p_elem.find(f"{{{W_NS}}}pPr")
        if ppr is not None:
            ppr.addnext(ins)
        else:
            p_elem.insert(0, ins)

        # Also mark the paragraph property change (so the paragraph mark itself
        # is shown as an insertion in Word's revision view)
        ppr = p_elem.find(f"{{{W_NS}}}pPr")
        if ppr is None:
            ppr = etree.SubElement(p_elem, f"{{{W_NS}}}pPr")
            p_elem.insert(0, ppr)
        rpr = ppr.find(f"{{{W_NS}}}rPr")
        if rpr is None:
            rpr = etree.SubElement(ppr, f"{{{W_NS}}}rPr")
        rpr_ins = etree.SubElement(rpr, f"{{{W_NS}}}ins")
        rpr_ins.set(f"{{{W_NS}}}id", _next_id())
        rpr_ins.set(f"{{{W_NS}}}author", author)
        rpr_ins.set(f"{{{W_NS}}}date", ts)

    def _make_del_para(text, ref_early_para=None):
        """Create a new paragraph with <w:del> containing deleted text (DELETED from Early Rev)."""
        del_para = etree.Element(f"{{{W_NS}}}p")

        # Copy pPr from the original early paragraph if available (preserves style)
        if ref_early_para is not None:
            early_ppr = ref_early_para.find(f"{{{W_NS}}}pPr")
            if early_ppr is not None:
                del_para.append(deepcopy(early_ppr))

        # Mark paragraph property as deleted
        ppr = del_para.find(f"{{{W_NS}}}pPr")
        if ppr is None:
            ppr = etree.SubElement(del_para, f"{{{W_NS}}}pPr")
            del_para.insert(0, ppr)
        rpr = ppr.find(f"{{{W_NS}}}rPr")
        if rpr is None:
            rpr = etree.SubElement(ppr, f"{{{W_NS}}}rPr")
        del_rpr = etree.SubElement(rpr, f"{{{W_NS}}}del")
        del_rpr.set(f"{{{W_NS}}}id", _next_id())
        del_rpr.set(f"{{{W_NS}}}author", author)
        del_rpr.set(f"{{{W_NS}}}date", ts)

        # Add the deleted text inside <w:del> → <w:r> → <w:delText>
        del_wrapper = etree.SubElement(del_para, f"{{{W_NS}}}del")
        del_wrapper.set(f"{{{W_NS}}}id", _next_id())
        del_wrapper.set(f"{{{W_NS}}}author", author)
        del_wrapper.set(f"{{{W_NS}}}date", ts)

        run = etree.SubElement(del_wrapper, f"{{{W_NS}}}r")

        # Copy run properties from the early paragraph's first run if available
        if ref_early_para is not None:
            first_run = ref_early_para.find(f"{{{W_NS}}}r")
            if first_run is not None:
                early_rpr = first_run.find(f"{{{W_NS}}}rPr")
                if early_rpr is not None:
                    run.append(deepcopy(early_rpr))

        del_text = etree.SubElement(run, f"{{{W_NS}}}delText")
        del_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        del_text.text = text

        return del_para

    # Process opcodes in REVERSE order to avoid index shifts when inserting
    # deleted paragraphs into the body.
    insert_ops = []  # (position_in_body, element_to_insert_or_mark)

    # Track statistics
    stats = {"added": 0, "deleted": 0, "updated": 0}

    # We process backwards so inserting elements doesn't shift later indices
    for op, i1, i2, j1, j2 in reversed(opcodes):
        if op == "equal":
            continue

        if op == "insert":
            # Paragraphs ADDED in Latest Rev (exist in latest, not in early)
            # Mark them as insertions in the output
            for j in range(j1, j2):
                if j < len(body_paras):
                    _make_ins_para(body_paras[j])
                    stats["added"] += 1

        elif op == "delete":
            # Paragraphs DELETED from Early Rev (exist in early, not in latest)
            # Insert deleted paragraphs before position j1 in the body
            # Find the insertion point in the body
            if j1 < len(body_paras):
                anchor = body_paras[j1]
            else:
                # Append at end (before sectPr if present)
                sect_pr = body.find(f"{{{W_NS}}}sectPr")
                anchor = sect_pr  # may be None

            for idx in range(i2 - 1, i1 - 1, -1):
                early_p = early_body_paras[idx] if idx < len(early_body_paras) else None
                del_p = _make_del_para(early_texts[idx], ref_early_para=early_p)
                if anchor is not None:
                    anchor.addprevious(del_p)
                else:
                    body.append(del_p)
                stats["deleted"] += 1

        elif op == "replace":
            # Paragraphs UPDATED: old text deleted + new text inserted
            # 1) Mark latest paragraphs as insertions (new version)
            for j in range(j1, j2):
                if j < len(body_paras):
                    _make_ins_para(body_paras[j])

            # 2) Insert early paragraphs as deletions before the first
            #    inserted paragraph (old version)
            if j1 < len(body_paras):
                anchor = body_paras[j1]
            else:
                sect_pr = body.find(f"{{{W_NS}}}sectPr")
                anchor = sect_pr

            for idx in range(i2 - 1, i1 - 1, -1):
                early_p = early_body_paras[idx] if idx < len(early_body_paras) else None
                del_p = _make_del_para(early_texts[idx], ref_early_para=early_p)
                if anchor is not None:
                    anchor.addprevious(del_p)
                else:
                    body.append(del_p)

            stats["updated"] += max(i2 - i1, j2 - j1)

    # Write modified XML back to the docx
    modified_xml = etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8")

    with zipfile.ZipFile(output_path, "r") as zf_in:
        file_contents = {}
        for name in zf_in.namelist():
            if name == "word/document.xml":
                file_contents[name] = modified_xml
            else:
                file_contents[name] = zf_in.read(name)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
        for name, content in file_contents.items():
            zf_out.writestr(name, content)

    return True, (
        f"XML-based comparison complete: "
        f"{stats['added']} added, {stats['deleted']} deleted, "
        f"{stats['updated']} updated paragraph(s). "
        f"Note: Paragraph-level granularity only. "
        f"For word-level tracked changes, use Word COM."
    )
