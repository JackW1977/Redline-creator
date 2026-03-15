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

        logger.info("Opening latest revision in Word...")
        latest_doc = word.Documents.Open(
            str(latest_rev_path),
            ReadOnly=True,
            AddToRecentFiles=False,
        )

        logger.info("Running Word comparison engine...")
        # Document.Compare parameters:
        # Name, AuthorName, CompareTarget, DetectFormatChanges,
        # IgnoreAllComparisonWarnings, AddToRecentFiles, RemovePersonalInformation,
        # RemoveDateAndTime
        #
        # CompareTarget: 0=wdCompareTargetSelected, 1=wdCompareTargetCurrent, 2=wdCompareTargetNew
        # We use wdCompareTargetNew (2) to create a new document
        compared_doc = latest_doc.Compare(
            Name=str(early_rev_path),
            AuthorName=author,
            CompareTarget=2,  # wdCompareTargetNew - creates new comparison doc
            DetectFormatChanges=True,
            IgnoreAllComparisonWarnings=True,
            AddToRecentFiles=False,
        )

        # The comparison creates a new document that's now the active doc
        # It shows Early Rev as the base with changes to reach Latest Rev
        # We need to reverse this: Latest Rev as base with changes FROM Early Rev
        #
        # Word.Compare(A, B) produces: "B was the original, here's what changed to get A"
        # Since we opened Latest and compared against Early:
        # Result = "Early was original, changes show how to get to Latest"
        # This is exactly what we want: Latest Rev as the visual base,
        # tracked changes showing what changed from Early Rev.

        active_doc = word.ActiveDocument

        logger.info(f"Saving comparison document to {output_path}...")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        active_doc.SaveAs2(
            str(output_path),
            FileFormat=12,  # wdFormatXMLDocument (.docx)
        )

        # Close documents
        active_doc.Close(SaveChanges=0)
        latest_doc.Close(SaveChanges=0)

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

    This produces lower-fidelity tracked changes by:
    1. Starting from the Latest Rev as the base document
    2. Extracting paragraph-level text from both documents
    3. Using difflib to compute paragraph-level diffs
    4. Injecting <w:ins> and <w:del> markup into the output XML

    Limitations vs COM:
    - Only paragraph-level granularity (not word/character level within paragraphs)
    - Cannot detect moves (only shows as delete + insert)
    - Cannot track formatting-only changes
    - May not handle complex table diffs well
    """
    import difflib
    import zipfile
    from datetime import datetime, timezone

    from lxml import etree

    from text_extractor import extract_from_docx

    early_rev_path = Path(early_rev_path)
    latest_rev_path = Path(latest_rev_path)
    output_path = Path(output_path)

    # Start with a copy of the latest revision
    shutil.copy2(latest_rev_path, output_path)

    # Extract text from both
    early_paras = extract_from_docx(early_rev_path)
    latest_paras = extract_from_docx(latest_rev_path)

    early_texts = [p.text for p in early_paras]
    latest_texts = [p.text for p in latest_paras]

    # Compute paragraph-level diff
    sm = difflib.SequenceMatcher(None, early_texts, latest_texts)
    opcodes = sm.get_opcodes()

    # If no changes, just return the copy
    has_changes = any(op != "equal" for op, *_ in opcodes)
    if not has_changes:
        return True, "No differences found between documents."

    # Read the latest revision's document.xml
    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    import tempfile

    with zipfile.ZipFile(output_path, "r") as zf:
        doc_xml = zf.read("word/document.xml")

    doc_root = etree.fromstring(doc_xml)
    body = doc_root.find(f"{{{W_NS}}}body")

    if body is None:
        return False, "Could not find document body in latest revision."

    # Get all <w:p> elements in document order (excluding those in tables for now)
    all_paras = list(body.iter(f"{{{W_NS}}}p"))

    ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    change_id = 1000  # Start high to avoid conflicts

    # Process opcodes to inject tracked changes
    # We work backwards through the latest doc to avoid index shifting
    modifications = []  # (latest_index, operation, early_text)

    for op, i1, i2, j1, j2 in opcodes:
        if op == "equal":
            continue
        elif op == "delete":
            # Text in early but not in latest -> show as deletion before j1
            for idx in range(i1, i2):
                modifications.append(("delete", j1, early_texts[idx]))
        elif op == "insert":
            # Text in latest but not in early -> mark as insertion
            for idx in range(j1, j2):
                modifications.append(("insert", idx, latest_texts[idx]))
        elif op == "replace":
            # Changed paragraphs -> deletion of old + insertion of new
            for idx in range(i1, i2):
                modifications.append(("delete", j1, early_texts[idx]))
            for idx in range(j1, j2):
                modifications.append(("insert", idx, latest_texts[idx]))

    # Apply tracked change markup to the document
    # Mark insertions
    for mod_type, para_idx, text in modifications:
        if para_idx >= len(all_paras):
            continue

        target_para = all_paras[para_idx]

        if mod_type == "insert":
            # Wrap all runs in this paragraph with <w:ins>
            runs = list(target_para.iter(f"{{{W_NS}}}r"))
            if runs:
                ins_elem = etree.SubElement(target_para, f"{{{W_NS}}}ins")
                ins_elem.set(f"{{{W_NS}}}id", str(change_id))
                ins_elem.set(f"{{{W_NS}}}author", author)
                ins_elem.set(f"{{{W_NS}}}date", ts)
                change_id += 1

                # Move runs into the ins element
                for run in runs:
                    # We need to be careful: just add the ins wrapper attribute
                    # For simplicity in the XML fallback, we'll use a simpler approach
                    pass

            # Simplified approach: add rPr with ins marker
            # This is a limitation of the pure-XML approach
            ppr = target_para.find(f"{{{W_NS}}}pPr")
            if ppr is None:
                ppr = etree.SubElement(target_para, f"{{{W_NS}}}pPr")
                target_para.insert(0, ppr)
            rpr = ppr.find(f"{{{W_NS}}}rPr")
            if rpr is None:
                rpr = etree.SubElement(ppr, f"{{{W_NS}}}rPr")
            ins = etree.SubElement(rpr, f"{{{W_NS}}}ins")
            ins.set(f"{{{W_NS}}}id", str(change_id))
            ins.set(f"{{{W_NS}}}author", author)
            ins.set(f"{{{W_NS}}}date", ts)
            change_id += 1

        elif mod_type == "delete":
            # Insert a deleted paragraph before the target
            del_para = etree.Element(f"{{{W_NS}}}p")
            ppr = etree.SubElement(del_para, f"{{{W_NS}}}pPr")
            rpr = etree.SubElement(ppr, f"{{{W_NS}}}rPr")
            del_mark = etree.SubElement(rpr, f"{{{W_NS}}}del")
            del_mark.set(f"{{{W_NS}}}id", str(change_id))
            del_mark.set(f"{{{W_NS}}}author", author)
            del_mark.set(f"{{{W_NS}}}date", ts)
            change_id += 1

            # Add the deleted text as a del run
            del_wrapper = etree.SubElement(del_para, f"{{{W_NS}}}del")
            del_wrapper.set(f"{{{W_NS}}}id", str(change_id))
            del_wrapper.set(f"{{{W_NS}}}author", author)
            del_wrapper.set(f"{{{W_NS}}}date", ts)
            change_id += 1

            run = etree.SubElement(del_wrapper, f"{{{W_NS}}}r")
            del_text = etree.SubElement(run, f"{{{W_NS}}}delText")
            del_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            del_text.text = text

            # Insert before target paragraph
            parent = target_para.getparent()
            if parent is not None:
                idx = list(parent).index(target_para)
                parent.insert(idx, del_para)

    # Write modified XML back to the docx
    modified_xml = etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8")

    # Update the zip file
    with zipfile.ZipFile(output_path, "r") as zf_in:
        file_list = zf_in.namelist()
        file_contents = {}
        for name in file_list:
            if name == "word/document.xml":
                file_contents[name] = modified_xml
            else:
                file_contents[name] = zf_in.read(name)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
        for name, content in file_contents.items():
            zf_out.writestr(name, content)

    return True, (
        f"XML-based comparison complete. "
        f"Note: This is paragraph-level granularity only. "
        f"For word-level tracked changes, install pywin32 and use Word COM."
    )
