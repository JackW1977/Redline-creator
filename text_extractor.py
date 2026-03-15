"""Extract structured text from DOCX files for comparison and comment mapping.

Parses the document.xml inside a .docx to produce a list of StructuredParagraph
objects that preserve the mapping between text content and its XML location.
This is used by the comment mapper to find where anchor text appears in the
latest revision.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

from lxml import etree

from config import NAMESPACES

W = NAMESPACES["w"]


@dataclass
class RunInfo:
    """A single run of text within a paragraph."""
    text: str
    xml_element: etree._Element | None = None
    # Character offset within the paragraph
    char_offset: int = 0


@dataclass
class StructuredParagraph:
    """A paragraph with its text content and XML location info."""
    # Index of this paragraph in document order
    index: int
    # Full concatenated text of the paragraph
    text: str = ""
    # Individual runs that compose this paragraph
    runs: list[RunInfo] = field(default_factory=list)
    # The <w:p> XML element (only populated when working with unpacked XML)
    xml_element: etree._Element | None = None
    # Style name if available (e.g., "Heading1")
    style: str | None = None
    # Whether this paragraph is inside a table
    in_table: bool = False
    # Table cell coordinates if in a table: (table_idx, row, col)
    table_cell: tuple[int, int, int] | None = None
    # Whether this paragraph is in a header/footer
    in_header_footer: bool = False


def _get_paragraph_text(p_elem: etree._Element) -> tuple[str, list[RunInfo]]:
    """Extract text and run info from a <w:p> element."""
    runs = []
    full_text = ""

    for r_elem in p_elem.iter(f"{{{W}}}r"):
        run_text = ""
        for t_elem in r_elem.iter(f"{{{W}}}t"):
            if t_elem.text:
                run_text += t_elem.text
        # Also capture text in deleted runs for mapping purposes
        for dt_elem in r_elem.iter(f"{{{W}}}delText"):
            if dt_elem.text:
                run_text += dt_elem.text

        if run_text:
            runs.append(RunInfo(
                text=run_text,
                xml_element=r_elem,
                char_offset=len(full_text),
            ))
            full_text += run_text

    return full_text, runs


def _get_paragraph_style(p_elem: etree._Element) -> str | None:
    """Extract the style ID from a paragraph's <w:pPr><w:pStyle>."""
    ppr = p_elem.find(f"{{{W}}}pPr")
    if ppr is not None:
        pstyle = ppr.find(f"{{{W}}}pStyle")
        if pstyle is not None:
            return pstyle.get(f"{{{W}}}val")
    return None


def extract_from_xml(doc_xml_path: str | Path) -> list[StructuredParagraph]:
    """Extract structured paragraphs from a document.xml file.

    Args:
        doc_xml_path: Path to the document.xml file (unpacked).

    Returns:
        List of StructuredParagraph objects in document order.
    """
    tree = etree.parse(str(doc_xml_path))
    root = tree.getroot()
    paragraphs = []
    para_index = 0
    table_index = 0

    # Process body content
    body = root.find(f"{{{W}}}body")
    if body is None:
        return paragraphs

    for child in body:
        tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""

        if tag == "p":
            text, runs = _get_paragraph_text(child)
            paragraphs.append(StructuredParagraph(
                index=para_index,
                text=text,
                runs=runs,
                xml_element=child,
                style=_get_paragraph_style(child),
            ))
            para_index += 1

        elif tag == "tbl":
            for row_idx, tr in enumerate(child.iter(f"{{{W}}}tr")):
                for col_idx, tc in enumerate(tr.iter(f"{{{W}}}tc")):
                    for p in tc.iter(f"{{{W}}}p"):
                        text, runs = _get_paragraph_text(p)
                        paragraphs.append(StructuredParagraph(
                            index=para_index,
                            text=text,
                            runs=runs,
                            xml_element=p,
                            style=_get_paragraph_style(p),
                            in_table=True,
                            table_cell=(table_index, row_idx, col_idx),
                        ))
                        para_index += 1
            table_index += 1

        elif tag == "sdt":
            # Structured document tags - extract paragraphs from content
            for p in child.iter(f"{{{W}}}p"):
                text, runs = _get_paragraph_text(p)
                paragraphs.append(StructuredParagraph(
                    index=para_index,
                    text=text,
                    runs=runs,
                    xml_element=p,
                    style=_get_paragraph_style(p),
                ))
                para_index += 1

    return paragraphs


def extract_from_docx(docx_path: str | Path) -> list[StructuredParagraph]:
    """Extract structured paragraphs directly from a .docx file.

    Note: xml_element references will NOT be usable for modification since
    they point into an in-memory parse. Use extract_from_xml on an unpacked
    directory for modification workflows.

    Args:
        docx_path: Path to the .docx file.

    Returns:
        List of StructuredParagraph objects in document order.
    """
    docx_path = Path(docx_path)
    with zipfile.ZipFile(docx_path, "r") as zf:
        with zf.open("word/document.xml") as f:
            tree = etree.parse(f)

    root = tree.getroot()
    paragraphs = []
    para_index = 0
    table_index = 0

    body = root.find(f"{{{W}}}body")
    if body is None:
        return paragraphs

    for child in body:
        tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""

        if tag == "p":
            text, runs = _get_paragraph_text(child)
            paragraphs.append(StructuredParagraph(
                index=para_index,
                text=text,
                runs=runs,
                xml_element=child,
                style=_get_paragraph_style(child),
            ))
            para_index += 1

        elif tag == "tbl":
            for row_idx, tr in enumerate(child.iter(f"{{{W}}}tr")):
                for col_idx, tc in enumerate(tr.iter(f"{{{W}}}tc")):
                    for p in tc.iter(f"{{{W}}}p"):
                        text, runs = _get_paragraph_text(p)
                        paragraphs.append(StructuredParagraph(
                            index=para_index,
                            text=text,
                            runs=runs,
                            xml_element=p,
                            style=_get_paragraph_style(p),
                            in_table=True,
                            table_cell=(table_index, row_idx, col_idx),
                        ))
                        para_index += 1
            table_index += 1

        elif tag == "sdt":
            for p in child.iter(f"{{{W}}}p"):
                text, runs = _get_paragraph_text(p)
                paragraphs.append(StructuredParagraph(
                    index=para_index,
                    text=text,
                    runs=runs,
                    xml_element=p,
                    style=_get_paragraph_style(p),
                ))
                para_index += 1

    return paragraphs


def build_text_index(paragraphs: list[StructuredParagraph]) -> str:
    """Build a single concatenated text from all paragraphs for substring search.

    Returns the full text with paragraph boundaries marked by newlines.
    """
    return "\n".join(p.text for p in paragraphs)


def find_paragraph_at_offset(
    paragraphs: list[StructuredParagraph],
    global_offset: int,
) -> StructuredParagraph | None:
    """Given a character offset in the concatenated text, find which paragraph it falls in."""
    current = 0
    for p in paragraphs:
        end = current + len(p.text)
        if current <= global_offset < end:
            return p
        current = end + 1  # +1 for the newline separator
    return None
