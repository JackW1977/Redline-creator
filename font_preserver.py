"""Extract and apply font/style information from the Latest Rev.

Ensures the output document's injected content (unmapped comments appendix,
any system-generated text) uses the same default font family, size, and
theme as the Latest Rev rather than Word's built-in defaults.

Also copies the full styles.xml from the Latest Rev into the comparison
output so that heading styles, list styles, etc. match exactly.
"""

from __future__ import annotations

import logging
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

logger = logging.getLogger(__name__)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass
class DocumentFonts:
    """Font settings extracted from a document's styles.xml."""
    # Default body font (from <w:rFonts> in the default paragraph style or docDefaults)
    body_font: str | None = None
    body_font_cs: str | None = None  # complex-script font
    body_font_east_asia: str | None = None
    body_size_half_points: str | None = None  # in half-points (e.g. "24" = 12pt)
    body_size_cs: str | None = None

    # Heading font (from Heading1 style if present)
    heading_font: str | None = None
    heading_size_half_points: str | None = None

    # Theme fonts (from <w:majorFont>/<w:minorFont> in theme1.xml)
    theme_major_font: str | None = None  # typically headings
    theme_minor_font: str | None = None  # typically body


def extract_fonts(docx_path: str | Path) -> DocumentFonts:
    """Extract the default font configuration from a .docx file.

    Reads styles.xml and theme1.xml to determine the document's
    default body and heading fonts.
    """
    docx_path = Path(docx_path)
    fonts = DocumentFonts()

    with zipfile.ZipFile(docx_path, "r") as zf:
        names = zf.namelist()

        # ── Parse styles.xml ──────────────────────────────────────────
        if "word/styles.xml" in names:
            styles_root = etree.fromstring(zf.read("word/styles.xml"))

            # 1) docDefaults → rPrDefault → rPr
            doc_defaults = styles_root.find(f"{{{W}}}docDefaults")
            if doc_defaults is not None:
                rpr_default = doc_defaults.find(f"{{{W}}}rPrDefault")
                if rpr_default is not None:
                    rpr = rpr_default.find(f"{{{W}}}rPr")
                    if rpr is not None:
                        _extract_rpr_fonts(rpr, fonts, target="body")

            # 2) Named styles: find "Normal" and "Heading1"
            for style in styles_root.iter(f"{{{W}}}style"):
                style_id = style.get(f"{{{W}}}styleId", "")
                rpr = style.find(f"{{{W}}}rPr")

                if style_id == "Normal" and rpr is not None:
                    _extract_rpr_fonts(rpr, fonts, target="body")

                if style_id in ("Heading1", "heading1") and rpr is not None:
                    _extract_rpr_fonts(rpr, fonts, target="heading")

        # ── Parse theme1.xml ──────────────────────────────────────────
        theme_path = None
        for name in names:
            if "theme1.xml" in name or "theme/theme1.xml" in name:
                theme_path = name
                break

        if theme_path:
            A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
            theme_root = etree.fromstring(zf.read(theme_path))

            for major in theme_root.iter(f"{{{A_NS}}}majorFont"):
                latin = major.find(f"{{{A_NS}}}latin")
                if latin is not None:
                    fonts.theme_major_font = latin.get("typeface")
                break

            for minor in theme_root.iter(f"{{{A_NS}}}minorFont"):
                latin = minor.find(f"{{{A_NS}}}latin")
                if latin is not None:
                    fonts.theme_minor_font = latin.get("typeface")
                break

    # Resolve theme references
    if fonts.body_font is None and fonts.theme_minor_font:
        fonts.body_font = fonts.theme_minor_font
    if fonts.heading_font is None and fonts.theme_major_font:
        fonts.heading_font = fonts.theme_major_font

    logger.info(
        f"Extracted fonts: body='{fonts.body_font}' ({fonts.body_size_half_points}hp), "
        f"heading='{fonts.heading_font}', "
        f"theme major='{fonts.theme_major_font}', minor='{fonts.theme_minor_font}'"
    )
    return fonts


def _extract_rpr_fonts(rpr: etree._Element, fonts: DocumentFonts, target: str) -> None:
    """Extract font info from a <w:rPr> element into the DocumentFonts."""
    rfonts = rpr.find(f"{{{W}}}rFonts")
    if rfonts is not None:
        ascii_font = rfonts.get(f"{{{W}}}ascii")
        cs_font = rfonts.get(f"{{{W}}}cs")
        ea_font = rfonts.get(f"{{{W}}}eastAsia")
        hAnsi = rfonts.get(f"{{{W}}}hAnsi")

        # Use ascii or hAnsi (they're usually the same)
        font_name = ascii_font or hAnsi

        if target == "body":
            if font_name:
                fonts.body_font = font_name
            if cs_font:
                fonts.body_font_cs = cs_font
            if ea_font:
                fonts.body_font_east_asia = ea_font
        elif target == "heading":
            if font_name:
                fonts.heading_font = font_name

    sz = rpr.find(f"{{{W}}}sz")
    if sz is not None:
        val = sz.get(f"{{{W}}}val")
        if target == "body":
            fonts.body_size_half_points = val
        elif target == "heading":
            fonts.heading_size_half_points = val

    sz_cs = rpr.find(f"{{{W}}}szCs")
    if sz_cs is not None and target == "body":
        fonts.body_size_cs = sz_cs.get(f"{{{W}}}val")


def apply_fonts_to_rpr(rpr: etree._Element, fonts: DocumentFonts, is_heading: bool = False) -> None:
    """Apply the extracted font settings to a <w:rPr> element.

    This ensures any text we generate (unmapped comments appendix, etc.)
    matches the Latest Rev's font setup.
    """
    font_name = fonts.heading_font if is_heading else fonts.body_font
    size = fonts.heading_size_half_points if is_heading else fonts.body_size_half_points
    size_cs = fonts.body_size_cs

    if font_name:
        rfonts = rpr.find(f"{{{W}}}rFonts")
        if rfonts is None:
            rfonts = etree.SubElement(rpr, f"{{{W}}}rFonts")
            # Insert rFonts as first child of rPr for schema compliance
            rpr.insert(0, rfonts)
        rfonts.set(f"{{{W}}}ascii", font_name)
        rfonts.set(f"{{{W}}}hAnsi", font_name)
        if fonts.body_font_cs:
            rfonts.set(f"{{{W}}}cs", fonts.body_font_cs)
        if fonts.body_font_east_asia:
            rfonts.set(f"{{{W}}}eastAsia", fonts.body_font_east_asia)

    if size:
        sz = rpr.find(f"{{{W}}}sz")
        if sz is None:
            sz = etree.SubElement(rpr, f"{{{W}}}sz")
        sz.set(f"{{{W}}}val", size)

    if size_cs:
        szcs = rpr.find(f"{{{W}}}szCs")
        if szcs is None:
            szcs = etree.SubElement(rpr, f"{{{W}}}szCs")
        szcs.set(f"{{{W}}}val", size_cs)


def transplant_styles(
    latest_rev_path: str | Path,
    output_docx_path: str | Path,
) -> None:
    """Copy styles.xml and theme files from the Latest Rev into the output docx.

    This ensures the output document uses the exact same style definitions
    (fonts, sizes, colours, spacing) as the Latest Rev, even when Word COM
    comparison may have merged or altered some style entries.
    """
    latest_rev_path = Path(latest_rev_path)
    output_docx_path = Path(output_docx_path)

    # Read both zips
    latest_contents: dict[str, bytes] = {}
    with zipfile.ZipFile(latest_rev_path, "r") as zf:
        for name in zf.namelist():
            latest_contents[name] = zf.read(name)

    output_contents: dict[str, bytes] = {}
    with zipfile.ZipFile(output_docx_path, "r") as zf:
        for name in zf.namelist():
            output_contents[name] = zf.read(name)

    # Copy styles.xml from latest rev
    parts_to_copy = [
        "word/styles.xml",
        "word/theme/theme1.xml",
        "word/fontTable.xml",
    ]

    copied = []
    for part in parts_to_copy:
        if part in latest_contents:
            output_contents[part] = latest_contents[part]
            copied.append(part)

    if copied:
        # Rewrite the output zip
        with zipfile.ZipFile(output_docx_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, content in output_contents.items():
                zf.writestr(name, content)
        logger.info(f"Transplanted from Latest Rev: {', '.join(copied)}")
    else:
        logger.info("No style parts found to transplant.")
