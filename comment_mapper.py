"""Map comments from Early Rev to corresponding locations in Latest Rev.

Uses a multi-strategy approach:
1. Exact anchor text match
2. Fuzzy substring match (SequenceMatcher)
3. Paragraph-level fuzzy match using surrounding context
4. Heading/style-based proximity match
5. Unmapped fallback

Each mapped comment gets a MappingResult indicating confidence and strategy used.
"""

from __future__ import annotations

import difflib
import logging
import re
from dataclasses import dataclass
from enum import Enum

from comment_extractor import ExtractedComment
from config import FUZZY_MATCH_THRESHOLD, PARAGRAPH_MATCH_THRESHOLD
from text_extractor import StructuredParagraph

logger = logging.getLogger(__name__)


class MappingStrategy(Enum):
    EXACT_MATCH = "exact_match"
    FUZZY_SUBSTRING = "fuzzy_substring"
    PARAGRAPH_MATCH = "paragraph_match"
    HEADING_PROXIMITY = "heading_proximity"
    UNMAPPED = "unmapped"


@dataclass
class MappingResult:
    """Result of mapping a single comment to the latest revision."""
    comment: ExtractedComment
    # Target paragraph in the latest revision
    target_paragraph: StructuredParagraph | None = None
    # The text in the target to anchor the comment to (may differ from original)
    target_anchor_text: str | None = None
    # Character offset within the target paragraph where anchor starts
    target_char_offset: int | None = None
    # Length of anchor in target
    target_anchor_length: int | None = None
    # Which strategy was used
    strategy: MappingStrategy = MappingStrategy.UNMAPPED
    # Confidence score 0.0-1.0
    confidence: float = 0.0
    # Human-readable note about the mapping
    note: str = ""


def _normalize(text: str) -> str:
    """Normalize text for comparison: collapse whitespace, lowercase."""
    return re.sub(r"\s+", " ", text.strip().lower())


def _find_exact_match(
    anchor_text: str,
    paragraphs: list[StructuredParagraph],
) -> MappingResult | None:
    """Strategy 1: Find exact anchor text in a paragraph of the latest rev."""
    if not anchor_text or not anchor_text.strip():
        return None

    for para in paragraphs:
        idx = para.text.find(anchor_text)
        if idx != -1:
            return MappingResult(
                comment=None,  # Will be set by caller
                target_paragraph=para,
                target_anchor_text=anchor_text,
                target_char_offset=idx,
                target_anchor_length=len(anchor_text),
                strategy=MappingStrategy.EXACT_MATCH,
                confidence=1.0,
                note="Exact anchor text found in latest revision.",
            )

    # Try case-insensitive exact match
    anchor_lower = anchor_text.lower()
    for para in paragraphs:
        idx = para.text.lower().find(anchor_lower)
        if idx != -1:
            matched_text = para.text[idx : idx + len(anchor_text)]
            return MappingResult(
                comment=None,
                target_paragraph=para,
                target_anchor_text=matched_text,
                target_char_offset=idx,
                target_anchor_length=len(matched_text),
                strategy=MappingStrategy.EXACT_MATCH,
                confidence=0.95,
                note="Case-insensitive exact match.",
            )

    return None


def _find_fuzzy_substring(
    anchor_text: str,
    paragraphs: list[StructuredParagraph],
    threshold: float = FUZZY_MATCH_THRESHOLD,
) -> MappingResult | None:
    """Strategy 2: Find best fuzzy substring match across paragraphs."""
    if not anchor_text or len(anchor_text) < 3:
        return None

    norm_anchor = _normalize(anchor_text)
    best_score = 0.0
    best_para = None
    best_match_start = 0
    best_match_len = 0

    for para in paragraphs:
        if not para.text.strip():
            continue

        norm_para = _normalize(para.text)

        # Use SequenceMatcher for overall similarity
        sm = difflib.SequenceMatcher(None, norm_anchor, norm_para)

        # Find longest common substring ratio
        match = sm.find_longest_match(0, len(norm_anchor), 0, len(norm_para))
        if match.size == 0:
            continue

        # Score based on how much of the anchor is covered
        coverage = match.size / len(norm_anchor)
        # Also consider overall ratio
        ratio = sm.ratio()
        score = max(coverage, ratio)

        if score > best_score:
            best_score = score
            best_para = para
            # Map back to original text positions (approximate)
            best_match_start = match.b
            best_match_len = match.size

    if best_score >= threshold and best_para is not None:
        # Try to extract the best matching substring from the target
        target_text = best_para.text
        # Use get_close_matches approach: find the most similar window
        anchor_len = len(anchor_text)
        best_window_score = 0.0
        best_window_start = 0
        best_window_text = target_text

        # Slide a window of similar size across the paragraph
        for window_start in range(max(1, len(target_text) - anchor_len + 1)):
            window_end = min(window_start + anchor_len + 10, len(target_text))
            window = target_text[window_start:window_end]
            wscore = difflib.SequenceMatcher(None, anchor_text, window).ratio()
            if wscore > best_window_score:
                best_window_score = wscore
                best_window_start = window_start
                best_window_text = window

        return MappingResult(
            comment=None,
            target_paragraph=best_para,
            target_anchor_text=best_window_text,
            target_char_offset=best_window_start,
            target_anchor_length=len(best_window_text),
            strategy=MappingStrategy.FUZZY_SUBSTRING,
            confidence=best_score,
            note=f"Fuzzy match (score={best_score:.2f}). Original anchor: '{anchor_text[:80]}...'",
        )

    return None


def _find_paragraph_match(
    comment: ExtractedComment,
    paragraphs: list[StructuredParagraph],
    threshold: float = PARAGRAPH_MATCH_THRESHOLD,
) -> MappingResult | None:
    """Strategy 3: Match the full anchor paragraph text to find the closest paragraph."""
    if not comment.anchor_paragraph_text:
        return None

    norm_anchor_para = _normalize(comment.anchor_paragraph_text)
    best_score = 0.0
    best_para = None

    for para in paragraphs:
        if not para.text.strip():
            continue

        norm_text = _normalize(para.text)
        score = difflib.SequenceMatcher(None, norm_anchor_para, norm_text).ratio()

        # Bonus for matching style
        if comment.anchor_paragraph_style and para.style == comment.anchor_paragraph_style:
            score = min(1.0, score + 0.1)

        # Bonus for matching table context
        if comment.anchor_in_table and para.in_table:
            score = min(1.0, score + 0.05)

        if score > best_score:
            best_score = score
            best_para = para

    if best_score >= threshold and best_para is not None:
        return MappingResult(
            comment=None,
            target_paragraph=best_para,
            target_anchor_text=None,  # Anchor at paragraph level
            target_char_offset=0,
            target_anchor_length=len(best_para.text),
            strategy=MappingStrategy.PARAGRAPH_MATCH,
            confidence=best_score,
            note=(
                f"Paragraph-level match (score={best_score:.2f}). "
                f"Original paragraph: '{comment.anchor_paragraph_text[:80]}...'"
            ),
        )

    return None


def _find_heading_proximity(
    comment: ExtractedComment,
    early_paragraphs: list[StructuredParagraph],
    latest_paragraphs: list[StructuredParagraph],
) -> MappingResult | None:
    """Strategy 4: Find the nearest heading in Early Rev, then match that heading in Latest Rev.

    Anchors the comment to the paragraph immediately after the matched heading.
    """
    if not early_paragraphs or not latest_paragraphs:
        return None

    # Find the heading closest to (and before) the original anchor
    anchor_para_idx = comment.anchor_paragraph_index
    if anchor_para_idx is None:
        return None

    # Find preceding heading in early rev
    nearest_heading = None
    for para in reversed(early_paragraphs[:anchor_para_idx + 1]):
        if para.style and para.style.startswith("Heading"):
            nearest_heading = para
            break

    if nearest_heading is None:
        return None

    # Find the same heading in latest rev
    heading_text_norm = _normalize(nearest_heading.text)
    best_heading_match = None
    best_heading_score = 0.0

    for para in latest_paragraphs:
        if not (para.style and para.style.startswith("Heading")):
            continue
        score = difflib.SequenceMatcher(
            None, heading_text_norm, _normalize(para.text)
        ).ratio()
        if score > best_heading_score:
            best_heading_score = score
            best_heading_match = para

    if best_heading_match is None or best_heading_score < 0.5:
        return None

    # Anchor to the paragraph after the heading (or the heading itself)
    heading_idx = best_heading_match.index
    target = best_heading_match
    for para in latest_paragraphs:
        if para.index > heading_idx and para.text.strip():
            target = para
            break

    return MappingResult(
        comment=None,
        target_paragraph=target,
        target_anchor_text=None,
        target_char_offset=0,
        target_anchor_length=len(target.text) if target.text else 0,
        strategy=MappingStrategy.HEADING_PROXIMITY,
        confidence=best_heading_score * 0.7,  # Discount for indirect match
        note=(
            f"Mapped via heading proximity: '{nearest_heading.text[:60]}' "
            f"-> '{best_heading_match.text[:60]}' (heading score={best_heading_score:.2f})"
        ),
    )


def map_comments(
    comments: list[ExtractedComment],
    early_paragraphs: list[StructuredParagraph],
    latest_paragraphs: list[StructuredParagraph],
) -> list[MappingResult]:
    """Map all comments from the Early Rev to locations in the Latest Rev.

    Tries strategies in order of preference:
    1. Exact anchor text match
    2. Fuzzy substring match
    3. Paragraph-level match
    4. Heading proximity match
    5. Mark as unmapped

    Args:
        comments: Comments extracted from Early Rev.
        early_paragraphs: Paragraphs from Early Rev (for context).
        latest_paragraphs: Paragraphs from Latest Rev (target).

    Returns:
        List of MappingResult objects, one per comment.
    """
    results: list[MappingResult] = []

    # Index early paragraphs for heading proximity lookups
    for comment in comments:
        result = None

        # Strategy 1: Exact match
        result = _find_exact_match(comment.anchor_text, latest_paragraphs)
        if result and result.confidence >= 0.9:
            result.comment = comment
            results.append(result)
            logger.info(
                f"Comment {comment.comment_id} ({comment.author}): "
                f"EXACT match -> para {result.target_paragraph.index}"
            )
            continue

        # Strategy 2: Fuzzy substring
        fuzzy_result = _find_fuzzy_substring(comment.anchor_text, latest_paragraphs)
        if fuzzy_result and fuzzy_result.confidence >= FUZZY_MATCH_THRESHOLD:
            # Keep the better of exact (case-insensitive) and fuzzy
            if result and result.confidence >= fuzzy_result.confidence:
                result.comment = comment
                results.append(result)
            else:
                fuzzy_result.comment = comment
                results.append(fuzzy_result)
            logger.info(
                f"Comment {comment.comment_id} ({comment.author}): "
                f"FUZZY match (confidence={fuzzy_result.confidence:.2f}) "
                f"-> para {fuzzy_result.target_paragraph.index}"
            )
            continue

        # Strategy 3: Paragraph-level match
        para_result = _find_paragraph_match(comment, latest_paragraphs)
        if para_result and para_result.confidence >= PARAGRAPH_MATCH_THRESHOLD:
            para_result.comment = comment
            results.append(para_result)
            logger.info(
                f"Comment {comment.comment_id} ({comment.author}): "
                f"PARAGRAPH match (confidence={para_result.confidence:.2f}) "
                f"-> para {para_result.target_paragraph.index}"
            )
            continue

        # Strategy 4: Heading proximity
        heading_result = _find_heading_proximity(
            comment, early_paragraphs, latest_paragraphs
        )
        if heading_result:
            heading_result.comment = comment
            results.append(heading_result)
            logger.info(
                f"Comment {comment.comment_id} ({comment.author}): "
                f"HEADING proximity (confidence={heading_result.confidence:.2f}) "
                f"-> para {heading_result.target_paragraph.index}"
            )
            continue

        # Strategy 5: Unmapped
        unmapped = MappingResult(
            comment=comment,
            strategy=MappingStrategy.UNMAPPED,
            confidence=0.0,
            note=(
                f"Could not map comment. Original anchor: '{comment.anchor_text[:100]}'. "
                f"Original paragraph: '{comment.anchor_paragraph_text[:100]}'"
            ),
        )
        results.append(unmapped)
        logger.warning(
            f"Comment {comment.comment_id} ({comment.author}): UNMAPPED. "
            f"Anchor: '{comment.anchor_text[:60]}'"
        )

    return results
