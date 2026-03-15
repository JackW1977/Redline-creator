"""Main orchestrator for MS Word Redline Creator.

Compares two Word document revisions and produces a new document that:
1. Looks like the Latest Rev (formatting, layout, structure preserved)
2. Contains tracked changes showing differences from Early Rev
3. Optionally carries forward review comments from Early Rev, Latest Rev, or both

Usage:
    python compare_revisions.py early_rev.docx latest_rev.docx output.docx [options]

Example:
    python compare_revisions.py draft_v1.docx draft_v2.docx comparison_output.docx
    python compare_revisions.py draft_v1.docx draft_v2.docx output.docx --force-xml
    python compare_revisions.py draft_v1.docx draft_v2.docx output.docx --author "John Doe"
    python compare_revisions.py draft_v1.docx draft_v2.docx output.docx --comments both
    python compare_revisions.py draft_v1.docx draft_v2.docx output.docx --comments none
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
import time
from datetime import datetime
from pathlib import Path

from comment_extractor import ExtractedComment, extract_comments
from comment_inserter import insert_comments
from comment_mapper import MappingResult, MappingStrategy, map_comments
from config import DEFAULT_AUTHOR, LOG_FORMAT, LOG_LEVEL, TEMP_DIR
from font_preserver import transplant_styles
from text_extractor import extract_from_docx
from word_compare import compare_documents

logger = logging.getLogger(__name__)


def setup_logging(log_file: str | None = None, verbose: bool = False) -> None:
    """Configure logging."""
    level = logging.DEBUG if verbose else getattr(logging, LOG_LEVEL)
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    if log_file:
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))
    logging.basicConfig(level=level, format=LOG_FORMAT, handlers=handlers)


def run_comparison(
    early_rev_path: str | Path,
    latest_rev_path: str | Path,
    output_path: str | Path,
    author: str = DEFAULT_AUTHOR,
    force_xml: bool = False,
    comment_source: str = "early",
    skip_comments: bool = False,
    log_file: str | None = None,
    verbose: bool = False,
) -> dict:
    """Run the full comparison pipeline.

    Args:
        early_rev_path: Path to the earlier revision.
        latest_rev_path: Path to the later revision.
        output_path: Path for the output document.
        author: Author name for tracked changes.
        force_xml: Force pure-XML comparison (skip COM).
        comment_source: Which files to pull comments from:
            "none" - no comments, "early" - Early Rev only,
            "latest" - Latest Rev only, "both" - merge from both.
        skip_comments: Legacy flag; if True, overrides comment_source to "none".
        log_file: Optional path for detailed log output.
        verbose: Enable debug logging.

    Returns:
        Dict with pipeline results and statistics.
    """
    # Legacy compatibility: skip_comments overrides comment_source
    if skip_comments:
        comment_source = "none"
    comment_source = comment_source.lower()
    if comment_source not in ("none", "early", "latest", "both"):
        comment_source = "early"

    setup_logging(log_file, verbose)
    start_time = time.time()

    early_rev_path = Path(early_rev_path)
    latest_rev_path = Path(latest_rev_path)
    output_path = Path(output_path)

    results = {
        "early_rev": str(early_rev_path),
        "latest_rev": str(latest_rev_path),
        "output": str(output_path),
        "success": False,
        "steps": {},
        "comment_source": comment_source,
        "comments": {
            "total_extracted": 0,
            "early_extracted": 0,
            "latest_extracted": 0,
            "mapped_exact": 0,
            "mapped_fuzzy": 0,
            "mapped_paragraph": 0,
            "mapped_heading": 0,
            "unmapped": 0,
        },
        "errors": [],
        "warnings": [],
    }

    # ── Validate inputs ──────────────────────────────────────────────────
    logger.info("=" * 60)
    logger.info("MS Word Redline Creator")
    logger.info("=" * 60)
    _source_labels = {"none": "DISABLED", "early": "Early Rev",
                      "latest": "Latest Rev", "both": "Both (merged)"}
    logger.info(f"  Comment carry-over: {_source_labels.get(comment_source, comment_source)}")

    if not early_rev_path.exists():
        results["errors"].append(f"Early revision not found: {early_rev_path}")
        logger.error(results["errors"][-1])
        return results

    if not latest_rev_path.exists():
        results["errors"].append(f"Latest revision not found: {latest_rev_path}")
        logger.error(results["errors"][-1])
        return results

    if not early_rev_path.suffix.lower() == ".docx":
        results["errors"].append(f"Early revision must be .docx: {early_rev_path}")
        logger.error(results["errors"][-1])
        return results

    if not latest_rev_path.suffix.lower() == ".docx":
        results["errors"].append(f"Latest revision must be .docx: {latest_rev_path}")
        logger.error(results["errors"][-1])
        return results

    # ── Step 1: Extract comments from source file(s) ───────────────────
    early_comments: list[ExtractedComment] = []
    latest_comments: list[ExtractedComment] = []
    comments: list[ExtractedComment] = []

    if comment_source == "none":
        logger.info("")
        logger.info("Step 1: SKIPPED (comment carry-over disabled).")
    else:
        logger.info("")
        step_start = time.time()

        # Extract from Early Rev if needed
        if comment_source in ("early", "both"):
            logger.info("Step 1: Extracting comments from Early Rev...")
            try:
                early_comments = extract_comments(early_rev_path)
                results["comments"]["early_extracted"] = len(early_comments)
                logger.info(f"  Found {len(early_comments)} comment(s) in Early Rev.")
                for c in early_comments:
                    logger.debug(
                        f"  Comment {c.comment_id} by {c.author}: "
                        f"'{c.text[:60]}...' anchored to '{c.anchor_text[:40]}...'"
                    )
            except Exception as e:
                msg = f"Failed to extract comments from Early Rev: {e}"
                results["warnings"].append(msg)
                logger.warning(f"  {msg}")

        # Extract from Latest Rev if needed
        if comment_source in ("latest", "both"):
            logger.info("Step 1: Extracting comments from Latest Rev...")
            try:
                latest_comments = extract_comments(latest_rev_path)
                results["comments"]["latest_extracted"] = len(latest_comments)
                logger.info(f"  Found {len(latest_comments)} comment(s) in Latest Rev.")
                for c in latest_comments:
                    logger.debug(
                        f"  Comment {c.comment_id} by {c.author}: "
                        f"'{c.text[:60]}...' anchored to '{c.anchor_text[:40]}...'"
                    )
            except Exception as e:
                msg = f"Failed to extract comments from Latest Rev: {e}"
                results["warnings"].append(msg)
                logger.warning(f"  {msg}")

        comments = early_comments + latest_comments
        results["comments"]["total_extracted"] = len(comments)
        results["steps"]["extract_comments"] = {
            "success": True,
            "count": len(comments),
            "early_count": len(early_comments),
            "latest_count": len(latest_comments),
            "duration": time.time() - step_start,
        }

    # ── Step 2: Generate tracked changes via document comparison ─────────
    logger.info("")
    logger.info("Step 2: Generating tracked changes...")
    step_start = time.time()

    # Use a temp path for the comparison output before adding comments
    temp_comparison = output_path.parent / f"_temp_comparison_{output_path.name}"

    try:
        success, msg = compare_documents(
            early_rev_path, latest_rev_path, temp_comparison,
            author=author, force_xml=force_xml,
        )
        results["steps"]["comparison"] = {
            "success": success,
            "message": msg,
            "method": "xml" if force_xml else "auto",
            "duration": time.time() - step_start,
        }
        if success:
            logger.info(f"  {msg}")
        else:
            results["errors"].append(msg)
            logger.error(f"  {msg}")
            return results
    except Exception as e:
        msg = f"Document comparison failed: {e}"
        results["errors"].append(msg)
        logger.error(f"  {msg}")
        results["steps"]["comparison"] = {
            "success": False,
            "error": msg,
            "duration": time.time() - step_start,
        }
        return results

    # ── Step 2b: Transplant styles/fonts from Latest Rev ────────────────
    try:
        transplant_styles(latest_rev_path, temp_comparison)
        logger.info("  Transplanted styles.xml, theme, and fontTable from Latest Rev.")
    except Exception as e:
        results["warnings"].append(f"Style transplant warning: {e}")
        logger.warning(f"  Style transplant warning (non-fatal): {e}")

    # ── Step 3: Map comments to Latest Rev locations ─────────────────────
    mapping_results: list[MappingResult] = []
    if comments:
        logger.info("")
        logger.info("Step 3: Mapping comments to Latest Rev locations...")
        step_start = time.time()

        try:
            latest_paras = extract_from_docx(latest_rev_path)

            # Map Early Rev comments (text may have moved between revisions)
            if early_comments:
                early_paras = extract_from_docx(early_rev_path)
                early_mapping = map_comments(early_comments, early_paras, latest_paras)
                mapping_results.extend(early_mapping)
                logger.info(f"  Mapped {len(early_mapping)} Early Rev comment(s).")

            # Map Latest Rev comments (already at correct locations — use direct mapping)
            if latest_comments:
                # Latest Rev comments are at their original locations, which match
                # the output structure. Map them using latest_paras as both source
                # and target for exact matching.
                latest_mapping = map_comments(latest_comments, latest_paras, latest_paras)
                mapping_results.extend(latest_mapping)
                logger.info(f"  Mapped {len(latest_mapping)} Latest Rev comment(s).")

            # Tally mapping strategies
            for r in mapping_results:
                if r.strategy == MappingStrategy.EXACT_MATCH:
                    results["comments"]["mapped_exact"] += 1
                elif r.strategy == MappingStrategy.FUZZY_SUBSTRING:
                    results["comments"]["mapped_fuzzy"] += 1
                elif r.strategy == MappingStrategy.PARAGRAPH_MATCH:
                    results["comments"]["mapped_paragraph"] += 1
                elif r.strategy == MappingStrategy.HEADING_PROXIMITY:
                    results["comments"]["mapped_heading"] += 1
                elif r.strategy == MappingStrategy.UNMAPPED:
                    results["comments"]["unmapped"] += 1

            results["steps"]["map_comments"] = {
                "success": True,
                "duration": time.time() - step_start,
            }

            mapped_count = len(mapping_results) - results["comments"]["unmapped"]
            logger.info(
                f"  Total: {mapped_count}/{len(comments)} comments mapped: "
                f"{results['comments']['mapped_exact']} exact, "
                f"{results['comments']['mapped_fuzzy']} fuzzy, "
                f"{results['comments']['mapped_paragraph']} paragraph, "
                f"{results['comments']['mapped_heading']} heading. "
                f"{results['comments']['unmapped']} unmapped."
            )

        except Exception as e:
            msg = f"Comment mapping failed: {e}"
            results["warnings"].append(msg)
            logger.warning(f"  {msg}")
            mapping_results = []
            results["steps"]["map_comments"] = {
                "success": False,
                "error": msg,
                "duration": time.time() - step_start,
            }
    else:
        logger.info("")
        logger.info("Step 3: No comments to map (skipping).")

    # ── Step 4: Insert comments into output document ─────────────────────
    if mapping_results:
        logger.info("")
        logger.info("Step 4: Inserting comments into output document...")
        step_start = time.time()

        try:
            success, msg, log_entries = insert_comments(
                temp_comparison, output_path, mapping_results,
                latest_rev_path=latest_rev_path,
            )
            results["steps"]["insert_comments"] = {
                "success": success,
                "message": msg,
                "log_entries": log_entries,
                "duration": time.time() - step_start,
            }
            logger.info(f"  {msg}")
            for entry in log_entries:
                logger.debug(f"  {entry}")

        except Exception as e:
            msg = f"Comment insertion failed: {e}"
            results["warnings"].append(msg)
            logger.warning(f"  {msg}")
            # Fall back to the comparison document without comments
            import shutil
            shutil.copy2(temp_comparison, output_path)
            results["steps"]["insert_comments"] = {
                "success": False,
                "error": msg,
                "duration": time.time() - step_start,
            }
    else:
        # No comments to insert - just use the comparison output
        import shutil
        shutil.copy2(temp_comparison, output_path)
        logger.info("")
        logger.info("Step 4: No comments to insert (using comparison output as-is).")

    # ── Cleanup ──────────────────────────────────────────────────────────
    try:
        if temp_comparison.exists():
            temp_comparison.unlink()
    except Exception:
        pass

    # ── Summary ──────────────────────────────────────────────────────────
    elapsed = time.time() - start_time
    results["success"] = True
    results["duration"] = elapsed

    logger.info("")
    logger.info("=" * 60)
    logger.info("COMPLETE")
    logger.info(f"  Output: {output_path}")
    logger.info(f"  Duration: {elapsed:.1f}s")
    if comments:
        src_detail = ""
        if results["comments"]["early_extracted"] and results["comments"]["latest_extracted"]:
            src_detail = (f" ({results['comments']['early_extracted']} early + "
                          f"{results['comments']['latest_extracted']} latest)")
        logger.info(
            f"  Comments: {results['comments']['total_extracted']} extracted{src_detail}, "
            f"{results['comments']['total_extracted'] - results['comments']['unmapped']} mapped, "
            f"{results['comments']['unmapped']} unmapped"
        )
    logger.info("=" * 60)

    return results


def main():
    parser = argparse.ArgumentParser(
        description="Compare two Word document revisions and generate output with tracked changes and comments.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s early.docx latest.docx output.docx
  %(prog)s early.docx latest.docx output.docx --force-xml
  %(prog)s early.docx latest.docx output.docx --author "Jane Doe" --verbose
  %(prog)s early.docx latest.docx output.docx --log comparison.log --report report.json
        """,
    )
    parser.add_argument("early_rev", help="Path to the earlier revision (.docx)")
    parser.add_argument("latest_rev", help="Path to the later revision (.docx)")
    parser.add_argument("output", help="Path for the output document (.docx)")
    parser.add_argument(
        "--author",
        default=DEFAULT_AUTHOR,
        help=f"Author name for tracked changes (default: '{DEFAULT_AUTHOR}')",
    )
    parser.add_argument(
        "--force-xml",
        action="store_true",
        help="Force pure-XML comparison (skip Word COM automation)",
    )
    parser.add_argument(
        "--comments",
        choices=["none", "early", "latest", "both"],
        default="early",
        help="Which file(s) to carry comments from: none, early (default), latest, both",
    )
    parser.add_argument(
        "--skip-comments",
        action="store_true",
        help="Legacy flag; equivalent to --comments none",
    )
    parser.add_argument(
        "--log",
        metavar="FILE",
        help="Write detailed log to file",
    )
    parser.add_argument(
        "--report",
        metavar="FILE",
        help="Write JSON report of pipeline results",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose (debug) logging",
    )

    args = parser.parse_args()

    comment_source = "none" if args.skip_comments else args.comments
    results = run_comparison(
        early_rev_path=args.early_rev,
        latest_rev_path=args.latest_rev,
        output_path=args.output,
        author=args.author,
        force_xml=args.force_xml,
        comment_source=comment_source,
        log_file=args.log,
        verbose=args.verbose,
    )

    # Write JSON report if requested
    if args.report:
        report_path = Path(args.report)
        # Remove non-serializable items
        clean_results = json.loads(json.dumps(results, default=str))
        report_path.write_text(
            json.dumps(clean_results, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        logger.info(f"Report written to {args.report}")

    sys.exit(0 if results["success"] else 1)


if __name__ == "__main__":
    main()
