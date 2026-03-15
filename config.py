"""Configuration constants for the docx revision comparison system."""

from pathlib import Path

# ─── XML Namespaces ───────────────────────────────────────────────────────────
NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

# Relationship types
REL_COMMENTS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
REL_COMMENTS_EXTENDED = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)
REL_COMMENTS_IDS = (
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
)
REL_COMMENTS_EXTENSIBLE = (
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
)

# ─── Comment Mapping Thresholds ──────────────────────────────────────────────
# Minimum similarity ratio (0-1) for fuzzy text matching of comment anchors
FUZZY_MATCH_THRESHOLD = 0.6

# Minimum similarity for paragraph-level fallback matching
PARAGRAPH_MATCH_THRESHOLD = 0.4

# Maximum character distance to search for anchor text in the latest revision
MAX_SEARCH_WINDOW = 2000

# ─── Word COM Settings ───────────────────────────────────────────────────────
# Word comparison granularity: 0=CharLevel, 1=WordLevel
COMPARE_GRANULARITY = 1  # Word-level comparison

# Timeout in seconds for Word COM operations
COM_TIMEOUT = 120

# ─── Output Settings ─────────────────────────────────────────────────────────
# Author name used for system-generated tracked changes and comments
DEFAULT_AUTHOR = "Revision Compare System"

# Working directory for temporary files
TEMP_DIR = Path("_temp_compare")

# ─── Logging ──────────────────────────────────────────────────────────────────
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
LOG_LEVEL = "INFO"
