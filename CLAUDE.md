# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Location

Windows path: `C:\Users\wyzis\ClaudeProjects\docx-revision-compare`

## Project Overview

MS Word Redline Creator — a Windows application that compares two `.docx` files and produces a new document with tracked changes (redlines). Optionally carries forward review comments from source documents.

## Commands

**Run the application:**
```bash
python gui.py
```

**Install dependencies (Windows):**
```bash
install.bat           # Sets up venv, installs deps, creates desktop shortcut
# Or manually:
pip install -r requirements.txt
```

**Build standalone .exe:**
```bash
build.bat             # Creates dist/RedlineCreator/ and RedlineCreator_v1.0_Windows.zip
```

**Generate test documents:**
```bash
python create_test_docs.py   # Creates test_early_rev.docx and test_latest_rev.docx
```

There is no automated test suite. Manual testing uses the generated test documents.

## Architecture

The application follows a sequential pipeline orchestrated by `compare_revisions.py:run_comparison()`:

```
GUI (gui.py)
  └─→ run_comparison() [compare_revisions.py]
        ├─→ comment_extractor.py  — parse comments from source .docx files
        ├─→ word_compare.py       — diff the two documents (Word COM or XML)
        ├─→ comment_mapper.py     — map extracted comments to new text locations
        └─→ comment_inserter.py   — inject comments into the output .docx
```

### Key modules

| File | Role |
|------|------|
| `gui.py` | tkinter GUI (`RevisionCompareApp`); runs comparison in a background thread |
| `compare_revisions.py` | Orchestrator; validates inputs, calls all pipeline stages, returns stats dict |
| `word_compare.py` | Two comparison backends: `compare_with_word_com()` (pywin32 COM) and `compare_with_xml()` (pure Python fallback) |
| `comment_extractor.py` | Reads `comments.xml`, `commentsExtended.xml`, `document.xml` into `ExtractedComment` dataclasses |
| `comment_mapper.py` | Maps comments from Early Rev → Latest Rev using five strategies: EXACT_MATCH → FUZZY_SUBSTRING → PARAGRAPH_MATCH → HEADING_PROXIMITY → UNMAPPED |
| `comment_inserter.py` | Writes mapped comments into output .docx by directly editing Open XML ZIP internals; creates an "Unmapped Comments" appendix |
| `text_extractor.py` | Extracts text with XML location info (`StructuredParagraph`, `RunInfo`) for the mapper |
| `font_preserver.py` | Copies `styles.xml` from Latest Rev to output and applies matching fonts to injected content |
| `config.py` | Central constants: XML namespaces, fuzzy match thresholds, COM timeout, temp dir name |

### Comparison backends

- **Word COM** (`pywin32`, requires MS Word installed): highest fidelity, uses `Word.Application` via COM automation with a 120-second timeout.
- **Force XML / fallback**: pure Python `lxml`-based comparison; always available, cross-platform.

The GUI exposes a "Force XML comparison" checkbox; otherwise it tries COM first and falls back automatically.

### Comment mapping strategies (in order)

1. **EXACT_MATCH** — anchor text found verbatim in Latest Rev
2. **FUZZY_SUBSTRING** — ≥60% similarity match (configurable via `FUZZY_MATCH_THRESHOLD` in `config.py`)
3. **PARAGRAPH_MATCH** — ≥40% paragraph-level similarity (`PARAGRAPH_MATCH_THRESHOLD`)
4. **HEADING_PROXIMITY** — nearest heading context match
5. **UNMAPPED** — appended as an appendix in the output document

### Open XML manipulation

The inserter and extractor work directly with the `.docx` ZIP archive internals (`comments.xml`, `commentsExtended.xml`, `commentsIds.xml`, `commentsExtensible.xml`, `document.xml`, `_rels/`). All XML namespaces used are defined in `config.py`.

## Dependencies

- `lxml` — XML parsing (required)
- `pywin32` — Windows COM automation for Word (optional; enables Word COM mode)
- `tkinterdnd2` — drag-and-drop file support in GUI (optional)
- `python-docx` — only used by `create_test_docs.py` (testing utility, not runtime)
