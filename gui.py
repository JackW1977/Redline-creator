"""GUI for the MS Word Redline Creator.

A tkinter-based interface for comparing two Word document revisions.
Runs the comparison pipeline in a background thread to keep the UI responsive.

Usage:
    python gui.py
"""

from __future__ import annotations

import logging
import os
import subprocess
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from config import DEFAULT_AUTHOR

# Optional: drag-and-drop support via tkinterdnd2
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ---------------------------------------------------------------------------
# Logging handler that writes to a tkinter Text widget
# ---------------------------------------------------------------------------

class TextWidgetHandler(logging.Handler):
    """Routes log records into a scrolled Text widget (thread-safe)."""

    def __init__(self, text_widget: tk.Text):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record: logging.LogRecord):
        msg = self.format(record) + "\n"
        # schedule on the main thread
        self.text_widget.after(0, self._append, msg, record.levelno)

    def _append(self, msg: str, levelno: int):
        self.text_widget.configure(state="normal")
        tag = "error" if levelno >= logging.ERROR else (
            "warning" if levelno >= logging.WARNING else (
                "success" if "COMPLETE" in msg or "SKIPPED" in msg else "info"
            )
        )
        self.text_widget.insert(tk.END, msg, tag)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state="disabled")


# ---------------------------------------------------------------------------
# Hover tooltip
# ---------------------------------------------------------------------------

class ToolTip:
    """Dark-themed tooltip that appears when the mouse hovers over a widget."""

    BG = "#3b3b54"
    FG = "#cdd6f4"
    FG_DIM = "#a6adc8"
    BORDER = "#585b70"
    FONT = ("Segoe UI", 9)
    DELAY_MS = 400        # delay before showing
    WRAP_LENGTH = 340     # pixel width before wrapping

    def __init__(self, widget: tk.Widget, text: str):
        self.widget = widget
        self.text = text
        self._tip_window: tk.Toplevel | None = None
        self._after_id: str | None = None

        widget.bind("<Enter>", self._on_enter, add="+")
        widget.bind("<Leave>", self._on_leave, add="+")
        widget.bind("<ButtonPress>", self._on_leave, add="+")

    # -- public API ----------------------------------------------------------

    def update_text(self, text: str):
        self.text = text

    # -- internals -----------------------------------------------------------

    def _on_enter(self, event: tk.Event):
        self._cancel()
        self._after_id = self.widget.after(self.DELAY_MS, self._show)

    def _on_leave(self, _event: tk.Event):
        self._cancel()
        self._hide()

    def _cancel(self):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None

    def _show(self):
        if self._tip_window or not self.text:
            return
        # Position just below the widget
        x = self.widget.winfo_rootx() + 8
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4

        tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_attributes("-topmost", True)
        # Transparent background on Windows
        try:
            tw.wm_attributes("-transparentcolor", "")
        except Exception:
            pass
        tw.configure(bg=self.BORDER)

        inner = tk.Frame(tw, bg=self.BG, padx=10, pady=6,
                         highlightbackground=self.BORDER,
                         highlightthickness=1)
        inner.pack()

        label = tk.Label(
            inner, text=self.text, font=self.FONT,
            fg=self.FG, bg=self.BG,
            wraplength=self.WRAP_LENGTH, justify="left",
        )
        label.pack()

        tw.wm_geometry(f"+{x}+{y}")
        self._tip_window = tw

    def _hide(self):
        if self._tip_window:
            self._tip_window.destroy()
            self._tip_window = None


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------

class RevisionCompareApp:
    # Colour palette
    BG = "#1e1e2e"
    BG_LIGHT = "#2a2a3d"
    FG = "#cdd6f4"
    FG_DIM = "#6c7086"
    ACCENT = "#89b4fa"
    ACCENT_HOVER = "#74c7ec"
    GREEN = "#a6e3a1"
    RED = "#f38ba8"
    YELLOW = "#f9e2af"
    SURFACE = "#313244"
    BORDER = "#45475a"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("MS Word Redline Creator v1.0")
        self.root.configure(bg=self.BG)
        self.root.minsize(780, 620)
        self.root.geometry("860x700")

        # Try to set the window icon (silently skip if not available)
        try:
            self.root.iconbitmap(default="")
        except Exception:
            pass

        # Variables
        self.early_rev_var = tk.StringVar()
        self.latest_rev_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.author_var = tk.StringVar(value=DEFAULT_AUTHOR)
        self.force_xml_var = tk.BooleanVar(value=False)
        self.carry_comments_var = tk.BooleanVar(value=True)
        self.verbose_var = tk.BooleanVar(value=True)
        self.running = False
        self._output_was_auto = False  # tracks if output path was auto-generated

        # Configure ttk styles
        self._configure_styles()

        # Build UI
        self._build_ui()

        # Auto-fill output whenever Latest Revision changes
        self.latest_rev_var.trace_add("write", self._on_latest_rev_changed)

    # ----- Styles ---------------------------------------------------------

    def _configure_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("App.TFrame", background=self.BG)
        style.configure("Card.TFrame", background=self.BG_LIGHT)
        style.configure("App.TLabel", background=self.BG, foreground=self.FG,
                         font=("Segoe UI", 10))
        style.configure("Card.TLabel", background=self.BG_LIGHT, foreground=self.FG,
                         font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=self.BG, foreground=self.FG,
                         font=("Segoe UI", 16, "bold"))
        style.configure("Sub.TLabel", background=self.BG, foreground=self.FG_DIM,
                         font=("Segoe UI", 9))
        style.configure("Status.TLabel", background=self.BG, foreground=self.ACCENT,
                         font=("Segoe UI", 10))

        # Entry
        style.configure("App.TEntry", fieldbackground=self.SURFACE,
                         foreground=self.FG, insertcolor=self.FG,
                         bordercolor=self.BORDER, lightcolor=self.BORDER,
                         darkcolor=self.BORDER)
        style.map("App.TEntry",
                   bordercolor=[("focus", self.ACCENT)],
                   lightcolor=[("focus", self.ACCENT)])

        # Drop-highlight style for drag-and-drop visual feedback
        style.configure("DropHighlight.TEntry", fieldbackground="#3b3b54",
                         foreground=self.FG, insertcolor=self.FG,
                         bordercolor=self.GREEN, lightcolor=self.GREEN,
                         darkcolor=self.GREEN)

        # Buttons
        style.configure("Accent.TButton", background=self.ACCENT,
                         foreground="#1e1e2e", font=("Segoe UI", 10, "bold"),
                         padding=(16, 8), borderwidth=0)
        style.map("Accent.TButton",
                   background=[("active", self.ACCENT_HOVER),
                               ("disabled", self.SURFACE)],
                   foreground=[("disabled", self.FG_DIM)])

        style.configure("Browse.TButton", background=self.SURFACE,
                         foreground=self.FG, font=("Segoe UI", 9),
                         padding=(8, 4), borderwidth=0)
        style.map("Browse.TButton",
                   background=[("active", self.BORDER)])

        style.configure("Open.TButton", background=self.GREEN,
                         foreground="#1e1e2e", font=("Segoe UI", 9, "bold"),
                         padding=(10, 4), borderwidth=0)
        style.map("Open.TButton",
                   background=[("active", "#b5eab0"), ("disabled", self.SURFACE)],
                   foreground=[("disabled", self.FG_DIM)])

        # Checkbutton
        style.configure("App.TCheckbutton", background=self.BG_LIGHT,
                         foreground=self.FG, font=("Segoe UI", 9),
                         indicatorcolor=self.SURFACE)
        style.map("App.TCheckbutton",
                   indicatorcolor=[("selected", self.ACCENT)],
                   background=[("active", self.BG_LIGHT)])

        # Progressbar
        style.configure("Accent.Horizontal.TProgressbar",
                         troughcolor=self.SURFACE, background=self.ACCENT,
                         bordercolor=self.BG, lightcolor=self.ACCENT,
                         darkcolor=self.ACCENT)

    # ----- UI Layout ------------------------------------------------------

    def _build_ui(self):
        outer = ttk.Frame(self.root, style="App.TFrame", padding=20)
        outer.pack(fill="both", expand=True)

        # Title row
        title_frame = ttk.Frame(outer, style="App.TFrame")
        title_frame.pack(fill="x", pady=(0, 12))
        ttk.Label(title_frame, text="MS Word Redline Creator",
                  style="Header.TLabel").pack(side="left")
        ttk.Label(title_frame,
                  text="Compare two Word revisions and generate tracked changes",
                  style="Sub.TLabel").pack(side="left", padx=(12, 0), pady=(4, 0))

        help_btn = tk.Button(
            title_frame, text=" ? Help ", font=("Segoe UI", 9),
            bg=self.SURFACE, fg=self.ACCENT, activebackground=self.BORDER,
            activeforeground=self.ACCENT_HOVER, relief="flat", cursor="hand2",
            padx=8, pady=2, command=self._show_help,
        )
        help_btn.pack(side="right")
        ToolTip(help_btn, "Open the full help documentation.")

        # ── File selection card ──────────────────────────────────────────
        file_card = ttk.Frame(outer, style="Card.TFrame", padding=16)
        file_card.pack(fill="x", pady=(0, 10))

        self._file_row(
            file_card, "Early Revision:", self.early_rev_var, row=0,
            tooltip="The older version of the document (the \"before\" snapshot).\n"
                    "If comment carry-over is enabled, review comments in this\n"
                    "file will be extracted and re-anchored in the output.\n\n"
                    "Required: Must be a .docx file.",
        )
        self._file_row(
            file_card, "Latest Revision:", self.latest_rev_var, row=1,
            tooltip="The newer version of the document (the \"after\" snapshot).\n"
                    "The output inherits this file's fonts, styles, themes,\n"
                    "headers/footers, images, and layout.\n\n"
                    "Setting this auto-fills the Output File path.\n\n"
                    "Required: Must be a .docx file.",
        )
        self._output_row(
            file_card, "Output File:", self.output_var, row=2,
            tooltip="Where to save the comparison result (.docx).\n"
                    "Auto-populated as \"<LatestName>_redline.docx\" in the\n"
                    "same folder as the Latest Revision. Override manually\n"
                    "or via Save As if needed.\n\n"
                    "Required: Must end in .docx.",
        )

        # ── Options card ─────────────────────────────────────────────────
        opt_card = ttk.Frame(outer, style="Card.TFrame", padding=16)
        opt_card.pack(fill="x", pady=(0, 10))

        opt_row1 = ttk.Frame(opt_card, style="Card.TFrame")
        opt_row1.pack(fill="x")

        author_lbl = ttk.Label(opt_row1, text="Author:", style="Card.TLabel")
        author_lbl.pack(side="left")
        author_entry = ttk.Entry(opt_row1, textvariable=self.author_var,
                                  width=28, style="App.TEntry")
        author_entry.pack(side="left", padx=(6, 20))
        ToolTip(author_lbl,
                "The name stamped on every tracked change in the output.\n"
                "Appears when you hover over a redline in Word.\n\n"
                "Default: \"Revision Compare System\"")
        ToolTip(author_entry,
                "The name stamped on every tracked change in the output.\n"
                "Appears when you hover over a redline in Word.\n\n"
                "Default: \"Revision Compare System\"")

        force_xml_cb = ttk.Checkbutton(opt_row1, text="Force XML (no Word COM)",
                         variable=self.force_xml_var,
                         style="App.TCheckbutton")
        force_xml_cb.pack(side="left", padx=(0, 16))
        ToolTip(force_xml_cb,
                "Unchecked (default): Uses Word's built-in comparison\n"
                "engine via COM \u2014 highest-fidelity word/character diffs.\n\n"
                "Checked: Pure-Python XML diff. Paragraph-level only,\n"
                "but works without Word installed.\n\n"
                "Pre-condition: Microsoft Word must be installed for\n"
                "COM mode. Check this box if Word is unavailable.")

        verbose_cb = ttk.Checkbutton(opt_row1, text="Verbose logging",
                         variable=self.verbose_var,
                         style="App.TCheckbutton")
        verbose_cb.pack(side="left")
        ToolTip(verbose_cb,
                "When checked, shows debug-level detail in the log:\n"
                "comment mapping scores, paragraph indices,\n"
                "deduplication events, and XML operations.\n\n"
                "When unchecked, only high-level step summaries.")

        opt_row2 = ttk.Frame(opt_card, style="Card.TFrame")
        opt_row2.pack(fill="x", pady=(6, 0))

        carry_cb = ttk.Checkbutton(opt_row2, text="Carry over comments from Early Revision",
                         variable=self.carry_comments_var,
                         style="App.TCheckbutton")
        carry_cb.pack(side="left")
        ToolTip(carry_cb,
                "When checked, review comments from the Early Revision\n"
                "are extracted, mapped to the best matching location in\n"
                "the Latest Revision, and inserted into the output.\n\n"
                "When unchecked, only tracked changes (redlines) are\n"
                "generated \u2014 faster, no comments in the output.\n\n"
                "Pre-condition: Early Revision must contain comments.")

        # ── Action row ───────────────────────────────────────────────────
        action_frame = ttk.Frame(outer, style="App.TFrame")
        action_frame.pack(fill="x", pady=(0, 10))

        self.run_btn = ttk.Button(action_frame, text="  Run Comparison  ",
                                   style="Accent.TButton",
                                   command=self._on_run)
        self.run_btn.pack(side="left")
        ToolTip(self.run_btn,
                "Start the comparison pipeline.\n\n"
                "Pre-conditions:\n"
                "\u2022 Early Revision and Latest Revision must be set\n"
                "\u2022 Output file path must be set\n"
                "\u2022 Both input files must be valid .docx files\n"
                "\u2022 Word must be installed (unless Force XML is checked)")

        self.open_btn = ttk.Button(action_frame, text="Open Output",
                                    style="Open.TButton",
                                    command=self._open_output, state="disabled")
        self.open_btn.pack(side="left", padx=(10, 0))
        ToolTip(self.open_btn,
                "Open the output .docx in Microsoft Word\n"
                "(or your default .docx handler).\n\n"
                "Pre-condition: A comparison must have completed\n"
                "successfully first.")

        self.status_label = ttk.Label(action_frame, text="Ready",
                                       style="Status.TLabel")
        self.status_label.pack(side="right")

        self.progress = ttk.Progressbar(outer, mode="indeterminate",
                                         style="Accent.Horizontal.TProgressbar",
                                         length=200)
        self.progress.pack(fill="x", pady=(0, 10))

        # ── Log output ───────────────────────────────────────────────────
        log_frame = ttk.Frame(outer, style="App.TFrame")
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(
            log_frame,
            bg=self.SURFACE,
            fg=self.FG,
            font=("Consolas", 9),
            insertbackground=self.FG,
            selectbackground=self.ACCENT,
            selectforeground="#1e1e2e",
            relief="flat",
            borderwidth=0,
            padx=10,
            pady=8,
            wrap="word",
            state="disabled",
        )
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)

        # Text tags for coloured log lines
        self.log_text.tag_configure("info", foreground=self.FG)
        self.log_text.tag_configure("error", foreground=self.RED)
        self.log_text.tag_configure("warning", foreground=self.YELLOW)
        self.log_text.tag_configure("success", foreground=self.GREEN)
        self.log_text.tag_configure("heading", foreground=self.ACCENT,
                                     font=("Consolas", 9, "bold"))

        # Install the logging handler
        handler = TextWidgetHandler(self.log_text)
        handler.setFormatter(logging.Formatter("%(asctime)s  %(message)s",
                                                datefmt="%H:%M:%S"))
        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.DEBUG)

    # ----- Help dialog -----------------------------------------------------

    def _show_help(self):
        """Open a scrollable Help window explaining every part of the tool."""
        win = tk.Toplevel(self.root)
        win.title("Help - MS Word Redline Creator v1.0")
        win.configure(bg=self.BG)
        win.geometry("720x740")
        win.minsize(560, 480)
        win.transient(self.root)
        win.grab_set()

        # Scrollable text area
        frame = tk.Frame(win, bg=self.BG, padx=20, pady=16)
        frame.pack(fill="both", expand=True)

        txt = tk.Text(
            frame, bg=self.SURFACE, fg=self.FG, font=("Segoe UI", 10),
            relief="flat", wrap="word", padx=16, pady=14,
            insertbackground=self.FG, selectbackground=self.ACCENT,
            selectforeground="#1e1e2e", borderwidth=0, spacing1=2, spacing3=4,
        )
        sb = ttk.Scrollbar(frame, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)

        # Tags for formatting
        txt.tag_configure("h1", font=("Segoe UI", 16, "bold"), foreground=self.ACCENT,
                          spacing1=14, spacing3=6)
        txt.tag_configure("h2", font=("Segoe UI", 12, "bold"), foreground=self.ACCENT,
                          spacing1=12, spacing3=4)
        txt.tag_configure("h3", font=("Segoe UI", 10, "bold"), foreground=self.ACCENT_HOVER,
                          spacing1=8, spacing3=2)
        txt.tag_configure("body", font=("Segoe UI", 10), foreground=self.FG)
        txt.tag_configure("dim", font=("Segoe UI", 9), foreground=self.FG_DIM)
        txt.tag_configure("mono", font=("Consolas", 9), foreground=self.GREEN,
                          background="#262637")
        txt.tag_configure("bullet", font=("Segoe UI", 10), foreground=self.FG,
                          lmargin1=24, lmargin2=24)
        txt.tag_configure("warning", font=("Segoe UI", 9, "italic"),
                          foreground=self.YELLOW)

        def h1(text):  txt.insert("end", text + "\n", "h1")
        def h2(text):  txt.insert("end", text + "\n", "h2")
        def h3(text):  txt.insert("end", text + "\n", "h3")
        def p(text):   txt.insert("end", text + "\n\n", "body")
        def dim(text): txt.insert("end", text + "\n\n", "dim")
        def mono(text):txt.insert("end", text + "\n\n", "mono")
        def bullet(text): txt.insert("end", "  \u2022  " + text + "\n", "bullet")
        def warn(text):txt.insert("end", text + "\n\n", "warning")
        def gap():     txt.insert("end", "\n", "body")

        # ── Content ───────────────────────────────────────────────────
        h1("MS Word Redline Creator")
        dim("Version 1.0  \u2022  Author: Jack Wang")
        p("This tool compares two revisions of the same Microsoft Word document and "
          "produces a new .docx file that looks like the Latest Revision but includes "
          "tracked changes (redlines) showing every difference from the Early Revision. "
          "Optionally, it can also carry over all review comments from the Early "
          "Revision into the correct locations in the output.")

        h2("Document Selection")
        gap()
        h3("Early Revision")
        p("The older version of the document. This is the \"before\" snapshot. "
          "If \"Carry over comments\" is enabled, any review comments in this "
          "file will be extracted and re-anchored into the output. Must be a .docx file. "
          "You can browse for the file or drag & drop it onto the field.")

        h3("Latest Revision")
        p("The newer version of the document. This is the \"after\" snapshot and "
          "serves as the visual and formatting base for the output. The output "
          "document will use this file's fonts, styles, themes, headers/footers, "
          "images, and layout. Must be a .docx file. "
          "You can browse for the file or drag & drop it onto the field.")
        p("When this field is set, the Output File path is automatically "
          "generated as \"<LatestRevName>_redline.docx\" in the same folder.")

        h3("Output File")
        p("Where to save the result. Auto-populated when you set the Latest "
          "Revision \u2014 the default is \"<LatestRevName>_redline.docx\" in "
          "the same folder as the Latest Revision. You can override this by "
          "typing a path or using Save As. "
          "The output is a standard .docx file you can open in Microsoft Word.")

        h2("Options")
        gap()
        h3("Author")
        p("The name stamped on every tracked change in the output document. "
          "When you open the output in Word and hover over a redline "
          "(insertion or deletion), this is the author name that appears. "
          "Change it to your own name, a team name, or leave the default.")
        dim("Default: \"Revision Compare System\"")

        h3("Force XML (no Word COM)")
        p("Controls which comparison engine is used:")
        bullet("Unchecked (default) \u2014 Uses Microsoft Word's built-in comparison "
               "engine via COM automation. Requires Word to be installed. Produces "
               "the highest-fidelity tracked changes: word-level and character-level "
               "diffs, formatting change detection, and move tracking. This is "
               "equivalent to doing Review \u2192 Compare in Word itself.")
        bullet("Checked \u2014 Uses a pure-Python XML diff. Works on any machine "
               "without Word installed. Only paragraph-level granularity: cannot "
               "detect moves or formatting-only changes. Good enough for basic "
               "comparisons but not enterprise-grade.")
        gap()
        warn("Recommendation: Leave unchecked if you have Word installed. "
             "Only check this if Word is unavailable or COM is causing errors.")

        h3("Carry Over Comments from Early Revision")
        p("When checked (the default), comments from the Early Revision are "
          "extracted, mapped to the closest matching location in the Latest "
          "Revision, and inserted into the output document. This includes "
          "multi-strategy mapping, deduplication, and an Unmapped Comments "
          "appendix for any comments that couldn't be placed.")
        p("When unchecked, the comment pipeline (Steps 1, 3, and 4) is skipped "
          "entirely. The output will contain only tracked changes (redlines) "
          "with no review comments. This is faster and useful when you only "
          "need a redline comparison without comment carry-over.")
        dim("Default: Checked (comments are carried over)")

        h3("Verbose Logging")
        p("When checked, the log panel shows debug-level detail: every "
          "individual comment mapping decision, similarity scores, paragraph "
          "indices, deduplication events, and XML operations. "
          "When unchecked, only high-level step summaries are shown.")
        dim("Useful for troubleshooting when a comment maps to the wrong location.")

        h2("Buttons")
        gap()
        h3("Run Comparison")
        p("Starts the pipeline. The button is disabled while a comparison is "
          "running. Progress is shown in the progress bar and the log panel "
          "updates in real time. You can continue to resize the window but "
          "cannot start a second comparison until the first finishes.")

        h3("Open Output")
        p("Enabled after a successful comparison. Opens the output .docx "
          "directly in Microsoft Word (or your default .docx handler).")

        h2("What the Pipeline Does")
        p("The tool runs four steps in sequence:")
        bullet("Step 1 \u2014 Extract Comments: Reads all review comments from "
               "the Early Revision, including author, date, comment text, and the "
               "anchor text each comment was attached to.")
        bullet("Step 2 \u2014 Generate Tracked Changes: Compares the two documents "
               "and produces a new file with redline markup. Then transplants "
               "styles.xml, theme, and fontTable from the Latest Revision so "
               "the output's fonts match exactly.")
        bullet("Step 3 \u2014 Map Comments: For each extracted comment, finds the "
               "best matching location in the Latest Revision using four "
               "strategies in order: exact text match, fuzzy substring match, "
               "paragraph-level similarity, and heading proximity.")
        bullet("Step 4 \u2014 Insert Comments: Injects the mapped comments into "
               "the output document's XML with proper commentRangeStart/End "
               "markers. Comments that Word COM already carried forward are "
               "automatically deduplicated (skipped). Any comment that could "
               "not be mapped is preserved in an \"Unmapped Comments\" appendix "
               "at the end of the document.")

        h2("Comment Mapping Strategies")
        p("Comments are mapped in order of confidence. The first strategy that "
          "succeeds above its threshold is used:")
        bullet("Exact Match (confidence 0.95\u20131.0) \u2014 The original anchor "
               "text is found verbatim in the Latest Revision.")
        bullet("Fuzzy Substring (confidence 0.6\u20130.94) \u2014 A close textual "
               "match is found using sequence alignment (handles minor edits, "
               "rewordings, or whitespace changes).")
        bullet("Paragraph Match (confidence 0.4\u20130.6) \u2014 The full "
               "paragraph surrounding the original anchor is matched to the "
               "most similar paragraph in the Latest Revision. Style and "
               "table context give bonus score.")
        bullet("Heading Proximity (confidence varies) \u2014 The nearest heading "
               "before the original anchor is matched to a heading in the "
               "Latest Revision, and the comment is placed on the first "
               "paragraph after that heading.")
        bullet("Unmapped (confidence 0) \u2014 No reliable match found. "
               "The comment is preserved in the Unmapped Comments appendix "
               "with its original anchor text and reason for failure.")

        h2("Output Document Details")
        bullet("Format: Standard .docx, fully editable in Microsoft Word.")
        bullet("Tracked Changes: Native Word revision marks (not simulated). "
               "You can Accept/Reject changes normally in Word.")
        bullet("Comments: Native Word comments with original author and date. "
               "Threaded replies are preserved where possible.")
        bullet("Fonts & Styles: Copied from the Latest Revision (styles.xml, "
               "theme1.xml, fontTable.xml). The output should look identical "
               "to the Latest Revision when changes are accepted.")
        bullet("Unmapped Appendix: If any comments could not be placed, a "
               "new section titled \"Unmapped Comments from Early Revision\" "
               "is appended on a new page with full details.")

        h2("Command Line Usage")
        p("The tool can also be run from the command line without the GUI:")
        mono("python compare_revisions.py early.docx latest.docx output.docx")
        p("Additional flags:")
        mono("  --author \"Jane Doe\"    Author name for tracked changes\n"
             "  --force-xml            Use pure-XML comparison (no Word COM)\n"
             "  --skip-comments        Redlines only, no comment carry-over\n"
             "  --verbose              Enable debug-level logging\n"
             "  --log comparison.log   Write log to a file\n"
             "  --report report.json   Write a JSON pipeline report")

        h2("Deployment")
        p("Two deployment options are provided:")
        bullet("install.bat \u2014 Source install. Creates a virtual environment, "
               "installs dependencies, and optionally adds a desktop shortcut. "
               "Requires Python 3.10+ on the machine.")
        bullet("build.bat \u2014 Standalone build. Uses PyInstaller to create "
               "RedlineCreator.exe in a dist/ folder, then packages it as a "
               "ZIP file. The ZIP can be deployed to any Windows machine "
               "with no Python required.")
        gap()

        h2("Requirements")
        bullet("Microsoft Word \u2014 for highest-fidelity comparison (Word COM mode)")
        bullet("Python 3.10+ \u2014 only if running from source (not needed for .exe)")
        bullet("lxml \u2014 auto-installed by install.bat or bundled in .exe")
        bullet("pywin32 \u2014 auto-installed; only needed for Word COM mode")
        bullet("tkinterdnd2 \u2014 auto-installed; enables drag & drop in the GUI")
        gap()

        h2("Troubleshooting")
        h3("\"Word COM comparison failed\"")
        p("Word may not be installed, or another instance is blocking COM. "
          "Close all Word windows, try again, or check \"Force XML\" as a fallback.")
        h3("Comments mapped to wrong locations")
        p("Enable Verbose Logging and look for the confidence scores. Low scores "
          "(below 0.6) indicate the text changed substantially between revisions. "
          "The comment will still be preserved \u2014 either at the best-guess "
          "location or in the Unmapped Comments appendix.")
        h3("Output fonts don't match Latest Revision")
        p("The tool copies styles.xml, theme1.xml, and fontTable.xml from the "
          "Latest Revision. If fonts still differ, the Latest Revision may use "
          "fonts not installed on your machine, or Word's comparison engine "
          "may have introduced style conflicts. Open the output in Word and "
          "check Format \u2192 Styles to verify.")

        txt.configure(state="disabled")

        # Close button
        close_frame = tk.Frame(win, bg=self.BG, pady=10)
        close_frame.pack(fill="x")
        close_btn = tk.Button(
            close_frame, text="  Close  ", font=("Segoe UI", 10),
            bg=self.ACCENT, fg="#1e1e2e", activebackground=self.ACCENT_HOVER,
            activeforeground="#1e1e2e", relief="flat", cursor="hand2",
            padx=16, pady=4, command=win.destroy,
        )
        close_btn.pack()

    def _file_row(self, parent, label: str, var: tk.StringVar, row: int,
                  tooltip: str = ""):
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.pack(fill="x", pady=(0 if row == 0 else 6, 0))

        lbl = ttk.Label(frame, text=label, style="Card.TLabel", width=16, anchor="e")
        lbl.pack(side="left")

        entry = ttk.Entry(frame, textvariable=var, style="App.TEntry")
        entry.pack(side="left", fill="x", expand=True, padx=(6, 6))

        btn = ttk.Button(frame, text="Browse...", style="Browse.TButton",
                          command=lambda: self._browse_input(var))
        btn.pack(side="left")

        if tooltip:
            dnd_hint = tooltip + ("\n\nDrag & drop a .docx file here."
                                  if HAS_DND else "")
            ToolTip(lbl, dnd_hint)
            ToolTip(entry, dnd_hint)

        # Enable drag-and-drop on the entry
        self._enable_drop(entry, var)

    def _output_row(self, parent, label: str, var: tk.StringVar, row: int,
                    tooltip: str = ""):
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.pack(fill="x", pady=(6, 0))

        lbl = ttk.Label(frame, text=label, style="Card.TLabel", width=16, anchor="e")
        lbl.pack(side="left")

        entry = ttk.Entry(frame, textvariable=var, style="App.TEntry")
        entry.pack(side="left", fill="x", expand=True, padx=(6, 6))

        btn = ttk.Button(frame, text="Save As...", style="Browse.TButton",
                          command=lambda: self._browse_output(var))
        btn.pack(side="left")

        if tooltip:
            dnd_hint = tooltip + ("\n\nDrag & drop a .docx file here."
                                  if HAS_DND else "")
            ToolTip(lbl, dnd_hint)
            ToolTip(entry, dnd_hint)

        # Enable drag-and-drop on the entry
        self._enable_drop(entry, var)

    # ----- Drag-and-drop --------------------------------------------------

    def _enable_drop(self, widget: tk.Widget, var: tk.StringVar):
        """Register a widget as a drag-and-drop target (if tkinterdnd2 is available)."""
        if not HAS_DND:
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", lambda e: self._on_drop(e, var))
            widget.dnd_bind("<<DragEnter>>",
                            lambda e: widget.configure(
                                style="DropHighlight.TEntry") if isinstance(
                                widget, ttk.Entry) else None)
            widget.dnd_bind("<<DragLeave>>",
                            lambda e: widget.configure(
                                style="App.TEntry") if isinstance(
                                widget, ttk.Entry) else None)
        except Exception:
            pass  # silently skip if DnD registration fails

    def _on_drop(self, event, var: tk.StringVar):
        """Handle a file drop event."""
        raw = event.data
        # tkdnd may wrap paths in braces or return multiple files
        path = self._parse_drop_data(raw)
        if path:
            var.set(path)
        # Reset entry style after drop
        try:
            event.widget.configure(style="App.TEntry")
        except Exception:
            pass

    @staticmethod
    def _parse_drop_data(data: str) -> str:
        """Extract the first valid .docx path from drop data.

        tkdnd on Windows wraps paths containing spaces in {braces} and
        may deliver multiple space-separated paths.
        """
        data = data.strip()
        if not data:
            return ""

        # If wrapped in braces, extract the content
        if data.startswith("{"):
            end = data.index("}") if "}" in data else len(data)
            candidate = data[1:end]
        else:
            # Take the first token (space-separated)
            candidate = data.split()[0] if data else ""

        candidate = candidate.strip()

        # Accept .docx files (case-insensitive)
        if candidate.lower().endswith(".docx") and Path(candidate).exists():
            return candidate

        # If not .docx, still accept if it's a valid file
        if Path(candidate).is_file():
            return candidate

        return ""

    # ----- File dialogs ---------------------------------------------------

    def _browse_input(self, var: tk.StringVar):
        path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if path:
            var.set(path)

    def _browse_output(self, var: tk.StringVar):
        path = filedialog.asksaveasfilename(
            title="Save Output Document",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if path:
            var.set(path)
            self._output_was_auto = False  # user explicitly chose a path

    # ----- Auto-fill output -----------------------------------------------

    def _on_latest_rev_changed(self, *_args):
        """Called whenever latest_rev_var changes (browse, drop, or manual edit)."""
        self._auto_fill_output()

    def _auto_fill_output(self):
        """Auto-generate output path based on the Latest Revision.

        Fills in ``<latest_folder>/<latest_stem>_redline.docx``.
        Only overwrites the output field if it is empty or was previously
        auto-generated (so manual user edits are preserved).
        """
        current_output = self.output_var.get().strip()
        if current_output and not self._output_was_auto:
            return

        latest = self.latest_rev_var.get().strip()
        if not latest:
            return

        p = Path(latest)
        if not p.suffix.lower() == ".docx":
            return

        auto_path = str(p.parent / f"{p.stem}_redline{p.suffix}")
        self.output_var.set(auto_path)
        self._output_was_auto = True

    # ----- Run comparison -------------------------------------------------

    def _on_run(self):
        # Validate
        early = self.early_rev_var.get().strip()
        latest = self.latest_rev_var.get().strip()
        output = self.output_var.get().strip()

        errors = []
        if not early:
            errors.append("Early Revision is required.")
        elif not Path(early).exists():
            errors.append(f"Early Revision not found:\n{early}")
        if not latest:
            errors.append("Latest Revision is required.")
        elif not Path(latest).exists():
            errors.append(f"Latest Revision not found:\n{latest}")
        if not output:
            errors.append("Output file path is required.")

        if errors:
            messagebox.showerror("Validation Error", "\n\n".join(errors))
            return

        # Clear log
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

        # Disable controls
        self.running = True
        self.run_btn.configure(state="disabled")
        self.open_btn.configure(state="disabled")
        self.status_label.configure(text="Running...")
        self.progress.start(12)

        # Run in background thread
        thread = threading.Thread(target=self._run_pipeline, daemon=True)
        thread.start()

    def _run_pipeline(self):
        """Execute the comparison pipeline (called in a worker thread)."""
        from compare_revisions import run_comparison

        try:
            results = run_comparison(
                early_rev_path=self.early_rev_var.get().strip(),
                latest_rev_path=self.latest_rev_var.get().strip(),
                output_path=self.output_var.get().strip(),
                author=self.author_var.get().strip() or DEFAULT_AUTHOR,
                force_xml=self.force_xml_var.get(),
                skip_comments=not self.carry_comments_var.get(),
                verbose=self.verbose_var.get(),
            )
            self.root.after(0, self._on_complete, results)
        except Exception as e:
            logging.getLogger(__name__).error(f"Unhandled error: {e}")
            self.root.after(0, self._on_complete, {
                "success": False,
                "errors": [str(e)],
                "comments": {},
            })

    def _on_complete(self, results: dict):
        """Handle pipeline completion (called on main thread)."""
        self.running = False
        self.progress.stop()
        self.run_btn.configure(state="normal")

        if results.get("success"):
            self.status_label.configure(text="Complete", foreground=self.GREEN)
            self.open_btn.configure(state="normal")

            # Summary in log
            self._log_heading("\n--- SUMMARY ---")
            c = results.get("comments", {})
            total = c.get("total_extracted", 0)
            unmapped = c.get("unmapped", 0)
            if total:
                self._log_info(
                    f"Comments: {total} extracted, "
                    f"{total - unmapped} mapped, {unmapped} unmapped"
                )
            dur = results.get("duration", 0)
            self._log_info(f"Duration: {dur:.1f}s")
            self._log_info(f"Output: {results.get('output', '')}")
        else:
            self.status_label.configure(text="Failed", foreground=self.RED)
            for err in results.get("errors", ["Unknown error"]):
                self._log_error(err)

    # ----- Open output ----------------------------------------------------

    def _open_output(self):
        output = self.output_var.get().strip()
        if not output or not Path(output).exists():
            messagebox.showwarning("Not Found", "Output file does not exist yet.")
            return
        os.startfile(output)

    # ----- Log helpers ----------------------------------------------------

    def _log_msg(self, msg: str, tag: str):
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, msg + "\n", tag)
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _log_info(self, msg: str):
        self._log_msg(msg, "info")

    def _log_error(self, msg: str):
        self._log_msg(msg, "error")

    def _log_heading(self, msg: str):
        self._log_msg(msg, "heading")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = RevisionCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
