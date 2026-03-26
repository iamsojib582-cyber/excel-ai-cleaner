"""
╔══════════════════════════════════════════════════════════════╗
║            EXCEL AI CLEANER — by Sajeeb The Analyst         ║
║                        utils.py                             ║
║     Config manager · PDF report · Helper functions          ║
╚══════════════════════════════════════════════════════════════╝

Features:
  • Persistent config (API key, window size, last folder, theme)
  • PDF report generator — problems found + before/after stats
  • Branded with Sajeeb The Analyst
  • Helper functions used across all modules
"""

import json
import os
import re
from datetime import datetime
from typing import Any, Optional

import pandas as pd


# ════════════════════════════════════════════════════════════
#  PATHS
# ════════════════════════════════════════════════════════════

# Root folder = folder that contains this utils.py file
ROOT_DIR    = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR  = os.path.join(ROOT_DIR, "assets")
CONFIG_PATH = os.path.join(ASSETS_DIR, "config.json")

os.makedirs(ASSETS_DIR, exist_ok=True)


# ════════════════════════════════════════════════════════════
#  DEFAULT CONFIG
# ════════════════════════════════════════════════════════════

DEFAULT_CONFIG: dict = {
    "api_key"        : "",
    "window_width"   : 1400,
    "window_height"  : 820,
    "window_x"       : -1,          # -1 = center on screen
    "window_y"       : -1,
    "last_folder"    : "",
    "theme"          : "deep_purple_gold",
    "font_size"      : 10,
    "show_splash"    : True,
    "max_preview_rows": 500,
    "developer_name" : "Sajeeb The Analyst",
    "app_version"    : "1.0.0",
}


# ════════════════════════════════════════════════════════════
#  CONFIG MANAGER
# ════════════════════════════════════════════════════════════

class ConfigManager:
    """
    Loads and saves app configuration to assets/config.json.
    Thread-safe reads. Merges with defaults so new keys
    are never missing even on old config files.
    """

    def __init__(self):
        self._data: dict = {}
        self.load()

    # ── load ─────────────────────────────────────────────────
    def load(self):
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                # Merge: start from defaults, overlay saved values
                self._data = {**DEFAULT_CONFIG, **saved}
            except Exception:
                self._data = DEFAULT_CONFIG.copy()
        else:
            self._data = DEFAULT_CONFIG.copy()

    # ── save ─────────────────────────────────────────────────
    def save(self):
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(self._data, f, indent=2)
        except Exception as e:
            print(f"[Config] Save failed: {e}")

    # ── get / set ─────────────────────────────────────────────
    def get(self, key: str, fallback: Any = None) -> Any:
        return self._data.get(key, fallback)

    def set(self, key: str, value: Any):
        self._data[key] = value
        self.save()

    def set_many(self, updates: dict):
        self._data.update(updates)
        self.save()

    # ── convenience properties ───────────────────────────────
    @property
    def api_key(self) -> str:
        return self._data.get("api_key", "")

    @api_key.setter
    def api_key(self, value: str):
        self.set("api_key", value.strip())

    @property
    def last_folder(self) -> str:
        return self._data.get("last_folder", "")

    @last_folder.setter
    def last_folder(self, path: str):
        self.set("last_folder", path)

    @property
    def window_geometry(self) -> tuple[int, int, int, int]:
        return (
            self._data.get("window_width",  1400),
            self._data.get("window_height", 820),
            self._data.get("window_x", -1),
            self._data.get("window_y", -1),
        )

    def save_window_geometry(self, w: int, h: int, x: int, y: int):
        self.set_many({
            "window_width" : w,
            "window_height": h,
            "window_x"     : x,
            "window_y"     : y,
        })

    def __repr__(self):
        safe = {k: ("***" if k == "api_key" else v)
                for k, v in self._data.items()}
        return f"ConfigManager({safe})"


# ════════════════════════════════════════════════════════════
#  PDF REPORT GENERATOR
# ════════════════════════════════════════════════════════════

class PDFReport:
    """
    Generates a professional branded PDF report using only
    the standard library + reportlab (if available) or
    falls back to a plain .txt report if reportlab is missing.

    Report contains:
      • Header — app name, developer, date/time
      • File info section
      • Before vs After statistics table
      • Full list of AI-found problems
      • Footer
    """

    # ── Brand colors (as 0-1 RGB for reportlab) ──────────────
    PURPLE = (0.055, 0.043, 0.102)   # #0e0b1a
    GOLD   = (0.788, 0.659, 0.298)   # #c9a84c
    WHITE  = (0.941, 0.918, 1.000)   # #f0eaff
    GREY   = (0.478, 0.435, 0.604)   # #7a6f9a

    def __init__(self, config: Optional["ConfigManager"] = None):
        self.config = config or ConfigManager()

    # ── public entry point ───────────────────────────────────
    def generate(
        self,
        output_path  : str,
        file_name    : str,
        before_stats : Any,          # DataStats dataclass from cleaner.py
        after_stats  : Any,
        ai_issues    : list[str],    # list of issue strings
        diff         : dict,         # diff dict from cleaner.get_comparison()
    ) -> str:
        """
        Generate the report. Returns the output path on success.
        Tries PDF first, falls back to TXT.
        """
        try:
            return self._generate_pdf(
                output_path, file_name,
                before_stats, after_stats,
                ai_issues, diff,
            )
        except ImportError:
            # reportlab not installed — write plain text
            txt_path = output_path.replace(".pdf", ".txt")
            return self._generate_txt(
                txt_path, file_name,
                before_stats, after_stats,
                ai_issues, diff,
            )
        except Exception as e:
            raise RuntimeError(f"Report generation failed: {e}") from e

    # ── PDF generator ────────────────────────────────────────
    def _generate_pdf(
        self, path, file_name,
        before, after, issues, diff,
    ) -> str:
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles    import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units     import cm
            from reportlab.lib           import colors
            from reportlab.platypus      import (
                SimpleDocTemplate, Paragraph, Spacer, Table,
                TableStyle, HRFlowable,
            )
        except ImportError:
            raise ImportError("reportlab not installed - run 'pip install reportlab'")

        doc   = SimpleDocTemplate(
            path,
            pagesize     = A4,
            leftMargin   = 2*cm,
            rightMargin  = 2*cm,
            topMargin    = 2*cm,
            bottomMargin = 2*cm,
        )
        story = []
        W     = A4[0] - 4*cm     # usable width

        # ── color objects
        c_purple = colors.Color(*self.PURPLE)
        c_gold   = colors.Color(*self.GOLD)
        c_white  = colors.Color(*self.WHITE)
        c_grey   = colors.Color(*self.GREY)

        # ── styles
        styles = getSampleStyleSheet()

        s_title = ParagraphStyle(
            "Title",
            fontSize  = 22,
            textColor = c_gold,
            fontName  = "Helvetica-Bold",
            spaceAfter= 4,
        )
        s_subtitle = ParagraphStyle(
            "Subtitle",
            fontSize  = 11,
            textColor = c_grey,
            fontName  = "Helvetica",
            spaceAfter= 2,
        )
        s_section = ParagraphStyle(
            "Section",
            fontSize  = 13,
            textColor = c_gold,
            fontName  = "Helvetica-Bold",
            spaceBefore=14,
            spaceAfter = 4,
        )
        s_body = ParagraphStyle(
            "Body",
            fontSize  = 10,
            textColor = colors.black,
            fontName  = "Helvetica",
            spaceAfter= 3,
            leading   = 14,
        )
        s_issue = ParagraphStyle(
            "Issue",
            fontSize  = 9,
            textColor = colors.HexColor("#333333"),
            fontName  = "Helvetica",
            spaceAfter= 4,
            leftIndent= 10,
            leading   = 13,
        )

        # ══ HEADER BANNER ════════════════════════════════════
        story.append(Paragraph("Excel AI Cleaner", s_title))
        story.append(Paragraph(
            "Data Quality Report  •  Professional Edition", s_subtitle))
        story.append(HRFlowable(
            width=W, thickness=1.5,
            color=c_gold, spaceAfter=6))

        dev   = self.config.get("developer_name", "Sajeeb The Analyst")
        ver   = self.config.get("app_version", "1.0.0")
        now   = datetime.now().strftime("%Y-%m-%d  %H:%M:%S")

        story.append(Paragraph(
            f"<b>Developer:</b>  {dev}  &nbsp;|&nbsp; "
            f"<b>Version:</b> {ver}  &nbsp;|&nbsp; "
            f"<b>Generated:</b> {now}",
            s_body,
        ))
        story.append(Spacer(1, 0.3*cm))

        # ══ FILE INFO ════════════════════════════════════════
        story.append(Paragraph("File Information", s_section))
        story.append(Paragraph(f"<b>File:</b>  {file_name}", s_body))

        # ══ BEFORE vs AFTER TABLE ════════════════════════════
        story.append(Paragraph("Before vs After Cleaning", s_section))

        tbl_data = [
            ["Metric",       "Before Cleaning",          "After Cleaning",           "Improvement"],
            ["Total Rows",
             f"{before.rows:,}",
             f"{after.rows:,}",
             _delta(before.rows, after.rows, lower_better=True)],
            ["Columns",
             str(before.columns),
             str(after.columns),
             "—"],
            ["Missing Cells",
             f"{before.missing:,}",
             f"{after.missing:,}",
             _delta(before.missing, after.missing, lower_better=True)],
            ["Duplicate Rows",
             f"{before.duplicates:,}",
             f"{after.duplicates:,}",
             _delta(before.duplicates, after.duplicates, lower_better=True)],
            ["Numeric Columns",
             str(before.numeric_cols),
             str(after.numeric_cols),
             _delta(after.numeric_cols, before.numeric_cols, lower_better=False)],
            ["Date Columns",
             str(before.date_cols),
             str(after.date_cols),
             _delta(after.date_cols, before.date_cols, lower_better=False)],
        ]

        tbl = Table(tbl_data, colWidths=[W*0.32, W*0.22, W*0.22, W*0.24])
        tbl.setStyle(TableStyle([
            # Header row
            ("BACKGROUND",  (0,0), (-1,0), c_gold),
            ("TEXTCOLOR",   (0,0), (-1,0), colors.black),
            ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",    (0,0), (-1,0), 10),
            ("ALIGN",       (0,0), (-1,0), "CENTER"),
            ("BOTTOMPADDING",(0,0),(-1,0), 6),
            ("TOPPADDING",  (0,0), (-1,0), 6),
            # Body rows
            ("FONTSIZE",    (0,1), (-1,-1), 9),
            ("ALIGN",       (1,1), (-1,-1), "CENTER"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),
             [colors.HexColor("#f9f6ff"), colors.white]),
            ("GRID",        (0,0), (-1,-1), 0.4, colors.HexColor("#cccccc")),
            ("TOPPADDING",  (0,1), (-1,-1), 4),
            ("BOTTOMPADDING",(0,1),(-1,-1), 4),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 0.3*cm))

        # ══ AI ISSUES LIST ═══════════════════════════════════
        if issues:
            story.append(Paragraph(
                f"AI-Detected Problems  ({len(issues)} found)", s_section))
            for i, issue in enumerate(issues, 1):
                # Clean up any emoji that reportlab can't render
                clean = _strip_non_latin(issue)
                story.append(Paragraph(f"{i}.  {clean}", s_issue))
        else:
            story.append(Paragraph("AI-Detected Problems", s_section))
            story.append(Paragraph(
                "No AI scan was run or no issues were found.", s_body))

        # ══ FOOTER ═══════════════════════════════════════════
        story.append(Spacer(1, 0.5*cm))
        story.append(HRFlowable(
            width=W, thickness=0.8,
            color=c_gold, spaceAfter=4))
        story.append(Paragraph(
            f"Generated by Excel AI Cleaner v{ver}  •  "
            f"Developed by <b>{dev}</b>  •  Powered by Groq AI (Llama-3)",
            s_subtitle,
        ))

        doc.build(story)
        return path

    # ── TXT fallback ─────────────────────────────────────────
    def _generate_txt(
        self, path, file_name,
        before, after, issues, diff,
    ) -> str:
        dev = self.config.get("developer_name", "Sajeeb The Analyst")
        ver = self.config.get("app_version", "1.0.0")
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        w   = 62

        lines = [
            "=" * w,
            "  EXCEL AI CLEANER — DATA QUALITY REPORT",
            f"  Developed by: {dev}",
            f"  Version: {ver}",
            f"  Generated: {now}",
            "=" * w,
            "",
            f"FILE: {file_name}",
            "",
            "─" * w,
            "BEFORE vs AFTER CLEANING",
            "─" * w,
            f"{'Metric':<22} {'Before':>10} {'After':>10} {'Change':>10}",
            f"{'─'*22} {'─'*10} {'─'*10} {'─'*10}",
            f"{'Total Rows':<22} {before.rows:>10,} {after.rows:>10,} "
            f"{_delta(before.rows, after.rows, True):>10}",
            f"{'Missing Cells':<22} {before.missing:>10,} {after.missing:>10,} "
            f"{_delta(before.missing, after.missing, True):>10}",
            f"{'Duplicate Rows':<22} {before.duplicates:>10,} {after.duplicates:>10,} "
            f"{_delta(before.duplicates, after.duplicates, True):>10}",
            f"{'Numeric Columns':<22} {before.numeric_cols:>10} {after.numeric_cols:>10}",
            f"{'Date Columns':<22} {before.date_cols:>10} {after.date_cols:>10}",
            "",
            "─" * w,
            f"AI-DETECTED PROBLEMS ({len(issues)} found)",
            "─" * w,
        ]

        for i, issue in enumerate(issues, 1):
            lines.append(f"{i:>3}. {issue}")

        lines += [
            "",
            "=" * w,
            f"Generated by Excel AI Cleaner v{ver}",
            f"Developer: {dev}",
            f"Powered by Groq AI (Llama-3)",
            "=" * w,
        ]

        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return path


# ════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ════════════════════════════════════════════════════════════

def _delta(before: int, after: int, lower_better: bool) -> str:
    """Format a numeric delta for the report table."""
    diff = after - before
    if diff == 0:
        return "No change"
    sign   = "+" if diff > 0 else ""
    better = (diff < 0 and lower_better) or (diff > 0 and not lower_better)
    tag    = "✓" if better else "✗"
    return f"{tag} {sign}{diff:,}"


def _strip_non_latin(text: str) -> str:
    """Remove emoji and non-latin characters for reportlab compatibility."""
    return re.sub(r"[^\x00-\x7F\u00C0-\u024F]", "", text).strip()


def format_number(n: Any) -> str:
    """Format a number for display — handles int, float, NA."""
    if pd.isna(n):
        return ""
    if isinstance(n, float):
        return f"{n:,.4g}"
    if isinstance(n, int):
        return f"{n:,}"
    return str(n)


def truncate(text: str, max_len: int = 30) -> str:
    """Truncate long strings for display in table cells."""
    text = str(text)
    return text if len(text) <= max_len else text[:max_len - 1] + "…"


def safe_filename(name: str) -> str:
    """Strip illegal characters from a filename."""
    return re.sub(r'[<>:"/\\|?*]', "_", name)


def timestamp_str() -> str:
    """Return current timestamp as a safe filename string."""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def file_size_str(path: str) -> str:
    """Return human-readable file size."""
    try:
        size = os.path.getsize(path)
        if size < 1024:
            return f"{size} B"
        elif size < 1024 ** 2:
            return f"{size/1024:.1f} KB"
        elif size < 1024 ** 3:
            return f"{size/1024**2:.1f} MB"
        else:
            return f"{size/1024**3:.1f} GB"
    except Exception:
        return "unknown size"


def center_window(root, width: int, height: int):
    """Center a tkinter window on screen."""
    root.update_idletasks()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x  = (sw - width)  // 2
    y  = (sh - height) // 2
    root.geometry(f"{width}x{height}+{x}+{y}")


def apply_theme_to_ttk(style, theme: dict):
    """
    Apply the Deep Purple & Gold theme to ttk widgets.
    Call once after creating the main window.
    """
    bg     = theme.get("bg",           "#0e0b1a")
    panel  = theme.get("panel",        "#1a1530")
    gold   = theme.get("gold",         "#c9a84c")
    bright = theme.get("gold_bright",  "#e8c96d")
    text   = theme.get("text",         "#f0eaff")
    dim    = theme.get("dim",          "#7a6f9a")
    border = theme.get("border",       "#2e2550")
    select = theme.get("select",       "#2a1f5e")

    style.theme_use("clam")

    style.configure(".",
        background        = bg,
        foreground        = text,
        fieldbackground   = panel,
        selectbackground  = select,
        selectforeground  = text,
        bordercolor       = border,
        darkcolor         = bg,
        lightcolor        = panel,
        troughcolor       = border,
        font              = ("Segoe UI", 10),
    )

    # ── TFrame
    style.configure("TFrame",       background=bg)
    style.configure("Card.TFrame",  background=panel,
                    relief="flat",  borderwidth=1)

    # ── TLabel
    style.configure("TLabel",       background=bg,   foreground=text)
    style.configure("Dim.TLabel",   background=bg,   foreground=dim)
    style.configure("Gold.TLabel",  background=bg,   foreground=gold,
                    font=("Segoe UI", 10, "bold"))
    style.configure("Title.TLabel", background=bg,   foreground=bright,
                    font=("Segoe UI", 14, "bold"))

    # ── TButton
    style.configure("TButton",
        background  = panel,
        foreground  = gold,
        bordercolor = gold,
        focuscolor  = gold,
        relief      = "flat",
        padding     = (10, 5),
        font        = ("Segoe UI", 9, "bold"),
    )
    style.map("TButton",
        background   = [("active", select),   ("pressed", border)],
        foreground   = [("active", bright),   ("pressed", bright)],
        bordercolor  = [("active", bright)],
    )

    # ── Gold primary button
    style.configure("Gold.TButton",
        background  = gold,
        foreground  = "#0e0b1a",
        bordercolor = gold,
        relief      = "flat",
        padding     = (12, 6),
        font        = ("Segoe UI", 9, "bold"),
    )
    style.map("Gold.TButton",
        background  = [("active", bright), ("pressed", "#b8952a")],
        foreground  = [("active", "#0e0b1a")],
    )

    # ── TNotebook
    style.configure("TNotebook",
        background  = bg,
        bordercolor = border,
        tabmargins  = [4, 4, 0, 0],
    )
    style.configure("TNotebook.Tab",
        background  = panel,
        foreground  = dim,
        padding     = [14, 6],
        font        = ("Segoe UI", 9),
    )
    style.map("TNotebook.Tab",
        background  = [("selected", bg)],
        foreground  = [("selected", gold)],
    )

    # ── Treeview
    style.configure("Treeview",
        background       = panel,
        foreground       = text,
        fieldbackground  = panel,
        rowheight        = 26,
        bordercolor      = border,
        font             = ("Segoe UI", 9),
    )
    style.configure("Treeview.Heading",
        background  = border,
        foreground  = gold,
        relief      = "flat",
        font        = ("Segoe UI", 9, "bold"),
        padding     = [6, 5],
    )
    style.map("Treeview",
        background  = [("selected", select)],
        foreground  = [("selected", bright)],
    )

    # ── TEntry
    style.configure("TEntry",
        fieldbackground = panel,
        foreground      = text,
        insertcolor     = gold,
        bordercolor     = border,
        relief          = "flat",
        padding         = 5,
    )
    style.map("TEntry",
        bordercolor = [("focus", gold)],
    )

    # ── TCombobox
    style.configure("TCombobox",
        fieldbackground = panel,
        foreground      = text,
        selectbackground= select,
        arrowcolor      = gold,
        bordercolor     = border,
        padding         = 5,
    )

    # ── Progressbar
    style.configure("TProgressbar",
        troughcolor = border,
        background  = gold,
        thickness   = 6,
    )

    # ── TScrollbar
    style.configure("TScrollbar",
        background  = panel,
        troughcolor = bg,
        arrowcolor  = dim,
        bordercolor = bg,
        relief      = "flat",
    )
    style.map("TScrollbar",
        background  = [("active", border)],
    )

    # ── TSeparator
    style.configure("TSeparator",
        background  = border,
    )


# ════════════════════════════════════════════════════════════
#  THEME DICTIONARY  (imported by ui.py)
# ════════════════════════════════════════════════════════════

THEME = {
    "bg"          : "#0e0b1a",
    "panel"       : "#1a1530",
    "gold"        : "#c9a84c",
    "gold_bright" : "#e8c96d",
    "text"        : "#f0eaff",
    "dim"         : "#7a6f9a",
    "border"      : "#2e2550",
    "select"      : "#2a1f5e",
    "success"     : "#4caf80",
    "error"       : "#ff6b6b",
    "warning"     : "#f0a84c",
    "gold_cell"   : "#3a2e10",   # background for modified cells
}


# ════════════════════════════════════════════════════════════
#  SINGLETON CONFIG INSTANCE
# ════════════════════════════════════════════════════════════

_config_instance: Optional[ConfigManager] = None


def get_config() -> ConfigManager:
    """Return the shared ConfigManager instance."""
    global _config_instance
    if _config_instance is None:
        _config_instance = ConfigManager()
    return _config_instance