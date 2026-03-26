"""
╔══════════════════════════════════════════════════════════════════╗
║         EXCEL AI CLEANER  —  by Sajeeb The Analyst              ║
║         Deep Purple & Gold  •  Advanced  •  Single File         ║
╠══════════════════════════════════════════════════════════════════╣
║  SETUP (run once in VS Code terminal):                          ║
║      pip install groq pandas openpyxl matplotlib numpy          ║
║                                                                 ║
║  FREE GROQ API KEY:                                             ║
║      1. Go to  https://console.groq.com                        ║
║      2. Sign up with Google (free)                             ║
║      3. Click API Keys → Create API Key                        ║
║      4. Paste it inside the app → Settings tab                 ║
║                                                                 ║
║  RUN:   python excel_ai_cleaner.py                             ║
╚══════════════════════════════════════════════════════════════════╝
"""

import json
import os
import re
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, simpledialog, ttk

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


# ══════════════════════════════════════════════════════════════════
#  THEME  ─  Deep Purple & Gold
# ══════════════════════════════════════════════════════════════════

T = {
    "bg"         : "#0d0d0f",
    "panel"      : "#16161a",
    "panel2"     : "#1f1f23",
    "gold"       : "#ffb703",
    "gold_bright": "#ffcc33",
    "gold_dim"   : "#b38b00",
    "text"       : "#e0e0e6",
    "dim"        : "#94949e",
    "border"     : "#2d2d34",
    "select"     : "#32323d",
    "success"    : "#2ec4b6",
    "error"      : "#e63946",
    "warning"    : "#f77f00",
    "gold_cell"  : "#1a1608",
}


def apply_theme(style: ttk.Style) -> None:
    bg, panel, gold, bright, text, dim, border, sel = (
        T["bg"], T["panel"], T["gold"], T["gold_bright"],
        T["text"], T["dim"], T["border"], T["select"],
    )
    style.theme_use("clam")
    style.configure(".",
        background=bg, foreground=text, fieldbackground=panel,
        selectbackground=sel, selectforeground=text,
        bordercolor=border, troughcolor=border,
        font=("Segoe UI", 10),
    )
    style.configure("TFrame",    background=bg)
    style.configure("TLabel",    background=bg, foreground=text)
    style.configure("Dim.TLabel",  background=bg, foreground=dim,  font=("Segoe UI", 9))
    style.configure("Gold.TLabel", background=bg, foreground=gold, font=("Segoe UI", 10, "bold"))

    style.configure("TButton",
        background=panel, foreground=gold, bordercolor=gold,
        focuscolor=gold, relief="flat", padding=(10, 5),
        font=("Segoe UI", 9, "bold"),
    )
    style.map("TButton",
        background=[("active", sel),    ("pressed", border)],
        foreground=[("active", bright), ("pressed", bright)],
    )
    style.configure("Gold.TButton",
        background=gold, foreground="#0e0b1a", bordercolor=gold,
        relief="flat", padding=(12, 6), font=("Segoe UI", 9, "bold"),
    )
    style.map("Gold.TButton",
        background=[("active", bright), ("pressed", T["gold_dim"])],
        foreground=[("active", "#0e0b1a")],
    )
    style.configure("TNotebook", background=bg, bordercolor=border)
    style.configure("TNotebook.Tab",
        background=panel, foreground=dim, padding=[14, 7], font=("Segoe UI", 9),
    )
    style.map("TNotebook.Tab",
        background=[("selected", bg)], foreground=[("selected", gold)],
    )
    style.configure("Treeview",
        background=panel, foreground=text, fieldbackground=panel,
        rowheight=26, font=("Segoe UI", 9),
    )
    style.configure("Treeview.Heading",
        background=border, foreground=gold, relief="flat",
        font=("Segoe UI", 9, "bold"), padding=[6, 5],
    )
    style.map("Treeview",
        background=[("selected", sel)], foreground=[("selected", bright)],
    )
    style.configure("TEntry",
        fieldbackground=panel, foreground=text, insertcolor=gold,
        bordercolor=border, relief="flat", padding=5,
    )
    style.map("TEntry", bordercolor=[("focus", gold)])
    style.configure("TCombobox",
        fieldbackground=panel, foreground=text, selectbackground=sel,
        arrowcolor=gold, bordercolor=border, padding=5,
    )
    style.configure("TProgressbar",
        troughcolor=border, background=gold, thickness=6,
    )
    style.configure("TScrollbar",
        background=panel, troughcolor=bg, arrowcolor=dim,
        bordercolor=bg, relief="flat",
    )
    style.map("TScrollbar", background=[("active", border)])
    style.configure("TSeparator", background=border)


# ══════════════════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════════════════

_CFG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "excel_ai_config.json")
_CFG_DEF  = {"api_key": "", "last_folder": "", "window_w": 1420, "window_h": 840}


def cfg_load() -> dict:
    try:
        with open(_CFG_PATH) as f:
            return {**_CFG_DEF, **json.load(f)}
    except Exception:
        return _CFG_DEF.copy()


def cfg_save(d: dict) -> None:
    try:
        with open(_CFG_PATH, "w") as f:
            json.dump(d, f, indent=2)
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════
#  FILE LOADER
# ══════════════════════════════════════════════════════════════════

_NULLS = [
    "NULL", "null", "N/A", "NA", "n/a", "na", "None", "none",
    "NaN", "nan", "#N/A", "#n/a", "#NA", "?", "-", "",
    "INF", "inf", "-inf", "NIL", "nil", "NONE",
]


def _sep(text: str) -> str:
    s = "\n".join(text.splitlines()[:20])
    return max([";", ",", "\t", "|"], key=s.count)


def _hdr(raw: list) -> list:
    seen, out = {}, []
    for i, h in enumerate(raw):
        h = re.sub(r"[^\x00-\x7F]", "", str(h)).strip().replace(" ", "_") or f"Col_{i}"
        if h in seen:
            seen[h] += 1
            out.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            out.append(h)
    return out


def _norm(df: pd.DataFrame) -> pd.DataFrame:
    df = df.replace(_NULLS, pd.NA)
    for c in df.select_dtypes(include=["object"]).columns:
        s = df[c].dropna().astype(str).str.strip()
        if s.empty:
            continue
        num = pd.to_numeric(s.str.replace(r"[$,\s%€£¥]", "", regex=True), errors="coerce")
        if num.notna().mean() > 0.82:
            df[c] = pd.to_numeric(
                df[c].astype(str).str.replace(r"[$,\s%€£¥]", "", regex=True),
                errors="coerce",
            )
    return df


def load_file(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()

    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(path)
        if df.shape[1] <= 3 and df.iloc[:, 0].astype(str).str.contains(";").mean() > 0.5:
            rows = [str(r).split(";") for r in df.iloc[:, 0]]
            w    = max(len(r) for r in rows)
            rows = [(r + [""] * (w - len(r)))[:w] for r in rows]
            df   = pd.DataFrame(rows[1:], columns=_hdr(rows[0]))
        return _norm(df)

    if ext == ".json":
        with open(path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        return _norm(pd.json_normalize(raw if isinstance(raw, list) else [raw]))

    # CSV
    raw = None
    for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1252"]:
        try:
            with open(path, "r", encoding=enc, errors="replace") as f:
                raw = f.read()
            break
        except Exception:
            continue
    if not raw:
        raise ValueError("Cannot read this file.")

    sep   = _sep(raw)
    lines = [l.strip() for l in raw.splitlines()
             if l.strip() and not re.fullmatch(r"[;,|\t\s]+", l.strip())]
    if not lines:
        raise ValueError("File is empty.")

    rows  = [l.split(sep) for l in lines]
    width = max(set(len(r) for r in rows), key=[len(r) for r in rows].count)
    rows  = [(r + [""] * (width - len(r)))[:width] for r in rows]
    return _norm(pd.DataFrame(rows[1:], columns=_hdr(rows[0])))


# ══════════════════════════════════════════════════════════════════
#  UNDO / REDO STACK
# ══════════════════════════════════════════════════════════════════

class UndoStack:
    MAX = 40

    def __init__(self):
        self._u: list = []
        self._r: list = []

    def push(self, label: str, before: pd.DataFrame, after: pd.DataFrame) -> None:
        self._u.append((label, before.copy()))
        if len(self._u) > self.MAX:
            self._u.pop(0)
        self._r.clear()

    def undo(self):
        if not self._u:
            return None, None
        l, df = self._u.pop()
        self._r.append((l, df))
        return l, df

    def redo(self):
        if not self._r:
            return None, None
        l, df = self._r.pop()
        self._u.append((l, df))
        return l, df

    def can_undo(self) -> bool: return bool(self._u)
    def can_redo(self) -> bool: return bool(self._r)
    def clear(self): self._u.clear(); self._r.clear()


# ══════════════════════════════════════════════════════════════════
#  AI ENGINE  ─  Groq / Llama-3
# ══════════════════════════════════════════════════════════════════

def build_profile(df: pd.DataFrame) -> str:
    lines = [
        f"SHAPE: {len(df):,} rows x {len(df.columns)} columns",
        f"MISSING CELLS: {int(df.isna().sum().sum()):,}",
        f"DUPLICATE ROWS: {int(df.duplicated().sum()):,}",
        f"JUNK ROWS (>80% empty): {int((df.isna().mean(axis=1) > 0.8).sum())}",
        "",
    ]
    for col in df.columns:
        s    = df[col]
        miss = int(s.isna().sum())
        pct  = f"{miss / max(len(df), 1) * 100:.1f}%"
        lines.append(
            f"COLUMN `{col}` | dtype={s.dtype} | "
            f"missing={miss}({pct}) | unique={s.nunique()}"
        )
        nn = s.dropna().astype(str).str.strip()
        if nn.empty:
            lines.append("  (all values missing)")
            lines.append("")
            continue
        lines.append(f"  samples: {nn.value_counts().head(8).index.tolist()}")
        if pd.api.types.is_numeric_dtype(s):
            num = pd.to_numeric(s, errors="coerce").dropna()
            if len(num):
                lines.append(
                    f"  min={num.min():.4g}  max={num.max():.4g}  "
                    f"mean={num.mean():.4g}  std={num.std():.4g}"
                )
            if int(np.isinf(pd.to_numeric(s, errors="coerce").fillna(0)).sum()):
                lines.append("  WARNING: contains inf / -inf values")
        else:
            nr = pd.to_numeric(
                nn.str.replace(r"[$,\s%€£¥]", "", regex=True), errors="coerce"
            ).notna().mean()
            if nr > 0.7:
                lines.append(f"  numeric-as-text rate: {nr * 100:.0f}%")
            try:
                dr = pd.to_datetime(
                    nn, errors="coerce", infer_datetime_format=True
                ).notna().mean()
                if dr > 0.5:
                    lines.append(f"  date-as-text rate: {dr * 100:.0f}%")
            except Exception:
                pass
            fmts = set()
            for v in nn.head(60):
                if re.search(r"\d{4}-\d{2}-\d{2}", v): fmts.add("YYYY-MM-DD")
                if re.search(r"\d{2}/\d{2}/\d{4}", v): fmts.add("DD/MM/YYYY")
                if re.search(r"\d{2} \w{3}", v):        fmts.add("DD Mon")
            if len(fmts) > 1:
                lines.append(f"  mixed date formats: {fmts}")
            vm = {}
            for v in nn.unique():
                vm.setdefault(v.lower().strip(), []).append(v)
            conflicts = {k: v for k, v in vm.items() if len(v) > 1}
            if conflicts:
                ex = "; ".join(
                    "/".join(x for x in v[:3])
                    for v in list(conflicts.values())[:3]
                )
                lines.append(f"  case conflicts: {ex}")
        lines.append("")
    return "\n".join(lines)


_AI_SYSTEM = """You are a world-class data analyst with 15 years of experience.
Given a dataset profile, find EVERY data quality problem like a senior data engineer would.

Look for:
- Missing / null values (which columns, how severe)
- Duplicate rows
- Numbers or dates stored as text
- Mixed date formats in same column
- Inconsistent capitalisation (e.g. "usa" vs "USA" vs "Us")
- Typos and misspellings (e.g. "New Zeland", "Pg-13" vs "PG-13")
- Wrong decimal scale (e.g. score=86.0 should be 8.6)
- Impossible values (inf, negative duration, 0 revenue)
- Ghost / unnamed columns (Col_0, Unnamed, completely empty)
- Truncated column names (very short, ends in consonant)
- Multi-value cells (genres/tags separated by commas)
- ID columns with duplicate values
- Columns that are >80% empty
- Outliers detected via IQR

Format your response EXACTLY like this — one block per issue:

ISSUE: [short clear title]
DETAIL: [specific description — use exact column names]
FIX: [what to do]
ACTION: [one of: fix_nulls | trim_whitespace | drop_blank_rows | drop_duplicates | remove_junk_rows | fix_numeric_columns | fix_date_columns | standardise_case | standardise_col_names | manual_review]
---

After all issues write:

SUMMARY:
[2-3 sentence overall assessment of the dataset quality]

Be specific. Use exact column names. No preamble. No filler."""


def run_groq_scan(
    df        : pd.DataFrame,
    api_key   : str,
    on_chunk  = None,
    on_done   = None,
    on_error  = None,
    on_status = None,
) -> None:
    def _worker():
        try:
            from groq import Groq
        except ImportError:
            if on_error:
                on_error(
                    "Package 'groq' not installed.\n\n"
                    "Run this in your VS Code terminal:\n\n"
                    "    pip install groq"
                )
            return

        if on_status: on_status("Building data profile…")
        profile = build_profile(df)

        if on_status: on_status("AI is thinking…")
        try:
            client = Groq(api_key=api_key)
            stream = client.chat.completions.create(
                model    = "llama-3.3-70b-versatile",
                messages = [
                    {"role": "system", "content": _AI_SYSTEM},
                    {"role": "user",   "content": f"Dataset profile:\n\n{profile}"},
                ],
                max_tokens  = 2500,
                temperature = 0.2,
                stream      = True,
            )
            chunks = []
            for chunk in stream:
                piece = chunk.choices[0].delta.content or ""
                chunks.append(piece)
                if on_chunk: on_chunk(piece)
            if on_done: on_done("".join(chunks))
        except Exception as e:
            if on_error:
                on_error(
                    f"Groq API error:\n\n{e}\n\n"
                    "Check your API key in the Settings tab."
                )

    threading.Thread(target=_worker, daemon=True).start()


def parse_cards(text: str) -> list:
    _EMAP = {
        "missing": "❓", "null": "❓", "blank": "❓",
        "duplicate": "⚠️", "numeric": "🔢", "number": "🔢",
        "date": "📅", "time": "📅", "case": "🔠", "capital": "🔠",
        "typo": "🔤", "spelling": "🔤", "outlier": "📊",
        "column": "👻", "ghost": "👻", "whitespace": "✂️",
        "space": "✂️", "currency": "💰", "junk": "🗑️",
        "scale": "🔢", "decimal": "🔢", "format": "📅",
        "truncat": "✂️", "id": "🔑", "empty": "🗑️",
    }
    cards = []
    for i, block in enumerate(re.split(r"\n---+\n?", text)):
        block = block.strip()
        if not block or "SUMMARY" in block.upper():
            continue
        im = re.search(r"ISSUE:\s*(.+)",  block, re.I)
        dm = re.search(r"DETAIL:\s*(.+)", block, re.I | re.S)
        am = re.search(r"ACTION:\s*(.+)", block, re.I)
        if not im:
            continue
        title  = im.group(1).strip()
        detail = ""
        if dm:
            detail = re.split(r"\nFIX:", dm.group(1), flags=re.I)[0].strip()
        action = (
            am.group(1).strip().lower().replace(" ", "_")
            if am else "manual_review"
        )
        tl     = (title + " " + detail).lower()
        emoji  = next((e for k, e in _EMAP.items() if k in tl), "⚠️")
        cards.append({
            "id": i, "emoji": emoji, "title": title,
            "detail": detail, "action": action, "applied": False,
        })
    return cards


_ACTION_MAP = {
    "fix_nulls"            : "fix_nulls",
    "trim_whitespace"      : "trim_whitespace",
    "drop_blank_rows"      : "drop_blank_rows",
    "drop_duplicates"      : "drop_duplicates",
    "remove_junk_rows"     : "remove_junk_rows",
    "fix_numeric_columns"  : "fix_numeric_columns",
    "fix_date_columns"     : "fix_date_columns",
    "standardise_case"     : "standardise_case",
    "standardise_col_names": "standardise_col_names",
}


# ══════════════════════════════════════════════════════════════════
#  CLEANER  ─  all operations + undo/redo
# ══════════════════════════════════════════════════════════════════

class Cleaner:
    def __init__(self, df: pd.DataFrame):
        self.df       = df.copy()
        self.stack    = UndoStack()
        self.modified : set = set()          # (row_idx, col_name) → gold highlight
        self._before  = self._snap()

    def _snap(self) -> dict:
        return {
            "rows": len(self.df),
            "cols": len(self.df.columns),
            "miss": int(self.df.isna().sum().sum()),
            "dupe": int(self.df.duplicated().sum()),
        }

    def _commit(self, label: str, before: pd.DataFrame, msg: str = "") -> str:
        self.stack.push(label, before, self.df)
        return msg or label

    # ── undo / redo ──────────────────────────────────────────────
    def undo(self) -> str | None:
        label, df = self.stack.undo()
        if df is not None:
            self.df = df.copy()
        return f"↩  Undone: {label}" if label else None

    def redo(self) -> str | None:
        label, df = self.stack.redo()
        if df is not None:
            self.df = df.copy()
        return f"↪  Redone: {label}" if label else None

    # ── cleaning operations ──────────────────────────────────────
    def fix_nulls(self) -> str:
        b = self.df.copy()
        n = int(self.df.isin(_NULLS).sum().sum())
        self.df = self.df.replace(_NULLS, pd.NA)
        return self._commit("Fix Nulls", b,
                            f"🚫  Replaced {n} null-like tokens with proper blanks.")

    def trim_whitespace(self) -> str:
        b = self.df.copy()
        for c in self.df.select_dtypes(include=["object", "string"]).columns:
            self.df[c] = self.df[c].astype("string").str.strip()
        return self._commit("Trim Whitespace", b,
                            "✂️   Trimmed whitespace from all text columns.")

    def drop_blank_rows(self) -> str:
        b = self.df.copy()
        n = int(self.df.isna().all(axis=1).sum())
        self.df = self.df.dropna(how="all").reset_index(drop=True)
        return self._commit("Drop Blank Rows", b,
                            f"🗑️   Dropped {n} fully blank rows.")

    def drop_duplicates(self) -> str:
        b = self.df.copy()
        n = int(self.df.duplicated().sum())
        self.df = self.df.drop_duplicates().reset_index(drop=True)
        return self._commit("Drop Duplicates", b,
                            f"🗑️   Dropped {n} duplicate rows.")

    def remove_junk_rows(self) -> str:
        b    = self.df.copy()
        mask = self.df.isna().mean(axis=1) > 0.80
        n    = int(mask.sum())
        self.df = self.df[~mask].reset_index(drop=True)
        return self._commit("Remove Junk Rows", b,
                            f"🗑️   Removed {n} junk rows (>80% empty).")

    def fix_numeric_columns(self) -> str:
        b     = self.df.copy()
        fixed = []
        for c in self.df.select_dtypes(include=["object", "string"]).columns:
            s = self.df[c].astype(str).str.replace(r"[$,\s%€£¥]", "", regex=True)
            if pd.to_numeric(s, errors="coerce").notna().mean() > 0.80:
                self.df[c] = pd.to_numeric(s, errors="coerce")
                fixed.append(c)
        msg = (f"🔢  Converted to numeric: {', '.join(f'`{c}`' for c in fixed)}."
               if fixed else "🔢  No text-as-numeric columns found.")
        return self._commit("Fix Numeric", b, msg)

    def fix_date_columns(self) -> str:
        b     = self.df.copy()
        fixed = []
        for c in self.df.select_dtypes(include=["object", "string"]).columns:
            try:
                p = pd.to_datetime(
                    self.df[c].astype(str),
                    errors="coerce", infer_datetime_format=True,
                )
                if p.notna().mean() > 0.70:
                    self.df[c] = p
                    fixed.append(c)
            except Exception:
                pass
        msg = (f"📅  Converted to datetime: {', '.join(f'`{c}`' for c in fixed)}."
               if fixed else "📅  No text-as-date columns found.")
        return self._commit("Fix Dates", b, msg)

    def standardise_case(self) -> str:
        b = self.df.copy()
        for c in self.df.select_dtypes(include=["object", "string"]).columns:
            self.df[c] = self.df[c].astype("string").str.title()
        return self._commit("Title Case", b,
                            "🔠  Applied Title Case to all text columns.")

    def standardise_col_names(self) -> str:
        b = self.df.copy()
        self.df.columns = [
            re.sub(r"\s+", "_", re.sub(r"[^\w\s]", "", str(c)).strip()).upper()
            for c in self.df.columns
        ]
        return self._commit("Fix Col Names", b,
                            "🏷️   Column names standardised to UPPER_SNAKE_CASE.")

    def replace_values(self, col: str, find, repl) -> str:
        if col not in self.df.columns:
            return f"❌  Column `{col}` not found."
        b = self.df.copy()
        n = int((self.df[col].astype(str) == str(find)).sum())
        self.df[col] = self.df[col].replace(find, repl)
        return self._commit("Replace Values", b,
                            f"🔄  Replaced '{find}' → '{repl}' in `{col}` ({n} cells).")

    def fill_missing(self, col: str, method: str = "value", value="") -> str:
        if col not in self.df.columns:
            return f"❌  Column `{col}` not found."
        b = self.df.copy()
        n = int(self.df[col].isna().sum())
        if   method == "value":  self.df[col] = self.df[col].fillna(value)
        elif method == "mean":   self.df[col] = self.df[col].fillna(self.df[col].mean())
        elif method == "median": self.df[col] = self.df[col].fillna(self.df[col].median())
        elif method == "mode":
            m = self.df[col].mode()
            if not m.empty: self.df[col] = self.df[col].fillna(m[0])
        elif method == "ffill":  self.df[col] = self.df[col].ffill()
        elif method == "bfill":  self.df[col] = self.df[col].bfill()
        return self._commit("Fill Missing", b,
                            f"❓  Filled {n} missing values in `{col}` (method: {method}).")

    def one_click_clean(self) -> list:
        return [
            fn() for fn in [
                self.fix_nulls, self.trim_whitespace, self.remove_junk_rows,
                self.drop_blank_rows, self.drop_duplicates,
                self.fix_numeric_columns, self.fix_date_columns,
                self.standardise_col_names,
            ]
        ]

    # ── data entry ───────────────────────────────────────────────
    def add_row(self, idx: int | None = None) -> str:
        b     = self.df.copy()
        blank = pd.DataFrame(
            [[pd.NA] * len(self.df.columns)], columns=self.df.columns
        )
        if idx is None or idx >= len(self.df) - 1:
            self.df = pd.concat([self.df, blank], ignore_index=True)
            pos = "bottom"
        else:
            self.df = pd.concat(
                [self.df.iloc[: idx + 1], blank, self.df.iloc[idx + 1 :]],
                ignore_index=True,
            )
            pos = f"after row {idx}"
        return self._commit("Add Row", b, f"➕  Added 1 blank row ({pos}).")

    def delete_row(self, idx: int) -> str:
        if idx < 0 or idx >= len(self.df):
            return "❌  Row index out of range."
        b = self.df.copy()
        self.df = self.df.drop(index=idx).reset_index(drop=True)
        return self._commit("Delete Row", b, f"🗑️   Deleted row {idx}.")

    def add_column(self, name: str, pos: int | None = None) -> str:
        if name in self.df.columns:
            return f"❌  Column `{name}` already exists."
        b = self.df.copy()
        if pos is None or pos >= len(self.df.columns):
            self.df[name] = pd.NA
        else:
            self.df.insert(pos, name, pd.NA)
        return self._commit("Add Column", b, f"➕  Added column `{name}`.")

    def delete_column(self, col: str) -> str:
        if col not in self.df.columns:
            return f"❌  Column `{col}` not found."
        b = self.df.copy()
        self.df = self.df.drop(columns=[col])
        return self._commit("Delete Column", b, f"🗑️   Deleted column `{col}`.")

    def rename_column(self, old: str, new: str) -> str:
        if old not in self.df.columns: return f"❌  `{old}` not found."
        if new in self.df.columns:     return f"❌  `{new}` already exists."
        b = self.df.copy()
        self.df = self.df.rename(columns={old: new})
        self.modified = {(r, new if c == old else c) for r, c in self.modified}
        return self._commit("Rename Column", b, f"✏️   Renamed `{old}` → `{new}`.")

    def edit_cell(self, row: int, col: str, val) -> str:
        if col not in self.df.columns:  return "❌  Column not found."
        if row < 0 or row >= len(self.df): return "❌  Row out of range."
        b   = self.df.copy()
        old = self.df.at[row, col]
        if pd.api.types.is_numeric_dtype(self.df[col]):
            try:
                val = float(val) if "." in str(val) else int(val)
            except (ValueError, TypeError):
                return f"⚠️  `{col}` is numeric — '{val}' is not a valid number."
        self.df.at[row, col] = val
        self.modified.add((row, col))
        return self._commit("Edit Cell", b,
                            f"✏️   [{row}, `{col}`]:  '{old}'  →  '{val}'.")

    def get_comparison(self) -> dict:
        a = self._snap()
        b = self._before
        return {
            "before"        : b,
            "after"         : a,
            "rows_removed"  : b["rows"] - a["rows"],
            "missing_fixed" : b["miss"] - a["miss"],
            "dupes_removed" : b["dupe"] - a["dupe"],
        }

    def reset(self, original: pd.DataFrame) -> None:
        self.df       = original.copy()
        self.modified.clear()
        self.stack.clear()
        self._before  = self._snap()


# ══════════════════════════════════════════════════════════════════
#  SPLASH SCREEN
# ══════════════════════════════════════════════════════════════════

class Splash:
    def __init__(self, root: tk.Tk, on_done):
        self.root    = root
        self.on_done = on_done
        root.overrideredirect(True)
        root.configure(bg=T["bg"])
        root.attributes("-topmost", True)
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w, h   = 620, 390
        root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")
        self._build()
        self._steps = [
            (12, "Loading UI components…"),
            (30, "Initializing AI engine…"),
            (52, "Loading data modules…"),
            (74, "Setting up workspace…"),
            (92, "Almost ready…"),
            (100,"Welcome, Sajeeb! 🚀"),
        ]
        self._idx = 0
        self._step()

    def _build(self):
        c = tk.Canvas(self.root, width=620, height=390,
                      bg=T["bg"], highlightthickness=0)
        c.pack(fill="both", expand=True)
        self.c = c
        # Borders
        c.create_rectangle(2,  2, 618, 388, outline=T["gold"],   width=2)
        c.create_rectangle(7,  7, 613, 383, outline=T["border"], width=1)
        # Logo
        self._draw_logo(c, 310, 98)
        # Title
        c.create_text(310, 168, text="Excel AI Cleaner",
                      fill=T["gold_bright"], font=("Segoe UI", 26, "bold"))
        c.create_text(310, 198,
                      text="Smart Data Cleaning  •  Powered by Groq AI  (Free)",
                      fill=T["dim"], font=("Segoe UI", 10))
        # Divider
        c.create_line(70, 222, 550, 222, fill=T["border"], width=1)
        # Developer
        c.create_text(310, 244, text="Developed by",
                      fill=T["dim"], font=("Segoe UI", 9))
        c.create_text(310, 266, text="Sajeeb The Analyst",
                      fill=T["gold"], font=("Segoe UI", 16, "bold"))
        c.create_text(310, 290, text="v1.0.0  •  Professional Edition",
                      fill=T["dim"], font=("Segoe UI", 8))
        # Progress track
        c.create_rectangle(70, 322, 550, 338,
                           fill=T["border"], outline="")
        # Progress fill
        self.bar = c.create_rectangle(70, 322, 70, 338,
                                      fill=T["gold"], outline="")
        # Status
        self.stat = c.create_text(310, 356, text="Initializing…",
                                  fill=T["dim"], font=("Segoe UI", 9))

    def _draw_logo(self, c, cx, cy):
        # Circle background
        c.create_oval(cx-46, cy-46, cx+46, cy+46,
                      fill=T["panel"], outline=T["gold"], width=2)
        # Chart bars
        for x1, y1, x2, y2, col in [
            (cx-27, cy+22, cx-15, cy- 8, T["gold"]),
            (cx- 8, cy+22, cx+ 4, cy-26, T["gold_bright"]),
            (cx+10, cy+22, cx+22, cy+ 2, T["gold"]),
        ]:
            c.create_rectangle(x1, y1, x2, y2, fill=col, outline="")
        # Sparkles
        for dx, dy, r in [(-33, -30, 3), (35, -24, 2), (31, 20, 2), (-31, 26, 2)]:
            c.create_oval(cx+dx-r, cy+dy-r, cx+dx+r, cy+dy+r,
                          fill=T["gold_bright"], outline="")

    def _step(self):
        if self._idx >= len(self._steps):
            self.root.after(500, self.on_done)
            return
        pct, msg = self._steps[self._idx]
        self._idx += 1
        bx = 70 + int(480 * pct / 100)
        self.c.coords(self.bar, 70, 322, bx, 338)
        self.c.itemconfig(self.stat, text=msg)
        self.root.after(360 if pct < 100 else 520, self._step)


# ══════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ══════════════════════════════════════════════════════════════════

class App:
    def __init__(self, root: tk.Tk):
        self.root    = root
        self.cfg     = cfg_load()

        self.cleaner  : Cleaner | None = None
        self.original : pd.DataFrame   = pd.DataFrame()
        self.filepath : str            = ""
        self.ai_cards : list           = []
        self._card_widgets: list       = []

        # ── window setup
        w  = self.cfg.get("window_w", 1420)
        h  = self.cfg.get("window_h", 840)
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        root.title("Excel AI Cleaner  —  by Sajeeb The Analyst")
        root.configure(bg=T["bg"])
        root.protocol("WM_DELETE_WINDOW", self._on_close)

        apply_theme(ttk.Style(root))
        self._build_ui()
        self._bind_shortcuts()

    # ══════════════════════════════════════════════════════════
    #  UI  BUILD
    # ══════════════════════════════════════════════════════════

    def _build_ui(self):
        # ── Top toolbar ──────────────────────────────────────
        bar = tk.Frame(self.root, bg=T["bg"], pady=7)
        bar.pack(fill="x", padx=12)

        tk.Label(
            bar, text="⬡  Excel AI Cleaner",
            bg=T["bg"], fg=T["gold"],
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left", padx=(0, 18))

        _btns = [
            ("📂  Open File",         self.open_file,      "TButton"),
            ("🤖  AI Scan",           self.ai_scan,        "Gold.TButton"),
            ("✨  One Click Clean",   self.one_click,      "TButton"),
            ("↩  Undo",               self.undo,           "TButton"),
            ("↪  Redo",               self.redo,           "TButton"),
            ("↺  Reset",              self.reset,          "TButton"),
            ("💾  Export CSV",        self.export_csv,     "TButton"),
            ("📊  Export Excel",      self.export_excel,   "TButton"),
            ("📄  Report",            self.export_report,  "TButton"),
        ]
        for label, cmd, sty in _btns:
            ttk.Button(bar, text=label, command=cmd,
                       style=sty).pack(side="left", padx=3)

        self.lbl_file = ttk.Label(bar, text="No file opened",
                                  style="Dim.TLabel")
        self.lbl_file.pack(side="left", padx=14)

        self.var_status = tk.StringVar(value="Ready")
        tk.Label(
            bar, textvariable=self.var_status,
            bg=T["bg"], fg=T["gold"],
            font=("Segoe UI", 9, "bold"),
        ).pack(side="right", padx=10)

        # ── Progress bar ─────────────────────────────────────
        self.pbar = ttk.Progressbar(self.root, mode="indeterminate")
        self.pbar.pack(fill="x", padx=12, pady=2)

        # ── Notebook ─────────────────────────────────────────
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=12, pady=4)

        self._tabs = {}
        for key, label in [
            ("preview",  "📋  Data Preview"),
            ("ai",       "🤖  AI Issues"),
            ("clean",    "🧹  Clean"),
            ("edit",     "✏️   Edit Data"),
            ("pivot",    "📊  Pivot"),
            ("chart",    "📈  Chart"),
            ("settings", "⚙️   Settings"),
        ]:
            f = ttk.Frame(self.nb)
            self.nb.add(f, text=label)
            self._tabs[key] = f

        self._build_preview(self._tabs["preview"])
        self._build_ai(self._tabs["ai"])
        self._build_clean(self._tabs["clean"])
        self._build_edit(self._tabs["edit"])
        self._build_pivot(self._tabs["pivot"])
        self._build_chart(self._tabs["chart"])
        self._build_settings(self._tabs["settings"])

        # ── Bottom status bar ─────────────────────────────────
        tk.Frame(self.root, bg=T["border"], height=1).pack(fill="x")
        bot = tk.Frame(self.root, bg=T["panel"], pady=5)
        bot.pack(fill="x")

        self.lbl_sum = tk.Label(
            bot,
            text="Rows: —  |  Columns: —  |  Missing: —  |  Duplicates: —",
            bg=T["panel"], fg=T["dim"], font=("Segoe UI", 9),
        )
        self.lbl_sum.pack(side="left", padx=12)

        tk.Label(
            bot,
            text="Sajeeb The Analyst  •  Excel AI Cleaner  v1.0",
            bg=T["panel"], fg=T["gold_dim"],
            font=("Segoe UI", 8),
        ).pack(side="right", padx=12)

    # ── Preview tab ──────────────────────────────────────────────

    def _build_preview(self, p):
        tf = tk.Frame(p, bg=T["bg"])
        tf.pack(fill="both", expand=True, padx=8, pady=8)

        self.tree = ttk.Treeview(tf, show="headings", selectmode="browse")
        self.tree.pack(side="left", fill="both", expand=True)

        sy = ttk.Scrollbar(tf, orient="vertical",   command=self.tree.yview)
        sy.pack(side="right", fill="y")
        sx = ttk.Scrollbar(p,  orient="horizontal", command=self.tree.xview)
        sx.pack(fill="x", padx=8)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)

        # interactions
        self.tree.bind("<Double-1>",  self._cell_dbl_click)
        self.tree.bind("<Button-3>",  self._tree_right_click)

    # ── AI tab ───────────────────────────────────────────────────

    def _build_ai(self, p):
        # header row
        hdr = tk.Frame(p, bg=T["bg"])
        hdr.pack(fill="x", padx=12, pady=8)
        tk.Label(
            hdr, text="🤖  AI Problem Finder  —  Groq / Llama-3.3-70B",
            bg=T["bg"], fg=T["gold_bright"],
            font=("Segoe UI", 12, "bold"),
        ).pack(side="left")
        ttk.Button(hdr, text="▶  Run Scan",
                   command=self.ai_scan,
                   style="Gold.TButton").pack(side="right", padx=4)
        ttk.Button(hdr, text="🗑  Clear",
                   command=self._clear_ai).pack(side="right", padx=4)

        # paned: left=stream, right=cards
        pane = tk.PanedWindow(
            p, orient="horizontal",
            bg=T["border"], sashwidth=5, sashrelief="flat",
        )
        pane.pack(fill="both", expand=True, padx=8, pady=4)

        # ── left: streaming text ──────────────────────────────
        lf = tk.Frame(pane, bg=T["bg"])
        pane.add(lf, minsize=320)

        tk.Label(lf, text="Live AI Analysis",
                 bg=T["bg"], fg=T["dim"],
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=6, pady=4)

        self.ai_box = tk.Text(
            lf, wrap="word", font=("Consolas", 9),
            bg="#0a0816", fg=T["text"],
            insertbackground=T["gold"],
            selectbackground=T["select"],
            relief="flat", bd=0, padx=10, pady=10,
        )
        self.ai_box.pack(fill="both", expand=True)
        self.ai_box.tag_config("head",  foreground=T["gold_bright"],
                               font=("Consolas", 10, "bold"))
        self.ai_box.tag_config("issue", foreground=T["warning"])
        self.ai_box.tag_config("ok",    foreground=T["success"])
        self.ai_box.tag_config("dim",   foreground=T["dim"])

        # ── right: fix cards ──────────────────────────────────
        rf = tk.Frame(pane, bg=T["bg"])
        pane.add(rf, minsize=270)

        tk.Label(rf, text="Fix Cards  —  click Apply to fix instantly",
                 bg=T["bg"], fg=T["dim"],
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=6, pady=4)

        card_outer = tk.Frame(rf, bg=T["bg"])
        card_outer.pack(fill="both", expand=True)

        self._cards_canvas = tk.Canvas(card_outer, bg=T["bg"], highlightthickness=0)
        self._cards_canvas.pack(side="left", fill="both", expand=True)

        sc = ttk.Scrollbar(card_outer, orient="vertical",
                           command=self._cards_canvas.yview)
        sc.pack(side="right", fill="y")
        self._cards_canvas.configure(yscrollcommand=sc.set)

        self._cards_frame = tk.Frame(self._cards_canvas, bg=T["bg"])
        self._cards_canvas.create_window((0, 0), window=self._cards_frame, anchor="nw")
        self._cards_frame.bind(
            "<Configure>",
            lambda e: self._cards_canvas.configure(
                scrollregion=self._cards_canvas.bbox("all")
            ),
        )

    # ── Clean tab ────────────────────────────────────────────────

    def _build_clean(self, p):
        f = tk.Frame(p, bg=T["bg"])
        f.pack(fill="both", expand=True, padx=14, pady=12)

        # Row 1 — find / replace
        r1 = tk.Frame(f, bg=T["bg"]); r1.pack(fill="x", pady=5)
        self._lbl(r1, "Column"); self.cb_col = ttk.Combobox(r1, state="readonly", width=18)
        self.cb_col.pack(side="left", padx=4)
        self._lbl(r1, "Find"); self.ent_find = ttk.Entry(r1, width=16)
        self.ent_find.pack(side="left", padx=4)
        self._lbl(r1, "Replace With"); self.ent_repl = ttk.Entry(r1, width=16)
        self.ent_repl.pack(side="left", padx=4)
        ttk.Button(r1, text="Replace All",
                   command=self._do_replace).pack(side="left", padx=6)

        # Row 2 — fill missing
        r2 = tk.Frame(f, bg=T["bg"]); r2.pack(fill="x", pady=5)
        self._lbl(r2, "Fill Missing in Column:")
        self.cb_fill_col = ttk.Combobox(r2, state="readonly", width=18)
        self.cb_fill_col.pack(side="left", padx=4)
        self._lbl(r2, "Method:")
        self.cb_fill_m = ttk.Combobox(
            r2, state="readonly", width=10,
            values=["value", "mean", "median", "mode", "ffill", "bfill"],
        )
        self.cb_fill_m.set("value"); self.cb_fill_m.pack(side="left", padx=4)
        self.ent_fill_v = ttk.Entry(r2, width=12)
        self.ent_fill_v.pack(side="left", padx=4)
        ttk.Button(r2, text="Fill", command=self._do_fill).pack(side="left", padx=4)

        # Row 3 — action buttons
        r3 = tk.Frame(f, bg=T["bg"]); r3.pack(fill="x", pady=5)
        for label, method in [
            ("Drop Blank Rows",    "drop_blank_rows"),
            ("Drop Duplicates",    "drop_duplicates"),
            ("Trim Whitespace",    "trim_whitespace"),
            ("Fix Dates",          "fix_date_columns"),
            ("Fix Numbers",        "fix_numeric_columns"),
            ("Title Case",         "standardise_case"),
            ("Junk Rows",          "remove_junk_rows"),
            ("Fix Col Names",      "standardise_col_names"),
        ]:
            ttk.Button(
                r3, text=label,
                command=lambda m=method: self._run_clean(m),
            ).pack(side="left", padx=3)

        # Log
        self._lbl_section(f, "Clean Log")
        self.clean_log = tk.Text(
            f, wrap="word", font=("Consolas", 9),
            bg=T["panel"], fg=T["text"],
            relief="flat", bd=0, padx=8, pady=8,
            insertbackground=T["gold"],
        )
        self.clean_log.pack(fill="both", expand=True)

    # ── Edit tab ─────────────────────────────────────────────────

    def _build_edit(self, p):
        f = tk.Frame(p, bg=T["bg"])
        f.pack(fill="both", expand=True, padx=14, pady=12)

        tk.Label(
            f, text="✏️  Direct Data Editing",
            bg=T["bg"], fg=T["gold_bright"],
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", pady=(0, 10))

        # Row/column operations
        r1 = tk.Frame(f, bg=T["bg"]); r1.pack(fill="x", pady=5)
        self._lbl(r1, "Row Index:")
        self.ent_ridx = ttk.Entry(r1, width=8)
        self.ent_ridx.pack(side="left", padx=4)
        for label, cmd in [
            ("➕  Add Row Below", self._add_row),
            ("🗑   Delete Row",   self._del_row),
        ]:
            ttk.Button(r1, text=label, command=cmd).pack(side="left", padx=4)

        r2 = tk.Frame(f, bg=T["bg"]); r2.pack(fill="x", pady=5)
        self._lbl(r2, "Column Name:")
        self.ent_cname = ttk.Entry(r2, width=20)
        self.ent_cname.pack(side="left", padx=4)
        for label, cmd in [
            ("➕  Add Column",    self._add_col),
            ("🗑   Delete Column", self._del_col),
            ("✏️   Rename Column", self._ren_col),
        ]:
            ttk.Button(r2, text=label, command=cmd).pack(side="left", padx=4)

        # Edit single cell
        r3 = tk.Frame(f, bg=T["bg"]); r3.pack(fill="x", pady=5)
        self._lbl(r3, "Edit Cell  —  Row:")
        self.ent_erow = ttk.Entry(r3, width=7)
        self.ent_erow.pack(side="left", padx=4)
        self._lbl(r3, "Column:")
        self.cb_ecol = ttk.Combobox(r3, state="readonly", width=18)
        self.cb_ecol.pack(side="left", padx=4)
        self._lbl(r3, "New Value:")
        self.ent_eval = ttk.Entry(r3, width=20)
        self.ent_eval.pack(side="left", padx=4)
        ttk.Button(
            r3, text="✅  Apply Edit",
            style="Gold.TButton",
            command=self._edit_cell,
        ).pack(side="left", padx=8)

        # Tips
        tk.Label(
            f,
            text=(
                "💡  Tip: Double-click any cell in Data Preview to edit it directly.\n"
                "    Right-click a row in Data Preview → Add Row Below / Delete Row.\n"
                "    Modified cells are highlighted in gold.  Every action is undoable (Ctrl+Z)."
            ),
            bg=T["bg"], fg=T["dim"],
            font=("Segoe UI", 9), justify="left",
        ).pack(anchor="w", pady=10)

        # Edit log
        self._lbl_section(f, "Edit Log")
        self.edit_log = tk.Text(
            f, wrap="word", font=("Consolas", 9),
            bg=T["panel"], fg=T["text"],
            relief="flat", bd=0, padx=8, pady=8,
            insertbackground=T["gold"],
        )
        self.edit_log.pack(fill="both", expand=True)

    # ── Pivot tab ────────────────────────────────────────────────

    def _build_pivot(self, p):
        ctrl = tk.Frame(p, bg=T["bg"]); ctrl.pack(fill="x", padx=12, pady=10)

        self._lbl(ctrl, "Group By")
        self.cb_piv_idx = ttk.Combobox(ctrl, state="readonly", width=18)
        self.cb_piv_idx.pack(side="left", padx=4)
        self._lbl(ctrl, "Value")
        self.cb_piv_val = ttk.Combobox(ctrl, state="readonly", width=18)
        self.cb_piv_val.pack(side="left", padx=4)
        self._lbl(ctrl, "Aggregation")
        self.cb_piv_agg = ttk.Combobox(ctrl, state="readonly", width=10,
                                        values=["sum","mean","count","min","max"])
        self.cb_piv_agg.set("sum"); self.cb_piv_agg.pack(side="left", padx=4)
        ttk.Button(ctrl, text="Create Pivot",
                   style="Gold.TButton",
                   command=self._make_pivot).pack(side="left", padx=10)

        tf = tk.Frame(p, bg=T["bg"]); tf.pack(fill="both", expand=True, padx=12, pady=4)
        self.pivot_tree = ttk.Treeview(tf, show="headings")
        self.pivot_tree.pack(side="left", fill="both", expand=True)
        ttk.Scrollbar(tf, orient="vertical",
                      command=self.pivot_tree.yview).pack(side="right", fill="y")

    # ── Chart tab ────────────────────────────────────────────────

    def _build_chart(self, p):
        ctrl = tk.Frame(p, bg=T["bg"]); ctrl.pack(fill="x", padx=12, pady=10)

        self._lbl(ctrl, "Chart Type")
        self.cb_cht_type = ttk.Combobox(ctrl, state="readonly", width=12,
                                         values=["bar","line","pie","histogram","scatter"])
        self.cb_cht_type.set("bar"); self.cb_cht_type.pack(side="left", padx=4)
        self._lbl(ctrl, "X / Category")
        self.cb_cht_x = ttk.Combobox(ctrl, state="readonly", width=20)
        self.cb_cht_x.pack(side="left", padx=4)
        self._lbl(ctrl, "Y / Value")
        self.cb_cht_y = ttk.Combobox(ctrl, state="readonly", width=20)
        self.cb_cht_y.pack(side="left", padx=4)
        ttk.Button(ctrl, text="Show Chart",
                   style="Gold.TButton",
                   command=self._show_chart).pack(side="left", padx=10)

        tk.Label(
            p,
            text="Tip: select a text column as X (category) and a numeric column as Y (value).",
            bg=T["bg"], fg=T["dim"], font=("Segoe UI", 9),
        ).pack(anchor="w", padx=14, pady=4)

    # ── Settings tab ─────────────────────────────────────────────

    def _build_settings(self, p):
        f = tk.Frame(p, bg=T["bg"])
        f.pack(fill="both", expand=True, padx=30, pady=26)

        tk.Label(f, text="🔑  Groq API Key  (100% Free)",
                 bg=T["bg"], fg=T["gold_bright"],
                 font=("Segoe UI", 13, "bold")).pack(anchor="w")

        tk.Label(f, text="Get your free key → https://console.groq.com  "
                         "(sign up free, click API Keys, Create Key)",
                 bg=T["bg"], fg=T["dim"],
                 font=("Segoe UI", 9)).pack(anchor="w", pady=5)

        key_row = tk.Frame(f, bg=T["bg"]); key_row.pack(fill="x", pady=10)
        self.var_key = tk.StringVar(value=self.cfg.get("api_key", ""))
        self.ent_key = ttk.Entry(key_row, textvariable=self.var_key,
                                 width=56, show="*")
        self.ent_key.pack(side="left", padx=4)
        ttk.Button(key_row, text="👁  Show / Hide",
                   command=lambda: self.ent_key.config(
                       show="" if self.ent_key.cget("show") == "*" else "*"
                   )).pack(side="left", padx=4)
        ttk.Button(key_row, text="💾  Save Key",
                   command=self._save_key).pack(side="left", padx=4)
        ttk.Button(key_row, text="✅  Test Key",
                   command=self._test_key).pack(side="left", padx=4)

        tk.Frame(f, bg=T["border"], height=1).pack(fill="x", pady=18)

        tk.Label(f, text="ℹ️  How the App Works",
                 bg=T["bg"], fg=T["gold"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")

        tk.Label(f, justify="left", bg=T["bg"], fg=T["dim"],
                 font=("Segoe UI", 9), text=(
            "1.  Open any CSV / Excel / JSON file — even huge or messy ones.\n"
            "2.  Click  🤖 AI Scan  — the app builds a statistical profile of your data\n"
            "    and sends ONLY that summary to Groq AI.  Your raw data never leaves your PC.\n"
            "3.  Llama-3.3-70B reads the profile and thinks like a real data analyst —\n"
            "    it finds every problem: typos, wrong types, mixed dates, ghost columns,\n"
            "    wrong decimal scales, inconsistent capitalisation, and more.\n"
            "4.  Each problem becomes a clickable Fix Card.  Press Apply to fix it instantly.\n"
            "5.  ✨ One Click Clean  auto-fixes ALL common issues in one shot.\n"
            "6.  Edit cells, add/delete rows & columns directly in the  ✏️ Edit Data  tab.\n"
            "    Double-click any cell in Data Preview to edit inline.\n"
            "7.  Full Undo / Redo history — Ctrl+Z  /  Ctrl+Shift+Z.\n"
            "8.  Export cleaned data as CSV or Excel.  Generate a text Quality Report.\n\n"
            "✅  100% Free  •  Works on any file size  •  No internet needed for cleaning"
        )).pack(anchor="w", pady=6)

        tk.Frame(f, bg=T["border"], height=1).pack(fill="x", pady=14)
        tk.Label(f, text="⌨️  Keyboard Shortcuts",
                 bg=T["bg"], fg=T["gold"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(f, justify="left", bg=T["bg"], fg=T["dim"],
                 font=("Segoe UI", 9), text=(
            "Ctrl+O  — Open file          Ctrl+S — Export CSV\n"
            "Ctrl+Z  — Undo               Ctrl+Shift+Z — Redo\n"
            "Ctrl+Shift+S  — AI Scan      Ctrl+E — Export Report"
        )).pack(anchor="w", pady=4)

    # ── small helpers ────────────────────────────────────────────

    def _lbl(self, parent, text: str) -> ttk.Label:
        lbl = ttk.Label(parent, text=text, style="Dim.TLabel")
        lbl.pack(side="left", padx=4)
        return lbl

    def _lbl_section(self, parent, text: str):
        tk.Label(
            parent, text=text,
            bg=T["bg"], fg=T["dim"],
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w", pady=(10, 3))

    # ══════════════════════════════════════════════════════════
    #  KEYBOARD SHORTCUTS
    # ══════════════════════════════════════════════════════════

    def _bind_shortcuts(self):
        r = self.root
        r.bind("<Control-o>",       lambda e: self.open_file())
        r.bind("<Control-s>",       lambda e: self.export_csv())
        r.bind("<Control-z>",       lambda e: self.undo())
        r.bind("<Control-Z>",       lambda e: self.redo())
        r.bind("<Control-Shift-Z>", lambda e: self.redo())
        r.bind("<Control-Shift-s>", lambda e: self.ai_scan())
        r.bind("<Control-e>",       lambda e: self.export_report())

    # ══════════════════════════════════════════════════════════
    #  FILE OPERATIONS
    # ══════════════════════════════════════════════════════════

    def open_file(self):
        folder = self.cfg.get("last_folder", "") or "/"
        path   = filedialog.askopenfilename(
            initialdir=folder,
            filetypes=[
                ("Data Files", "*.csv *.xlsx *.xls *.json"),
                ("All Files",  "*.*"),
            ],
        )
        if not path:
            return
        self._busy(True, "Loading file…")
        threading.Thread(target=self._load_thread,
                         args=(path,), daemon=True).start()

    def _load_thread(self, path: str):
        try:
            df = load_file(path)
        except Exception as e:
            self.root.after(0, lambda: (
                messagebox.showerror("Load Error", str(e)),
                self._busy(False, "Load failed"),
            ))
            return
        self.root.after(0, lambda: self._finish_load(path, df))

    def _finish_load(self, path: str, df: pd.DataFrame):
        self.filepath = path
        self.original = df.copy()
        self.cleaner  = Cleaner(df)
        self.cfg["last_folder"] = os.path.dirname(path)
        cfg_save(self.cfg)
        self.lbl_file.config(
            text=os.path.basename(path),
            foreground=T["gold"],
        )
        self._refresh()
        self._busy(False, f"Loaded  {len(df):,} rows × {len(df.columns)} columns")

    # ══════════════════════════════════════════════════════════
    #  REFRESH
    # ══════════════════════════════════════════════════════════

    def _refresh(self):
        self._update_summary()
        self._fill_tree()
        self._update_combos()

    def _update_summary(self):
        if not self.cleaner or self.cleaner.df.empty:
            self.lbl_sum.config(
                text="Rows: —  |  Columns: —  |  Missing: —  |  Duplicates: —"
            )
            return
        df = self.cleaner.df
        self.lbl_sum.config(
            text=(
                f"Rows: {len(df):,}  |  "
                f"Columns: {len(df.columns)}  |  "
                f"Missing: {int(df.isna().sum().sum()):,}  |  "
                f"Duplicates: {int(df.duplicated().sum()):,}"
            )
        )

    def _fill_tree(self):
        self.tree.delete(*self.tree.get_children())
        if not self.cleaner:
            return
        df = self.cleaner.df.head(500)
        self.tree["columns"] = list(df.columns)
        for c in df.columns:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120, anchor="center", minwidth=50)
        self.tree.tag_configure(
            "modified",
            background=T["gold_cell"],
            foreground=T["gold_bright"],
        )
        for i, (_, row) in enumerate(df.iterrows()):
            vals = ["" if pd.isna(v) else str(v) for v in row]
            tags = ("modified",) if any(
                (i, c) in self.cleaner.modified for c in df.columns
            ) else ()
            self.tree.insert("", "end", values=vals, tags=tags)

    def _update_combos(self):
        cols = [] if not self.cleaner else self.cleaner.df.columns.tolist()
        nums = [] if not self.cleaner else \
               self.cleaner.df.select_dtypes(include="number").columns.tolist()

        for cb in [self.cb_col, self.cb_fill_col, self.cb_ecol,
                   self.cb_piv_idx, self.cb_cht_x]:
            cb["values"] = cols
            if cols: cb.set(cols[0])

        for cb in [self.cb_piv_val, self.cb_cht_y]:
            cb["values"] = nums or cols
            if (nums or cols): cb.set((nums or cols)[0])

    # ══════════════════════════════════════════════════════════
    #  AI SCAN
    # ══════════════════════════════════════════════════════════

    def ai_scan(self):
        if not self.cleaner:
            messagebox.showwarning("No Data", "Please open a file first.")
            return
        key = self.cfg.get("api_key", "").strip()
        if not key:
            messagebox.showwarning(
                "No API Key",
                "Please add your free Groq API key in the ⚙️ Settings tab.\n\n"
                "Get it free at:  https://console.groq.com",
            )
            self.nb.select(self._tabs["settings"])
            return

        self.ai_box.delete("1.0", "end")
        self._clear_cards()
        self._ai_write("🤖  AI Problem Finder  —  Groq / Llama-3.3-70B\n", "head")
        self._ai_write(f"File : {os.path.basename(self.filepath)}\n", "dim")
        self._ai_write(
            f"Size : {len(self.cleaner.df):,} rows × "
            f"{len(self.cleaner.df.columns)} columns\n\n", "dim",
        )
        self.nb.select(self._tabs["ai"])
        self._busy(True, "AI is thinking…")

        run_groq_scan(
            df        = self.cleaner.df,
            api_key   = key,
            on_chunk  = lambda p: self.root.after(
                0, lambda piece=p: self._stream(piece)),
            on_done   = lambda t: self.root.after(
                0, lambda txt=t: self._scan_done(txt)),
            on_error  = lambda e: self.root.after(
                0, lambda err=e: self._scan_error(err)),
            on_status = lambda s: self.root.after(
                0, lambda m=s: self.var_status.set(m)),
        )

    def _stream(self, piece: str):
        self.ai_box.insert("end", piece)
        self.ai_box.see("end")

    def _scan_done(self, full_text: str):
        self._busy(False, "✅  AI scan complete")
        cards = parse_cards(full_text)
        self.ai_cards = cards
        self._render_cards(cards)

    def _scan_error(self, err: str):
        self._busy(False, "❌  Scan failed")
        self._ai_write("\n" + err + "\n", "issue")

    def _ai_write(self, text: str, tag: str | None = None):
        self.ai_box.insert("end", text, tag or "")
        self.ai_box.see("end")

    def _clear_ai(self):
        self.ai_box.delete("1.0", "end")
        self._clear_cards()

    # ── Fix Cards ────────────────────────────────────────────────

    def _clear_cards(self):
        for w in self._card_widgets:
            w.destroy()
        self._card_widgets.clear()

    def _render_cards(self, cards: list):
        self._clear_cards()
        for card in cards:
            self._make_card(card)

    def _make_card(self, card: dict):
        outer = tk.Frame(
            self._cards_frame,
            bg=T["panel2"],
            highlightbackground=T["border"],
            highlightthickness=1,
            pady=9, padx=10,
        )
        outer.pack(fill="x", padx=6, pady=5)
        self._card_widgets.append(outer)

        # Title
        tk.Label(
            outer,
            text=f"{card['emoji']}  {card['title']}",
            bg=T["panel2"], fg=T["gold"],
            font=("Segoe UI", 9, "bold"),
            wraplength=224, justify="left",
        ).pack(anchor="w")

        # Detail
        if card.get("detail"):
            tk.Label(
                outer,
                text=card["detail"],
                bg=T["panel2"], fg=T["dim"],
                font=("Segoe UI", 8),
                wraplength=234, justify="left",
            ).pack(anchor="w", pady=(3, 5))

        # Apply / manual button
        if card.get("action", "manual_review") != "manual_review":
            ttk.Button(
                outer,
                text="✅  Apply Fix",
                style="Gold.TButton",
                command=lambda c=card, f=outer: self._apply_card(c, f),
            ).pack(anchor="e", pady=3)
        else:
            tk.Label(
                outer,
                text="⚠️  Needs manual review",
                bg=T["panel2"], fg=T["warning"],
                font=("Segoe UI", 8),
            ).pack(anchor="e")

    def _apply_card(self, card: dict, frame: tk.Frame):
        if card.get("applied"):
            messagebox.showinfo("Already Applied",
                                f"'{card['title']}' was already applied.")
            return
        if not self.cleaner:
            return
        method_name = _ACTION_MAP.get(card.get("action", ""))
        if not method_name:
            messagebox.showinfo("Manual",
                                "This issue requires manual review — no auto-fix available.")
            return
        method = getattr(self.cleaner, method_name, None)
        if method is None:
            return
        msg             = method()
        card["applied"] = True
        # Dim the card to show it's done
        frame.configure(bg=T["panel"], highlightbackground=T["gold_dim"])
        for w in frame.winfo_children():
            try:
                w.configure(bg=T["panel"])
            except Exception:
                pass
        self._refresh()
        self._log_clean(msg)
        self.var_status.set(f"✅  Applied: {card['title']}")

    # ══════════════════════════════════════════════════════════
    #  ONE CLICK CLEAN
    # ══════════════════════════════════════════════════════════

    def one_click(self):
        if not self.cleaner:
            messagebox.showwarning("No Data", "Open a file first.")
            return
        self._busy(True, "Cleaning…")

        def _run():
            log = self.cleaner.one_click_clean()
            self.root.after(0, lambda: self._finish_clean(log))

        threading.Thread(target=_run, daemon=True).start()

    def _finish_clean(self, log: list):
        self._refresh()
        self.clean_log.insert("end", "✨  One Click Clean — Done\n" + "─" * 52 + "\n")
        for line in log:
            self.clean_log.insert("end", line + "\n")
        self.clean_log.insert("end", "\n")
        self.clean_log.see("end")
        self.nb.select(self._tabs["clean"])
        self._busy(False, "✅  Clean complete")

    # ══════════════════════════════════════════════════════════
    #  CLEAN ACTIONS
    # ══════════════════════════════════════════════════════════

    def _run_clean(self, method: str):
        if not self.cleaner:
            return
        fn = getattr(self.cleaner, method, None)
        if fn:
            self._log_clean(fn())
            self._refresh()

    def _do_replace(self):
        if not self.cleaner:
            return
        msg = self.cleaner.replace_values(
            self.cb_col.get(),
            self.ent_find.get(),
            self.ent_repl.get(),
        )
        self._log_clean(msg)
        self._refresh()

    def _do_fill(self):
        if not self.cleaner:
            return
        msg = self.cleaner.fill_missing(
            self.cb_fill_col.get(),
            self.cb_fill_m.get(),
            self.ent_fill_v.get(),
        )
        self._log_clean(msg)
        self._refresh()

    def _log_clean(self, msg: str):
        if msg:
            self.clean_log.insert("end", msg + "\n")
            self.clean_log.see("end")

    # ══════════════════════════════════════════════════════════
    #  EDIT DATA ACTIONS
    # ══════════════════════════════════════════════════════════

    def _add_row(self):
        if not self.cleaner:
            return
        try:
            idx = int(self.ent_ridx.get())
        except ValueError:
            idx = None
        self._log_edit(self.cleaner.add_row(idx))
        self._refresh()

    def _del_row(self):
        if not self.cleaner:
            return
        try:
            idx = int(self.ent_ridx.get())
        except ValueError:
            messagebox.showwarning("Input", "Enter a valid row index.")
            return
        self._log_edit(self.cleaner.delete_row(idx))
        self._refresh()

    def _add_col(self):
        if not self.cleaner:
            return
        name = self.ent_cname.get().strip()
        if not name:
            messagebox.showwarning("Input", "Enter a column name.")
            return
        self._log_edit(self.cleaner.add_column(name))
        self._refresh()

    def _del_col(self):
        if not self.cleaner:
            return
        name = self.ent_cname.get().strip()
        if not name:
            messagebox.showwarning("Input", "Enter the column name to delete.")
            return
        if not messagebox.askyesno("Confirm",
                f"Delete column  '{name}' ?\nThis can be undone with Ctrl+Z."):
            return
        self._log_edit(self.cleaner.delete_column(name))
        self._refresh()

    def _ren_col(self):
        if not self.cleaner:
            return
        old = self.ent_cname.get().strip()
        if not old:
            messagebox.showwarning("Input", "Enter the current column name.")
            return
        new = simpledialog.askstring(
            "Rename Column",
            f"New name for  '{old}' :",
            parent=self.root,
        )
        if not new:
            return
        self._log_edit(self.cleaner.rename_column(old, new.strip()))
        self._refresh()

    def _edit_cell(self):
        if not self.cleaner:
            return
        try:
            row = int(self.ent_erow.get())
        except ValueError:
            messagebox.showwarning("Input", "Enter a valid row index.")
            return
        col = self.cb_ecol.get()
        val = self.ent_eval.get()
        self._log_edit(self.cleaner.edit_cell(row, col, val))
        self._refresh()

    def _log_edit(self, msg: str):
        if msg:
            self.edit_log.insert("end", msg + "\n")
            self.edit_log.see("end")

    # ── double-click cell in tree ────────────────────────────────

    def _cell_dbl_click(self, event):
        if not self.cleaner:
            return
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        col_idx  = int(col_id.replace("#", "")) - 1
        cols     = self.cleaner.df.columns.tolist()
        if col_idx >= len(cols):
            return
        col_name = cols[col_idx]
        row_idx  = self.tree.index(row_id)
        cur_val  = self.cleaner.df.iat[row_idx, col_idx]

        new_val = simpledialog.askstring(
            "Edit Cell",
            f"Row {row_idx}  |  Column: {col_name}\n"
            f"Current value:  {cur_val}\n\nNew value:",
            parent=self.root,
        )
        if new_val is None:
            return
        self._log_edit(self.cleaner.edit_cell(row_idx, col_name, new_val))
        self._refresh()
        self.nb.select(self._tabs["preview"])

    # ── right-click context menu ─────────────────────────────────

    def _tree_right_click(self, event):
        if not self.cleaner:
            return
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        row_idx = self.tree.index(row_id)

        menu = tk.Menu(
            self.root, tearoff=0,
            bg=T["panel"], fg=T["text"],
            activebackground=T["select"],
            activeforeground=T["gold"],
            font=("Segoe UI", 9),
        )
        menu.add_command(
            label="➕  Add Row Below",
            command=lambda: (
                self._log_edit(self.cleaner.add_row(idx=row_idx)),
                self._refresh(),
            ),
        )
        menu.add_command(
            label="🗑   Delete This Row",
            command=lambda: (
                self._log_edit(self.cleaner.delete_row(row_idx)),
                self._refresh(),
            ),
        )
        menu.add_separator()
        menu.add_command(
            label="✏️   Edit Cell — double-click",
            state="disabled",
        )
        menu.post(event.x_root, event.y_root)

    # ══════════════════════════════════════════════════════════
    #  UNDO / REDO
    # ══════════════════════════════════════════════════════════

    def undo(self):
        if not self.cleaner:
            return
        msg = self.cleaner.undo()
        if msg:
            self._log_clean(msg)
            self._refresh()
            self.var_status.set(msg)

    def redo(self):
        if not self.cleaner:
            return
        msg = self.cleaner.redo()
        if msg:
            self._log_clean(msg)
            self._refresh()
            self.var_status.set(msg)

    def reset(self):
        if not self.cleaner:
            return
        if not messagebox.askyesno(
            "Reset Data",
            "Reset to the original loaded file?\nAll changes will be lost.",
        ):
            return
        self.cleaner.reset(self.original)
        self.clean_log.delete("1.0", "end")
        self.edit_log.delete("1.0",  "end")
        self._refresh()
        self.var_status.set("↺  Reset to original file")

    # ══════════════════════════════════════════════════════════
    #  PIVOT
    # ══════════════════════════════════════════════════════════

    def _make_pivot(self):
        if not self.cleaner:
            return
        try:
            piv = pd.pivot_table(
                self.cleaner.df,
                index   = self.cb_piv_idx.get(),
                values  = self.cb_piv_val.get(),
                aggfunc = self.cb_piv_agg.get(),
            ).reset_index()
        except Exception as e:
            messagebox.showerror("Pivot Error", str(e))
            return
        self.pivot_tree.delete(*self.pivot_tree.get_children())
        self.pivot_tree["columns"] = list(piv.columns)
        for c in piv.columns:
            self.pivot_tree.heading(c, text=c)
            self.pivot_tree.column(c, width=120, anchor="center")
        for _, row in piv.iterrows():
            self.pivot_tree.insert(
                "", "end",
                values=["" if pd.isna(v) else str(v) for v in row],
            )

    # ══════════════════════════════════════════════════════════
    #  CHART
    # ══════════════════════════════════════════════════════════

    def _show_chart(self):
        if not self.cleaner:
            return
        ct = self.cb_cht_type.get()
        xc = self.cb_cht_x.get()
        yc = self.cb_cht_y.get()
        if not xc:
            return
        try:
            plt.style.use("dark_background")
            fig, ax = plt.subplots(figsize=(10, 6))
            fig.patch.set_facecolor(T["bg"])
            ax.set_facecolor(T["panel"])
            ax.tick_params(colors=T["gold"])
            for spine in ax.spines.values():
                spine.set_color(T["border"])

            df = self.cleaner.df
            if ct == "bar":
                df.groupby(xc, dropna=False)[yc].sum().plot(
                    kind="bar", ax=ax, color=T["gold"], edgecolor=T["bg"])
            elif ct == "line":
                df.groupby(xc, dropna=False)[yc].sum().plot(
                    kind="line", ax=ax, color=T["gold"], marker="o")
            elif ct == "pie":
                df.groupby(xc, dropna=False)[yc].sum().plot(
                    kind="pie", ax=ax, autopct="%1.1f%%")
                ax.set_ylabel("")
            elif ct == "histogram":
                pd.to_numeric(df[yc], errors="coerce").dropna().plot(
                    kind="hist", ax=ax, bins=20,
                    color=T["gold"], edgecolor=T["bg"])
            elif ct == "scatter":
                ax.scatter(
                    pd.to_numeric(df[xc], errors="coerce"),
                    pd.to_numeric(df[yc], errors="coerce"),
                    color=T["gold"], alpha=0.6,
                )
                ax.set_xlabel(xc, color=T["text"])
                ax.set_ylabel(yc, color=T["text"])

            ax.set_title(f"{ct.title()} Chart  —  {yc} by {xc}",
                         color=T["gold_bright"], fontsize=13)
            plt.tight_layout()
            plt.show()
        except Exception as e:
            messagebox.showerror("Chart Error", str(e))

    # ══════════════════════════════════════════════════════════
    #  EXPORT
    # ══════════════════════════════════════════════════════════

    def export_csv(self):
        if not self.cleaner:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
        )
        if path:
            self.cleaner.df.to_csv(path, index=False)
            messagebox.showinfo("Saved", f"CSV saved:\n{path}")

    def export_excel(self):
        if not self.cleaner:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if path:
            self.cleaner.df.to_excel(path, index=False)
            messagebox.showinfo("Saved", f"Excel saved:\n{path}")

    def export_report(self):
        if not self.cleaner:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Report", "*.txt")],
        )
        if not path:
            return
        cmp = self.cleaner.get_comparison()
        b   = cmp["before"]
        a   = cmp["after"]
        now = datetime.now().strftime("%Y-%m-%d  %H:%M:%S")
        w   = 64
        lines = [
            "=" * w,
            "  EXCEL AI CLEANER  —  DATA QUALITY REPORT",
            "  Developer : Sajeeb The Analyst",
            "  Version   : 1.0.0  •  Powered by Groq AI (Llama-3)",
            f"  Generated : {now}",
            "=" * w, "",
            f"FILE: {os.path.basename(self.filepath)}", "",
            "─" * w, "BEFORE vs AFTER CLEANING", "─" * w,
            f"{'Metric':<24} {'Before':>10} {'After':>10} {'Change':>10}",
            f"{'─'*24} {'─'*10} {'─'*10} {'─'*10}",
            f"{'Total Rows':<24} {b['rows']:>10,} {a['rows']:>10,} "
            f"{a['rows']-b['rows']:>+10,}",
            f"{'Missing Cells':<24} {b['miss']:>10,} {a['miss']:>10,} "
            f"{a['miss']-b['miss']:>+10,}",
            f"{'Duplicate Rows':<24} {b['dupe']:>10,} {a['dupe']:>10,} "
            f"{a['dupe']-b['dupe']:>+10,}",
            "", "─" * w, "AI ANALYSIS", "─" * w,
        ]
        ai_text = self.ai_box.get("1.0", "end").strip()
        lines.append(ai_text if ai_text else "No AI scan was run.")
        lines += [
            "", "=" * w,
            "Excel AI Cleaner  v1.0",
            "Developer: Sajeeb The Analyst",
            "Powered by Groq AI (Llama-3)",
            "=" * w,
        ]
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
            messagebox.showinfo("Saved", f"Report saved:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ══════════════════════════════════════════════════════════
    #  SETTINGS ACTIONS
    # ══════════════════════════════════════════════════════════

    def _save_key(self):
        self.cfg["api_key"] = self.var_key.get().strip()
        cfg_save(self.cfg)
        messagebox.showinfo("Saved", "✅  API key saved successfully!")

    def _test_key(self):
        key = self.var_key.get().strip()
        if not key:
            messagebox.showwarning("No Key", "Enter your Groq API key first.")
            return
        self._busy(True, "Testing connection…")

        def _test():
            try:
                from groq import Groq
                Groq(api_key=key).chat.completions.create(
                    model    = "llama-3.3-70b-versatile",
                    messages = [{"role": "user", "content": "Say OK"}],
                    max_tokens=5,
                )
                self.root.after(0, lambda: (
                    messagebox.showinfo(
                        "Connected",
                        "✅  Groq API key is valid!\n"
                        "AI is ready to scan your data.",
                    ),
                    self._busy(False, "✅  API key OK"),
                ))
            except ImportError:
                self.root.after(0, lambda: (
                    messagebox.showerror(
                        "Package Missing",
                        "Run this in your VS Code terminal:\n\n    pip install groq",
                    ),
                    self._busy(False, ""),
                ))
            except Exception as e:
                self.root.after(0, lambda err=e: (
                    messagebox.showerror("Connection Failed", str(err)),
                    self._busy(False, "❌  Key failed"),
                ))

        threading.Thread(target=_test, daemon=True).start()

    # ══════════════════════════════════════════════════════════
    #  HELPERS
    # ══════════════════════════════════════════════════════════

    def _busy(self, on: bool, status: str = ""):
        def _do():
            self.var_status.set(status)
            if on:
                self.pbar.start(10)
            else:
                self.pbar.stop()
        self.root.after(0, _do)

    def _on_close(self):
        try:
            geo = self.root.geometry()
            m   = re.match(r"(\d+)x(\d+)", geo)
            if m:
                self.cfg["window_w"] = int(m.group(1))
                self.cfg["window_h"] = int(m.group(2))
            cfg_save(self.cfg)
        except Exception:
            pass
        self.root.destroy()


# ══════════════════════════════════════════════════════════════════
#  ENTRY POINT  —  Splash → App
# ══════════════════════════════════════════════════════════════════

def main():
    splash_root = tk.Tk()

    def launch():
        splash_root.destroy()
        app_root = tk.Tk()
        App(app_root)
        app_root.mainloop()

    Splash(splash_root, on_done=launch)
    splash_root.mainloop()


if __name__ == "__main__":
    main()