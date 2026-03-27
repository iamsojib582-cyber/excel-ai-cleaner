
"""
Excel AI Cleaner - by Sajeeb The Analyst
Deep Purple & Gold • Advanced • Single File

SETUP (run once):
pip install groq pandas openpyxl matplotlib numpy

FREE GROQ API KEY:
1. https://console.groq.com 
2. Sign up (free)
3. API Keys → Create Key
4. Paste in app → Settings tab

RUN: python Excel_Ai_Data_Cleaning_ADVANCED.py
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
    "cond_cell"  : "#2a1f0e",
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

_CFG_DIR  = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "ExcelAICleaner")
_CFG_PATH = os.path.join(_CFG_DIR, "config.json")
_LEGACY_CFG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "excel_ai_config.json")
_CFG_DEF  = {"api_key": "", "last_folder": "", "window_w": 1420, "window_h": 840}


def cfg_load() -> dict:
    # Prefer user profile config, fall back to legacy file (portable build)
    try:
        with open(_CFG_PATH) as f:
            return {**_CFG_DEF, **json.load(f)}
    except Exception:
        pass
    try:
        with open(_LEGACY_CFG_PATH) as f:
            data = {**_CFG_DEF, **json.load(f)}
        cfg_save(data)
        return data
    except Exception:
        return _CFG_DEF.copy()


def cfg_save(d: dict) -> None:
    try:
        os.makedirs(_CFG_DIR, exist_ok=True)
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

# Large file handling (keep UI responsive for millions of rows)
_BIG_FILE_MB = 60
_BIG_SAMPLE_BYTES = 1024 * 1024


def _guess_sep_from_sample(sample: str) -> str:
    sample = "\n".join(sample.splitlines()[:50])
    return max([";", ",", "\t", "|"], key=sample.count)


def _read_csv_pandas(path: str) -> pd.DataFrame:
    # Try fast pandas read with multiple encodings
    encs = ["utf-8", "utf-8-sig", "latin-1", "cp1252"]
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        sample = f.read(_BIG_SAMPLE_BYTES)
    sep = _guess_sep_from_sample(sample)
    last_err = None
    for enc in encs:
        try:
            return pd.read_csv(path, sep=sep, encoding=enc, low_memory=False)
        except Exception as e:
            last_err = e
            continue
    if last_err:
        raise last_err
    raise ValueError("CSV read failed.")


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

    if ext in (".parquet", ".pq"):
        try:
            return _norm(pd.read_parquet(path))
        except Exception as e:
            raise ValueError(f"Parquet read failed: {e}")

    # CSV / TSV / TXT
    if ext in (".tsv", ".txt"):
        ext = ".csv"

    # Fast path for big CSVs
    try:
        if ext == ".csv" and os.path.getsize(path) / (1024 * 1024) >= _BIG_FILE_MB:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                sample = f.read(_BIG_SAMPLE_BYTES)
            sep = _guess_sep_from_sample(sample)
            df = pd.read_csv(path, sep=sep, low_memory=False)
            return _norm(df)
    except Exception:
        # fallback to manual parsing below
        pass

    # CSV (robust parsing)
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


def load_file_sheets(path: str) -> dict:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        try:
            xls = pd.ExcelFile(path)
            sheets = {}
            for name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=name)
                sheets[name] = _norm(df)
            if not sheets:
                raise ValueError("Excel file has no sheets.")
            return sheets
        except Exception:
            # Fallback to single-sheet read
            return {"Sheet1": load_file(path)}
    return {"Sheet1": load_file(path)}


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
                    nn, errors="coerce"
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

Format your response EXACTLY like this - one block per issue:

ISSUE: [short clear title]
DETAIL: [specific description - use exact column names]
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
    extra_prompt: str = "",
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
            user_msg = f"Dataset profile:\n\n{profile}"
            if extra_prompt.strip():
                user_msg += f"\n\nUser instructions:\n{extra_prompt.strip()}"
            stream = client.chat.completions.create(
                model    = "llama-3.3-70b-versatile",
                messages = [
                    {"role": "system", "content": _AI_SYSTEM},
                    {"role": "user",   "content": user_msg},
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

    def clean_signs_all(self) -> str:
        b = self.df.copy()
        for c in self.df.select_dtypes(include=["object", "string"]).columns:
            self.df[c] = (
                self.df[c]
                .astype("string")
                .str.replace(r"[\\_\\:,'\\+=]", "", regex=True)
                .str.replace(r"\\s+", " ", regex=True)
                .str.strip()
            )
        return self._commit("Clean Symbols", b,
                            "🧽  Removed common symbols from all text columns.")

    def clean_signs_column(self, col: str) -> str:
        if col not in self.df.columns:
            return f"❌  Column `{col}` not found."
        b = self.df.copy()
        self.df[col] = (
            self.df[col]
            .astype("string")
            .str.replace(r"[\\_\\:,'\\+=]", "", regex=True)
            .str.replace(r"\\s+", " ", regex=True)
            .str.strip()
        )
        return self._commit("Clean Symbols (Column)", b,
                            f"🧽  Removed common symbols in `{col}`.")

    def trim_whitespace_column(self, col: str) -> str:
        if col not in self.df.columns:
            return f"❌  Column `{col}` not found."
        b = self.df.copy()
        self.df[col] = self.df[col].astype("string").str.strip()
        return self._commit("Trim Whitespace (Column)", b,
                            f"✂️   Trimmed whitespace in `{col}`.")


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

    def set_outside_range(self, col: str, min_v=None, max_v=None, fill="") -> str:
        if col not in self.df.columns:
            return f"❌  Column `{col}` not found."
        b = self.df.copy()
        s = pd.to_numeric(self.df[col], errors="coerce")
        mask = pd.Series(False, index=self.df.index)
        if min_v is not None:
            mask |= s < float(min_v)
        if max_v is not None:
            mask |= s > float(max_v)
        if fill in ("", None):
            self.df.loc[mask, col] = pd.NA
        else:
            self.df.loc[mask, col] = fill
        return self._commit(
            "Set Outside Range",
            b,
            f"🎯  Set out-of-range values in `{col}` to '{fill}'.",
        )

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

    def fix_date_columns_date_only(self) -> str:
        b     = self.df.copy()
        fixed = []
        for c in self.df.select_dtypes(include=["object", "string"]).columns:
            try:
                p = pd.to_datetime(
                    self.df[c].astype(str),
                    errors="coerce", infer_datetime_format=True,
                )
                if p.notna().mean() > 0.70:
                    self.df[c] = p.dt.date
                    fixed.append(c)
            except Exception:
                pass
        msg = (f"📅  Converted to date-only: {', '.join(f'`{c}`' for c in fixed)}."
               if fixed else "📅  No text-as-date columns found.")
        return self._commit("Fix Dates (Date Only)", b, msg)


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

    def auto_detect_column_type(self, col: str) -> str:
        if col not in self.df.columns:
            return f"❌  Column `{col}` not found."
        b = self.df.copy()
        s = self.df[col].astype(str)
        num_rate = pd.to_numeric(s.str.replace(r"[$,\\s%€£¥]", "", regex=True), errors="coerce").notna().mean()
        date_rate = pd.to_datetime(s, errors="coerce").notna().mean()
        if num_rate >= 0.8:
            self.df[col] = pd.to_numeric(s.str.replace(r"[$,\\s%€£¥]", "", regex=True), errors="coerce")
            return self._commit("Auto Detect Type", b, f"🔎  `{col}` detected as numeric.")
        if date_rate >= 0.7:
            self.df[col] = pd.to_datetime(s, errors="coerce")
            return self._commit("Auto Detect Type", b, f"🔎  `{col}` detected as date.")
        self.df[col] = s
        return self._commit("Auto Detect Type", b, f"🔎  `{col}` kept as text.")

    def convert_type(self, col: str, to_type: str, date_only: bool = False) -> str:
        if col not in self.df.columns:
            return f"❌  Column `{col}` not found."
        b = self.df.copy()
        to_type = to_type.lower()
        try:
            if to_type in ("number", "numeric", "float", "int"):
                self.df[col] = pd.to_numeric(self.df[col], errors="coerce")
            elif to_type in ("date", "datetime", "date_only"):
                series = pd.to_datetime(self.df[col], errors="coerce")
                if to_type in ("date", "date_only") or date_only:
                    series = series.dt.date
                self.df[col] = series
            elif to_type in ("text", "string"):
                self.df[col] = self.df[col].astype(str)
            else:
                return f"❌  Unsupported type: {to_type}"
        except Exception:
            return f"❌  Failed to convert `{col}` to {to_type}"
        suffix = " (date only)" if (to_type in ("date", "date_only") or date_only) else ""
        return self._commit("Convert Type", b, f"🔁  Converted `{col}` to {to_type}{suffix}.")

    def add_column_with_value(self, name: str, value) -> str:
        if name in self.df.columns:
            return f"❌  Column `{name}` already exists."
        b = self.df.copy()
        self.df[name] = value
        return self._commit("Add Column", b, f"➕  Added column `{name}` (default={value}).")

    def drop_columns(self, cols: list) -> str:
        present = [c for c in cols if c in self.df.columns]
        if not present:
            return "❌  No matching columns found."
        b = self.df.copy()
        self.df = self.df.drop(columns=present)
        msg = "🗑️  Deleted columns: " + ", ".join(present)
        return self._commit("Delete Columns", b, msg)

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

    def add_row_values(self, values: dict, idx: int | None = None) -> str:
        b = self.df.copy()
        row = {c: pd.NA for c in self.df.columns}
        for k, v in (values or {}).items():
            if k in row:
                row[k] = v
        blank = pd.DataFrame([row], columns=self.df.columns)
        if idx is None or idx >= len(self.df) - 1:
            self.df = pd.concat([self.df, blank], ignore_index=True)
            pos = "bottom"
        else:
            self.df = pd.concat(
                [self.df.iloc[: idx + 1], blank, self.df.iloc[idx + 1 :]],
                ignore_index=True,
            )
            pos = f"after row {idx}"
        return self._commit("Add Row Values", b, f"➕  Added 1 row with values ({pos}).")


    def delete_row(self, idx: int) -> str:
        if idx < 0 or idx >= len(self.df):
            return "❌  Row index out of range."
        b = self.df.copy()
        self.df = self.df.drop(index=idx).reset_index(drop=True)
        return self._commit("Delete Row", b, f"🗑️   Deleted row {idx}.")

    def duplicate_row(self, idx: int) -> str:
        if idx < 0 or idx >= len(self.df):
            return "❌  Row index out of range."
        b = self.df.copy()
        row = self.df.iloc[[idx]].copy()
        self.df = pd.concat(
            [self.df.iloc[: idx + 1], row, self.df.iloc[idx + 1 :]],
            ignore_index=True,
        )
        return self._commit("Duplicate Row", b, f"🧬  Duplicated row {idx}.")


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
                return f"⚠️  `{col}` is numeric - '{val}' is not a valid number."
        self.df.at[row, col] = val
        self.modified.add((row, col))
        return self._commit("Edit Cell", b,
                            f"✏️   [{row}, `{col}`]:  '{old}'  →  '{val}'.")

    def paste_block(self, start_row: int, start_col: int, block: list[list[str]]) -> str:
        if start_row < 0 or start_col < 0:
            return "❌  Invalid paste position."
        if not block:
            return "❌  Nothing to paste."
        b = self.df.copy()
        cols = self.df.columns.tolist()
        needed_rows = start_row + len(block)
        if needed_rows > len(self.df):
            add = needed_rows - len(self.df)
            extra = pd.DataFrame([[pd.NA] * len(cols)] * add, columns=cols)
            self.df = pd.concat([self.df, extra], ignore_index=True)
        max_cols = max(len(r) for r in block)
        for r, row_vals in enumerate(block):
            for c, val in enumerate(row_vals):
                col_idx = start_col + c
                if col_idx >= len(cols):
                    break
                col = cols[col_idx]
                v = val
                if pd.api.types.is_numeric_dtype(self.df[col]):
                    try:
                        v = float(v) if "." in str(v) else int(v)
                    except Exception:
                        pass
                self.df.at[start_row + r, col] = v
                self.modified.add((start_row + r, col))
        return self._commit("Paste", b, f"📋  Pasted {len(block)}x{max_cols} cells.")


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
        self.sheets   : dict           = {}
        self.originals: dict           = {}
        self.current_sheet: str        = ""
        self.filepath : str            = ""
        self.ai_cards : list           = []
        self._card_widgets: list       = []
        self.cond_rule: dict | None    = None
        self.active_filters: dict      = {}
        self.sort_state: dict          = {"col": None, "asc": True}
        self.view_index: list          = []
        self._sel_start: tuple | None  = None
        self._sel_end: tuple | None    = None

        # ── window setup
        w  = self.cfg.get("window_w", 1420)
        h  = self.cfg.get("window_h", 840)
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        root.title("Excel AI Cleaner - by Sajeeb The Analyst")
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

        ttk.Label(bar, text="Sheet:", style="Dim.TLabel").pack(side="left", padx=(6, 4))
        self.cb_sheet = ttk.Combobox(bar, state="readonly", width=16)
        self.cb_sheet.pack(side="left", padx=3)
        self.cb_sheet.bind("<<ComboboxSelected>>", self._on_sheet_select)
        ttk.Button(bar, text="➕  Add Sheet",
                   command=self._add_sheet).pack(side="left", padx=3)
        ttk.Button(bar, text="🗑   Delete Sheet",
                   command=self._delete_sheet).pack(side="left", padx=3)
        ttk.Button(bar, text="✏️  Rename Sheet",
                   command=self._rename_sheet).pack(side="left", padx=3)

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

        self.tree = ttk.Treeview(tf, show="headings", selectmode="extended")
        self.tree.pack(side="left", fill="both", expand=True)

        sy = ttk.Scrollbar(tf, orient="vertical",   command=self.tree.yview)
        sy.pack(side="right", fill="y")
        sx = ttk.Scrollbar(p,  orient="horizontal", command=self.tree.xview)
        sx.pack(fill="x", padx=8)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)

        # interactions
        self.tree.bind("<Double-1>",  self._cell_dbl_click)
        self.tree.bind("<Button-3>",  self._tree_right_click)
        self.tree.bind("<Button-1>",  self._tree_left_down)
        self.tree.bind("<B1-Motion>", self._tree_drag)

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

        # optional prompt
        pr = tk.Frame(p, bg=T["bg"])
        pr.pack(fill="x", padx=12, pady=(0, 6))
        tk.Label(
            pr, text="AI Instruction (optional) — tell the AI what to focus on",
            bg=T["bg"], fg=T["dim"], font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w")
        self.ai_prompt = tk.Text(
            pr, height=3, wrap="word", font=("Consolas", 9),
            bg=T["panel"], fg=T["text"],
            relief="flat", bd=0, padx=8, pady=6,
            insertbackground=T["gold"],
        )
        self.ai_prompt.pack(fill="x", pady=4)

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
            ("Clean Symbols",      "clean_signs_all"),
            ("Fix Dates",          "fix_date_columns"),
            ("Fix Dates (Date Only)", "fix_date_columns_date_only"),
            ("Fix Numbers",        "fix_numeric_columns"),
            ("Title Case",         "standardise_case"),
            ("Junk Rows",          "remove_junk_rows"),
            ("Fix Col Names",      "standardise_col_names"),
        ]:
            ttk.Button(
                r3, text=label,
                command=lambda m=method: self._run_clean(m),
            ).pack(side="left", padx=3)

        # Row 4 — AI prompt clean
        self._lbl_section(f, "AI Prompt Clean")
        r4 = tk.Frame(f, bg=T["bg"]); r4.pack(fill="x", pady=5)
        self.ai_clean_prompt = tk.Text(
            r4, height=3, wrap="word", font=("Consolas", 9),
            bg=T["panel"], fg=T["text"],
            relief="flat", bd=0, padx=8, pady=6,
            insertbackground=T["gold"],
        )
        self.ai_clean_prompt.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(
            r4, text="🤖  Run AI Prompt",
            style="Gold.TButton",
            command=self._ai_prompt_clean,
        ).pack(side="left")

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
        ttk.Button(r1, text="Data Entry Form", command=self._open_row_entry_dialog).pack(side="left", padx=6)


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

        # Conditional formatting (preview highlight)
        self._lbl_section(f, "Conditional Highlight")
        r4 = tk.Frame(f, bg=T["bg"]); r4.pack(fill="x", pady=5)
        self._lbl(r4, "Column:")
        self.cb_cond_col = ttk.Combobox(r4, state="readonly", width=18)
        self.cb_cond_col.pack(side="left", padx=4)
        self._lbl(r4, "Rule:")
        self.cb_cond_op = ttk.Combobox(
            r4, state="readonly", width=10,
            values=["=", "!=", ">", "<", ">=", "<=", "contains", "starts_with", "ends_with"],
        )
        self.cb_cond_op.set("="); self.cb_cond_op.pack(side="left", padx=4)
        self._lbl(r4, "Value:")
        self.ent_cond_val = ttk.Entry(r4, width=16)
        self.ent_cond_val.pack(side="left", padx=4)
        ttk.Button(r4, text="Apply Highlight", command=self._apply_cond_format).pack(side="left", padx=4)
        ttk.Button(r4, text="Clear", command=self._clear_cond_format).pack(side="left", padx=4)

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
        self._lbl(ctrl, "Aggregation")
        self.cb_cht_agg = ttk.Combobox(
            ctrl, state="readonly", width=10,
            values=["sum", "mean", "count", "min", "max"],
        )
        self.cb_cht_agg.set("sum"); self.cb_cht_agg.pack(side="left", padx=4)
        ttk.Button(ctrl, text="Show Chart",
                   style="Gold.TButton",
                   command=self._show_chart).pack(side="left", padx=10)

        tk.Label(
            p,
            text="Tip: select a text column as X (category), a numeric column as Y, and choose an aggregation.",
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
            "7.  Use the Sheet selector to add, delete, or rename sheets.\n"
            "8.  Conditional Highlight lets you preview rules like >, <, contains.\n"
            "9.  Full Undo / Redo history — Ctrl+Z  /  Ctrl+Shift+Z.\n"
            "10. Export cleaned data as CSV or Excel.  Generate a text Quality Report.\n\n"
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

    def _cond_match(self, row_idx: int, df_view: pd.DataFrame) -> bool:
        if not self.cond_rule:
            return False
        col = self.cond_rule.get("col")
        op  = self.cond_rule.get("op")
        val = self.cond_rule.get("val")
        if not col or col not in df_view.columns:
            return False
        cell = df_view.iloc[row_idx][col]
        if pd.isna(cell):
            return False

        cell_s = str(cell)
        val_s  = str(val)

        if op in {">", "<", ">=", "<="}:
            try:
                cell_n = float(cell)
                val_n  = float(val)
            except Exception:
                return False
            if op == ">":  return cell_n > val_n
            if op == "<":  return cell_n < val_n
            if op == ">=": return cell_n >= val_n
            if op == "<=": return cell_n <= val_n
        if op == "=":
            return cell_s == val_s
        if op == "!=":
            return cell_s != val_s
        if op == "contains":
            return val_s.lower() in cell_s.lower()
        if op == "starts_with":
            return cell_s.lower().startswith(val_s.lower())
        if op == "ends_with":
            return cell_s.lower().endswith(val_s.lower())
        return False

    # ══════════════════════════════════════════════════════════
    #  KEYBOARD SHORTCUTS
    # ══════════════════════════════════════════════════════════

    def _bind_shortcuts(self):
        r = self.root
        r.bind("<Control-o>",       lambda e: self.open_file())
        r.bind("<Control-s>",       lambda e: self.export_csv())
        r.bind("<Control-c>",       lambda e: self._copy_selection())
        r.bind("<Control-v>",       lambda e: self._paste_selection())
        r.bind("<Control-j>",       lambda e: self.export_json())
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
                ("Data Files", "*.csv *.tsv *.txt *.xlsx *.xls *.json *.parquet *.pq"),
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
            sheets = load_file_sheets(path)
        except Exception as e:
            self.root.after(0, lambda: (
                messagebox.showerror("Load Error", str(e)),
                self._busy(False, "Load failed"),
            ))
            return
        self.root.after(0, lambda: self._finish_load(path, sheets))

    def _finish_load(self, path: str, sheets: dict):
        self.filepath = path
        self.originals = {k: v.copy() for k, v in sheets.items()}
        self.sheets    = {k: Cleaner(v) for k, v in sheets.items()}
        self.current_sheet = next(iter(self.sheets)) if self.sheets else ""
        self.cleaner = self.sheets.get(self.current_sheet)
        self.original = self.originals.get(self.current_sheet, pd.DataFrame())
        self.cond_rule = None
        self.cfg["last_folder"] = os.path.dirname(path)
        cfg_save(self.cfg)
        self.lbl_file.config(
            text=os.path.basename(path),
            foreground=T["gold"],
        )
        self._update_sheet_list()
        if self.current_sheet:
            self._set_current_sheet(self.current_sheet)
        else:
            self._refresh()
        if self.cleaner:
            df = self.cleaner.df
            self._busy(False, f"Loaded  {len(df):,} rows × {len(df.columns)} columns")
            if len(df) >= 1_000_000:
                self.var_status.set("???  Large file mode: preview limited to 500 rows for speed.")
        else:
            self._busy(False, "Loaded file")

    # ══════════════════════════════════════════════════════════
    #  REFRESH
    # ══════════════════════════════════════════════════════════

    def _refresh(self):
        self._update_summary()
        self._fill_tree()
        self._update_combos()

    def _update_sheet_list(self):
        names = list(self.sheets.keys()) if self.sheets else []
        self.cb_sheet["values"] = names
        if names:
            if self.current_sheet not in names:
                self.current_sheet = names[0]
            self.cb_sheet.set(self.current_sheet)
        else:
            self.cb_sheet.set("")

    def _set_current_sheet(self, name: str):
        if name not in self.sheets:
            return
        self.current_sheet = name
        self.cleaner = self.sheets[name]
        self.original = self.originals.get(name, pd.DataFrame())
        base = os.path.basename(self.filepath) if self.filepath else "No file opened"
        self.lbl_file.config(
            text=f"{base}  |  Sheet: {name}",
            foreground=T["gold"] if self.filepath else T["dim"],
        )
        self._refresh()

    def _on_sheet_select(self, _event=None):
        name = self.cb_sheet.get()
        if name:
            self._set_current_sheet(name)

    def _add_sheet(self):
        if not self.sheets:
            return
        name = simpledialog.askstring("Add Sheet", "New sheet name:", parent=self.root)
        if not name:
            return
        if name in self.sheets:
            messagebox.showwarning("Sheet Exists", "A sheet with that name already exists.")
            return
        copy_cols = messagebox.askyesno(
            "Copy Columns?",
            "Copy columns from the current sheet?",
        )
        if copy_cols and self.cleaner:
            df = pd.DataFrame(columns=self.cleaner.df.columns)
        else:
            df = pd.DataFrame(columns=["Column1"])
        self.sheets[name] = Cleaner(df)
        self.originals[name] = df.copy()
        self._update_sheet_list()
        self._set_current_sheet(name)
        self.var_status.set(f"➕  Added sheet '{name}'")

    def _delete_sheet(self):
        if not self.sheets or len(self.sheets) <= 1:
            messagebox.showwarning("Not Allowed", "At least one sheet must remain.")
            return
        name = self.current_sheet
        if not messagebox.askyesno(
            "Delete Sheet",
            f"Delete sheet '{name}'?\nThis cannot be undone.",
        ):
            return
        self.sheets.pop(name, None)
        self.originals.pop(name, None)
        self.current_sheet = next(iter(self.sheets))
        self._update_sheet_list()
        self._set_current_sheet(self.current_sheet)
        self.var_status.set(f"🗑  Deleted sheet '{name}'")

    def _rename_sheet(self):
        if not self.sheets:
            return
        old = self.current_sheet
        new = simpledialog.askstring("Rename Sheet", f"New name for '{old}':", parent=self.root)
        if not new:
            return
        if new in self.sheets:
            messagebox.showwarning("Sheet Exists", "A sheet with that name already exists.")
            return
        self.sheets[new] = self.sheets.pop(old)
        self.originals[new] = self.originals.pop(old)
        self.current_sheet = new
        self._update_sheet_list()
        self._set_current_sheet(new)
        self.var_status.set(f"✏️  Renamed sheet to '{new}'")

    def _update_summary(self):
        if not self.cleaner or self.cleaner.df.empty:
            self.lbl_sum.config(
                text="Rows: ???  |  Columns: ???  |  Missing: ???  |  Duplicates: ???"
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

    def _get_view_df(self) -> pd.DataFrame | None:
        if not self.cleaner:
            return None
        df_view = self.cleaner.df
        df_view = self._apply_filters(df_view)
        col = self.sort_state.get("col")
        if col in df_view.columns:
            try:
                df_view = df_view.sort_values(by=col, ascending=bool(self.sort_state.get("asc", True)), kind="mergesort")
            except Exception:
                pass
        return df_view

    def _apply_filters(self, df: pd.DataFrame) -> pd.DataFrame:
        if not self.active_filters:
            return df
        out = df
        for col, rule in list(self.active_filters.items()):
            if col not in out.columns:
                continue
            op = (rule or {}).get("op", "contains")
            val = (rule or {}).get("value", "")
            try:
                if op == "contains":
                    mask = out[col].astype(str).str.contains(str(val), case=False, na=False)
                elif op == "equals":
                    mask = out[col].astype(str) == str(val)
                elif op == "starts_with":
                    mask = out[col].astype(str).str.startswith(str(val), na=False)
                elif op == "ends_with":
                    mask = out[col].astype(str).str.endswith(str(val), na=False)
                elif op in (">", ">=", "<", "<="):
                    s = pd.to_numeric(out[col], errors="coerce")
                    v = pd.to_numeric(pd.Series([val]), errors="coerce").iloc[0]
                    if pd.isna(v):
                        continue
                    if op == ">":
                        mask = s > v
                    elif op == ">=":
                        mask = s >= v
                    elif op == "<":
                        mask = s < v
                    else:
                        mask = s <= v
                elif op == "is_blank":
                    mask = out[col].isna() | (out[col].astype(str).str.strip() == "")
                elif op == "not_blank":
                    mask = ~(out[col].isna() | (out[col].astype(str).str.strip() == ""))
                else:
                    mask = out[col].astype(str).str.contains(str(val), case=False, na=False)
                out = out[mask]
            except Exception:
                continue
        return out

    def _map_view_row(self, view_row: int) -> int:
        if self.view_index and 0 <= view_row < len(self.view_index):
            return int(self.view_index[view_row])
        return view_row

    def _sort_column(self, col: str, asc: bool = True):
        self.sort_state = {"col": col, "asc": bool(asc)}
        self._refresh()

    def _clear_filters(self):
        self.active_filters.clear()
        self._refresh()

    def _filter_column_dialog(self, col: str):
        if not self.cleaner:
            return
        top = tk.Toplevel(self.root)
        top.title(f"Filter Column: {col}")
        top.configure(bg=T["bg"])
        top.geometry("380x420")

        tk.Label(top, text=f"Column: {col}", bg=T["bg"], fg=T["gold_bright"],
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=12, pady=(10, 6))

        row = tk.Frame(top, bg=T["bg"]); row.pack(fill="x", padx=12, pady=6)
        tk.Label(row, text="Rule:", bg=T["bg"], fg=T["text"], font=("Segoe UI", 9)).pack(side="left")
        op_box = ttk.Combobox(
            row, state="readonly", width=14,
            values=["contains", "equals", "starts_with", "ends_with", ">", ">=", "<", "<=", "is blank", "not blank"],
        )
        op_box.set("contains"); op_box.pack(side="left", padx=6)

        val_row = tk.Frame(top, bg=T["bg"]); val_row.pack(fill="x", padx=12, pady=6)
        tk.Label(val_row, text="Value:", bg=T["bg"], fg=T["text"], font=("Segoe UI", 9)).pack(side="left")
        val_ent = ttk.Entry(val_row, width=22)
        val_ent.pack(side="left", padx=6)

        tk.Label(top, text="Quick pick (unique values)", bg=T["bg"], fg=T["dim"],
                 font=("Segoe UI", 9)).pack(anchor="w", padx=12, pady=(8, 2))
        lb_frame = tk.Frame(top, bg=T["bg"])
        lb_frame.pack(fill="both", expand=True, padx=12, pady=4)
        lb = tk.Listbox(lb_frame, height=8, selectmode="browse")
        lb.pack(side="left", fill="both", expand=True)
        sc = ttk.Scrollbar(lb_frame, orient="vertical", command=lb.yview)
        sc.pack(side="right", fill="y")
        lb.configure(yscrollcommand=sc.set)

        try:
            vals = self.cleaner.df[col].dropna().astype(str).unique().tolist()
        except Exception:
            vals = []
        for v in vals[:200]:
            lb.insert("end", v)

        def _pick(_):
            sel = lb.curselection()
            if sel:
                val_ent.delete(0, "end")
                val_ent.insert(0, lb.get(sel[0]))
        lb.bind("<<ListboxSelect>>", _pick)

        btn_row = tk.Frame(top, bg=T["bg"]); btn_row.pack(fill="x", padx=12, pady=10)

        def _apply():
            op = op_box.get().strip().lower()
            val = val_ent.get().strip()
            op_key = op.replace(" ", "_")
            if op_key in ("is_blank", "not_blank"):
                val = ""
            if op_key == "":
                return
            self.active_filters[col] = {"op": op_key, "value": val}
            self._refresh()
            top.destroy()

        def _clear():
            if col in self.active_filters:
                self.active_filters.pop(col, None)
                self._refresh()
            top.destroy()

        ttk.Button(btn_row, text="Apply", style="Gold.TButton", command=_apply).pack(side="right")
        ttk.Button(btn_row, text="Clear Column", command=_clear).pack(side="right", padx=6)
        ttk.Button(btn_row, text="Clear All", command=lambda: (self._clear_filters(), top.destroy())).pack(side="right", padx=6)

    def _tree_left_down(self, event):
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        try:
            col_idx = int(col_id.replace("#", "")) - 1
        except Exception:
            return
        view_row_idx = self.tree.index(row_id)
        self._sel_start = (view_row_idx, col_idx)
        self._sel_end = (view_row_idx, col_idx)
        self._refresh()

    def _tree_drag(self, event):
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        try:
            col_idx = int(col_id.replace("#", "")) - 1
        except Exception:
            return
        view_row_idx = self.tree.index(row_id)
        self._sel_end = (view_row_idx, col_idx)
        self._refresh()

    def _copy_selection(self):
        if not self.cleaner or not self._sel_start or not self._sel_end:
            return
        r1, c1 = self._sel_start
        r2, c2 = self._sel_end
        r_start, r_end = sorted([r1, r2])
        c_start, c_end = sorted([c1, c2])
        df_view = self._get_view_df() if self._get_view_df() is not None else self.cleaner.df
        cols = df_view.columns.tolist()
        rows = []
        for vr in range(r_start, min(r_end + 1, len(df_view))):
            orow = self._map_view_row(vr)
            row_vals = []
            for ci in range(c_start, min(c_end + 1, len(cols))):
                col = cols[ci]
                try:
                    v = self.cleaner.df.at[orow, col]
                except Exception:
                    v = ""
                row_vals.append("" if pd.isna(v) else str(v))
            rows.append("\\t".join(row_vals))
        text = "\\n".join(rows)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)

    def _paste_selection(self):
        if not self.cleaner or not self._sel_start:
            return
        try:
            data = self.root.clipboard_get()
        except Exception:
            return
        if not data:
            return
        lines = [l for l in data.splitlines() if l.strip() != ""]
        if not lines:
            return
        block = [l.split("\\t") for l in lines]
        r0, c0 = self._sel_start
        start_row = self._map_view_row(r0)
        msg = self.cleaner.paste_block(start_row, c0, block)
        self._log_edit(msg)
        self._refresh()

    def _fill_tree(self):
        self.tree.delete(*self.tree.get_children())
        if not self.cleaner:
            return
        df_view = self._get_view_df() if self._get_view_df() is not None else self.cleaner.df
        self.view_index = df_view.index.tolist()
        df = df_view.head(500)
        raw_cols = [str(c) for c in df.columns]
        # Treeview needs unique column identifiers
        seen = {}
        tv_cols = []
        for c in raw_cols:
            if c in seen:
                seen[c] += 1
                tv_cols.append(f"{c}__{seen[c]}")
            else:
                seen[c] = 0
                tv_cols.append(c)
        self.tree["columns"] = tv_cols
        for tv, label in zip(tv_cols, raw_cols):
            self.tree.heading(tv, text=label)
            self.tree.column(tv, width=120, anchor="center", minwidth=50)
        self.tree.tag_configure(
            "modified",
            background=T["gold_cell"],
            foreground=T["gold_bright"],
        )
        self.tree.tag_configure(
            "cond",
            background=T["cond_cell"],
            foreground=T["gold_bright"],
        )
        self.tree.tag_configure(
            "range",
            background=T["select"],
            foreground=T["gold_bright"],
        )

        sel_start = self._sel_start[0] if self._sel_start else None
        sel_end = self._sel_end[0] if self._sel_end else None
        if sel_start is not None and sel_end is not None:
            r1, r2 = sorted([sel_start, sel_end])
        else:
            r1 = r2 = None

        for i, (_, row) in enumerate(df.iterrows()):
            vals = ["" if pd.isna(v) else str(v) for v in row]
            tags = []
            if i < len(self.view_index):
                orig_idx = self.view_index[i]
            else:
                orig_idx = i
            if any((orig_idx, c) in self.cleaner.modified for c in df.columns):
                tags.append("modified")
            if self.cond_rule and self._cond_match(i, df):
                tags.append("cond")
            if r1 is not None and r2 is not None and r1 <= i <= r2:
                tags.append("range")
            self.tree.insert("", "end", values=vals, tags=tuple(tags))

    def _update_combos(self):
        cols = [] if not self.cleaner else self.cleaner.df.columns.tolist()
        nums = [] if not self.cleaner else                self.cleaner.df.select_dtypes(include="number").columns.tolist()

        for cb in [self.cb_col, self.cb_fill_col, self.cb_ecol,
                   self.cb_piv_idx, self.cb_cht_x, self.cb_cond_col]:
            cb["values"] = cols
            if cols: cb.set(cols[0])

        for cb in [self.cb_piv_val, self.cb_cht_y]:
            # Allow all columns for Y, but prefer numeric if available
            cb["values"] = (nums + [c for c in cols if c not in nums]) if cols else []
            if cols:
                cb.set((nums + [c for c in cols if c not in nums])[0])

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
        if self.current_sheet:
            self._ai_write(f"Sheet: {self.current_sheet}\n", "dim")
        self._ai_write(
            f"Size : {len(self.cleaner.df):,} rows × "
            f"{len(self.cleaner.df.columns)} columns\n\n", "dim",
        )
        self.nb.select(self._tabs["ai"])
        self._busy(True, "AI is thinking…")

        run_groq_scan(
            df        = self.cleaner.df,
            api_key   = key,
            extra_prompt = self.ai_prompt.get("1.0", "end").strip() if hasattr(self, "ai_prompt") else "",
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

    

    def _confirm_ai_actions(self, actions: list) -> bool:
        if not actions:
            return False
        lines = ["AI will apply these actions:"]
        for a in actions:
            if not isinstance(a, dict):
                continue
            action = a.get("action", "")
            col = a.get("column", "")
            detail = ""
            if action == "replace":
                detail = f"{col} : {a.get('find','')} -> {a.get('replace','')}"
            elif action == "fill_missing":
                detail = f"{col} : method={a.get('method','')} value={a.get('value','')}"
            elif action == "convert_type":
                detail = f"{col} -> {a.get('type','')}"
            elif action == "set_outside_range":
                detail = f"{col} : min={a.get('min',None)} max={a.get('max',None)} fill={a.get('fill','')}"
            elif action == "drop_columns":
                detail = f"{', '.join(a.get('columns',[]))}"
            elif action == "rename_column":
                detail = f"{a.get('old','')} -> {a.get('new','')}"
            elif action == "add_column":
                detail = f"{a.get('name','')} default={a.get('default','')}"
            else:
                detail = ""
            lines.append(f"- {action} {detail}")
        msg = "\\n".join(lines)
        return messagebox.askyesno("Confirm AI Actions", msg)
    def _parse_json_actions(self, text: str):
        s = (text or "").strip()
        if not s:
            return []
        if not s.startswith("["):
            l = s.find("[")
            r = s.rfind("]")
            if l >= 0 and r > l:
                s = s[l:r+1]
        try:
            parsed = json.loads(s)
            if isinstance(parsed, dict):
                return [parsed]
            return parsed
        except Exception:
            return None

    # Quick prompt understanding (Excel-like commands without AI)
    def _match_col_by_prompt(self, prompt: str) -> str | None:
        if not self.cleaner:
            return None
        pl = (prompt or "").lower()
        cols = [str(c) for c in self.cleaner.df.columns]
        # quoted column names first
        for q in re.findall(r"[\"']([^\"']+)[\"']", prompt or ""):
            for c in cols:
                if q.strip().lower() == c.lower():
                    return c
        for c in cols:
            v1 = c.lower()
            v2 = v1.replace("_", " ")
            v3 = v1.replace(" ", "_")
            if any(v in pl for v in (v1, v2, v3)):
                return c
        return None

    def _quick_actions_from_prompt(self, prompt: str) -> list:
        p = (prompt or "").strip().lower()
        if not p:
            return []
        actions = []
        col = self._match_col_by_prompt(prompt)
        no_time = any(k in p for k in ["no time", "no timestamp", "remove time", "date only"])

        if any(k in p for k in ["fix all headers", "fix headers", "fix header", "fix column names", "fix col names"]):
            actions.append({"action": "standardise_col_names"})

        if "trim whitespace" in p or "remove extra spaces" in p:
            actions.append({"action": "trim_whitespace"})
        if "clean signs" in p or "remove symbols" in p or "remove signs" in p:
            if col:
                actions.append({"action": "clean_signs_column", "column": col})
            else:
                actions.append({"action": "clean_signs_all"})

        if "drop duplicates" in p or "remove duplicates" in p:
            actions.append({"action": "drop_duplicates"})

        if "drop blank rows" in p or "remove blank rows" in p:
            actions.append({"action": "drop_blank_rows"})

        if "remove junk rows" in p:
            actions.append({"action": "remove_junk_rows"})

        if "fix numbers" in p or "fix numeric" in p:
            actions.append({"action": "fix_numeric_columns"})
        if "auto detect" in p and "type" in p and col:
            actions.append({"action": "auto_detect_type", "column": col})

        if "fix dates" in p or "fix date column" in p or "fix date columns" in p:
            if col:
                actions.append({
                    "action": "convert_type",
                    "column": col,
                    "type": "date",
                    "date_only": bool(no_time),
                })
            else:
                actions.append({
                    "action": "fix_date_columns_date_only" if no_time else "fix_date_columns"
                })

        # Range rules: "score must be less than 10" / "values above 10 make blank or 0"
        m_max = re.search(r"(?:less than|under|<=|below)\s*(\d+(?:\.\d+)?)", p)
        m_min = re.search(r"(?:greater than|over|>=|above)\s*(\d+(?:\.\d+)?)", p)
        if col and (m_max or m_min):
            fill = ""
            if "0" in p or "zero" in p:
                fill = 0
            if "blank" in p or "empty" in p:
                fill = ""
            actions.append({
                "action": "set_outside_range",
                "column": col,
                "min": float(m_min.group(1)) if m_min else None,
                "max": float(m_max.group(1)) if m_max else None,
                "fill": fill,
            })

        m = re.search(r"convert\\s+(.+?)\\s+to\\s+(number|numeric|date|text|string)", p)
        if m:
            col_guess = self._match_col_by_prompt(m.group(1)) or col
            if col_guess:
                actions.append({
                    "action": "convert_type",
                    "column": col_guess,
                    "type": m.group(2),
                    "date_only": bool(no_time and m.group(2) == "date"),
                })

        if any(k in p for k in ["clean all", "fix all", "one click clean"]):
            actions += [{"action": a} for a in [
                "fix_nulls", "trim_whitespace", "remove_junk_rows",
                "drop_blank_rows", "drop_duplicates",
                "fix_numeric_columns", "fix_date_columns",
                "standardise_col_names",
            ]]

        # remove duplicates while preserving order
        seen = set()
        uniq = []
        for a in actions:
            key = json.dumps(a, sort_keys=True)
            if key not in seen:
                seen.add(key)
                uniq.append(a)
        return uniq

    def _apply_ai_actions(self, actions):
        if not self.cleaner:
            return
        if not isinstance(actions, list):
            messagebox.showerror("AI Prompt Error", "AI did not return a valid JSON action list.")
            return
        self._log_clean("🤖  AI Prompt Actions")
        self._log_clean("─" * 52)
        for act in actions:
            if not isinstance(act, dict):
                continue
            a = str(act.get("action", "")).lower()
            msg = ""
            if a == "replace":
                msg = self.cleaner.replace_values(
                    act.get("column",""),
                    act.get("find",""),
                    act.get("replace",""),
                )
            elif a == "fill_missing":
                msg = self.cleaner.fill_missing(
                    act.get("column",""),
                    act.get("method","value"),
                    act.get("value",""),
                )
            elif a == "convert_type":
                msg = self.cleaner.convert_type(
                    act.get("column",""),
                    act.get("type",""),
                    bool(act.get("date_only", False)),
                )
            elif a == "drop_columns":
                msg = self.cleaner.drop_columns(act.get("columns", []))
            elif a == "set_outside_range":
                msg = self.cleaner.set_outside_range(
                    act.get("column",""),
                    act.get("min", None),
                    act.get("max", None),
                    act.get("fill", ""),
                )
            elif a == "clean_signs_all":
                msg = self.cleaner.clean_signs_all()
            elif a == "clean_signs_column":
                msg = self.cleaner.clean_signs_column(act.get("column",""))
            elif a == "auto_detect_type":
                msg = self.cleaner.auto_detect_column_type(act.get("column",""))
            elif a == "rename_column":
                msg = self.cleaner.rename_column(
                    act.get("old",""),
                    act.get("new",""),
                )
            elif a == "add_column":
                msg = self.cleaner.add_column_with_value(
                    act.get("name",""),
                    act.get("default",""),
                )
            elif a in ("drop_blank_rows","drop_duplicates","trim_whitespace",
                       "fix_numeric_columns","fix_date_columns","standardise_case",
                       "standardise_col_names","remove_junk_rows","fix_nulls",
                       "fix_date_columns_date_only"):
                msg = getattr(self.cleaner, a)()
            else:
                msg = f"âŒ  Unsupported action: {a}"
            if msg:
                self._log_clean(msg)
        self._log_clean("")
        if hasattr(self, "ai_clean_prompt"):
            try:
                self.ai_clean_prompt.delete("1.0", "end")
            except Exception:
                pass
        self._log_clean("✅  Prompt done. Type a new command above.")
        self._refresh()

    def _ai_prompt_clean(self):
        if not self.cleaner:
            messagebox.showwarning("No Data", "Please open a file first.")
            return
        key = self.cfg.get("api_key", "").strip()
        if not key:
            messagebox.showwarning(
                "No API Key",
                "Please add your free Groq API key in the âš™ï¸ Settings tab.\n\n"
                "Get it free at:  https://console.groq.com",
            )
            self.nb.select(self._tabs["settings"])
            return
        prompt = self.ai_clean_prompt.get("1.0", "end").strip()
        if not prompt:
            messagebox.showwarning("No Prompt", "Please type an AI cleaning prompt.")
            return

        quick_actions = self._quick_actions_from_prompt(prompt)
        if quick_actions:
            self._busy(True, "Applying quick actions???")
            self._apply_ai_actions(quick_actions)
            self._busy(False, "???  Quick actions applied")
            return

        self._busy(True, "AI is building actions???")

        def _worker():
            try:
                from groq import Groq
            except ImportError:
                self.root.after(0, lambda: (
                    messagebox.showerror(
                        "Package Missing",
                        "Run this in your VS Code terminal:\n\n    pip install groq",
                    ),
                    self._busy(False, ""),
                ))
                return
            try:
                profile = build_profile(self.cleaner.df)
                sys = (
                    "You are a data cleaning assistant. Output ONLY a JSON array of actions.\n"
                    "Allowed actions and fields:\n"
                    "1) replace: {action:'replace', column:'Col', find:'x', replace:'y'}\n"
                    "2) fill_missing: {action:'fill_missing', column:'Col', method:'value|mean|median|mode|ffill|bfill', value:'x'}\n"
                    "3) convert_type: {action:'convert_type', column:'Col', type:'number|date|text', date_only:true|false}\n"
                    "4) set_outside_range: {action:'set_outside_range', column:'Col', min:0, max:10, fill:''|0}\n"
                    "5) clean_signs_all: {action:'clean_signs_all'}\n"
                    "6) clean_signs_column: {action:'clean_signs_column', column:'Col'}\n"
                    "7) auto_detect_type: {action:'auto_detect_type', column:'Col'}\n"
                    "8) drop_columns: {action:'drop_columns', columns:['Col1','Col2']}\n"
                    "9) rename_column: {action:'rename_column', old:'Old', new:'New'}\n"
                    "10) add_column: {action:'add_column', name:'NewCol', default:'value'}\n"
                    "11) bulk: drop_blank_rows, drop_duplicates, trim_whitespace, fix_numeric_columns, "
                    "fix_date_columns, fix_date_columns_date_only, standardise_case, standardise_col_names, remove_junk_rows. "
                    "Example: {action:'drop_duplicates'}\n"
                    "If user says remove time/no timestamp, use date_only:true or fix_date_columns_date_only.\nUse exact column names from the profile. No explanations."
                )
                user = f"DATA PROFILE:\\n{profile}\\n\\nUSER PROMPT:\\n{prompt}"
                client = Groq(api_key=key)
                resp = client.chat.completions.create(
                    model="llama-3.3-70b-versatile",
                    messages=[
                        {"role":"system","content":sys},
                        {"role":"user","content":user},
                    ],
                    max_tokens=1200,
                    temperature=0.2,
                )
                text = resp.choices[0].message.content or ""
                actions = self._parse_json_actions(text)
                self.root.after(0, lambda: (
                    self._busy(False, "?  AI actions ready"),
                    self._apply_ai_actions(actions) if self._confirm_ai_actions(actions) else self._log_clean("??  AI actions canceled."),
                ))
            except Exception as e:
                self.root.after(0, lambda err=e: (
                    messagebox.showerror("AI Prompt Error", str(err)),
                    self._busy(False, "âŒ  AI prompt failed"),
                ))

        threading.Thread(target=_worker, daemon=True).start()

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


    def _edit_cell_dialog(self, row_idx: int, col_name: str):
        if not self.cleaner:
            return
        try:
            cur_val = self.cleaner.df.at[row_idx, col_name]
        except Exception:
            cur_val = ""
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

    def _replace_in_column_prompt(self, col: str, current_val=None):
        if not self.cleaner:
            return
        find = simpledialog.askstring(
            "Replace in Column",
            f"Column: {col}\nFind:",
            initialvalue="" if current_val is None else str(current_val),
            parent=self.root,
        )
        if find is None:
            return
        repl = simpledialog.askstring(
            "Replace in Column",
            f"Column: {col}\nReplace with:",
            parent=self.root,
        )
        if repl is None:
            return
        self._log_clean(self.cleaner.replace_values(col, find, repl))
        self._refresh()

    def _fill_missing_in_column_prompt(self, col: str):
        if not self.cleaner:
            return
        method = simpledialog.askstring(
            "Fill Missing",
            "Method: value / mean / median / mode / ffill / bfill",
            initialvalue="value",
            parent=self.root,
        )
        if method is None:
            return
        method = method.strip().lower() or "value"
        value = ""
        if method == "value":
            value = simpledialog.askstring(
                "Fill Missing",
                f"Column: {col}\nValue:",
                parent=self.root,
            )
            if value is None:
                return
        self._log_clean(self.cleaner.fill_missing(col, method, value))
        self._refresh()

    def _convert_column(self, col: str, to_type: str, date_only: bool = False):
        if not self.cleaner:
            return
        msg = self.cleaner.convert_type(col, to_type, date_only)
        self._log_edit(msg)
        self._refresh()

    def _trim_column(self, col: str):
        if not self.cleaner:
            return
        msg = self.cleaner.trim_whitespace_column(col)
        self._log_edit(msg)
        self._refresh()

    def _rename_column_prompt(self, col: str):
        if not self.cleaner:
            return
        new = simpledialog.askstring(
            "Rename Column",
            f"New name for '{col}':",
            parent=self.root,
        )
        if not new:
            return
        self._log_edit(self.cleaner.rename_column(col, new.strip()))
        self._refresh()

    def _delete_column_prompt(self, col: str):
        if not self.cleaner:
            return
        if not messagebox.askyesno(
            "Confirm",
            f"Delete column '{col}' ?\nThis can be undone with Ctrl+Z.",
        ):
            return
        self._log_edit(self.cleaner.delete_column(col))
        self._refresh()

    def _open_row_entry_dialog(self, idx: int | None = None):
        if not self.cleaner:
            return
        top = tk.Toplevel(self.root)
        top.title("Data Entry - Add Row")
        top.configure(bg=T["bg"])
        top.geometry("420x520")

        hdr = tk.Label(
            top, text="Data Entry Form",
            bg=T["bg"], fg=T["gold_bright"],
            font=("Segoe UI", 12, "bold"),
        )
        hdr.pack(anchor="w", padx=12, pady=(10, 6))

        idx_row = tk.Frame(top, bg=T["bg"])
        idx_row.pack(fill="x", padx=12, pady=4)
        tk.Label(idx_row, text="Insert after row (optional):",
                 bg=T["bg"], fg=T["dim"], font=("Segoe UI", 9)).pack(side="left")
        idx_var = tk.StringVar(value="" if idx is None else str(idx))
        idx_ent = ttk.Entry(idx_row, width=10, textvariable=idx_var)
        idx_ent.pack(side="left", padx=6)

        canvas = tk.Canvas(top, bg=T["bg"], highlightthickness=0)
        canvas.pack(side="left", fill="both", expand=True, padx=(12, 0), pady=8)
        sc = ttk.Scrollbar(top, orient="vertical", command=canvas.yview)
        sc.pack(side="right", fill="y", padx=(0, 10), pady=8)
        canvas.configure(yscrollcommand=sc.set)

        inner = tk.Frame(canvas, bg=T["bg"])
        canvas.create_window((0, 0), window=inner, anchor="nw")

        entries = {}
        for c in self.cleaner.df.columns:
            row = tk.Frame(inner, bg=T["bg"])
            row.pack(fill="x", pady=3)
            tk.Label(row, text=str(c), bg=T["bg"], fg=T["text"],
                     font=("Segoe UI", 9)).pack(side="left")
            ent = ttk.Entry(row, width=26)
            ent.pack(side="right", padx=6)
            entries[str(c)] = ent

        def _on_config(_):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _on_config)

        btn_row = tk.Frame(top, bg=T["bg"])
        btn_row.pack(fill="x", padx=12, pady=10)

        def _add():
            values = {}
            for k, ent in entries.items():
                v = ent.get().strip()
                if v != "":
                    values[k] = v
            try:
                insert_idx = int(idx_var.get()) if idx_var.get().strip() != "" else None
            except ValueError:
                messagebox.showwarning("Input", "Row index must be a number.")
                return
            msg = self.cleaner.add_row_values(values, insert_idx)
            self._log_edit(msg)
            self._refresh()
            top.destroy()

        ttk.Button(btn_row, text="Add Row", style="Gold.TButton", command=_add).pack(side="right")
        ttk.Button(btn_row, text="Cancel", command=top.destroy).pack(side="right", padx=8)

    def _apply_cond_format(self):
        if not self.cleaner:
            return
        col = self.cb_cond_col.get()
        op  = self.cb_cond_op.get()
        val = self.ent_cond_val.get()
        if not col or not op:
            messagebox.showwarning("Input", "Select a column and rule.")
            return
        self.cond_rule = {"col": col, "op": op, "val": val}
        self._refresh()
        self.edit_log.insert("end", f"🎨  Highlight: {col} {op} {val}\n")

    def _clear_cond_format(self):
        self.cond_rule = None
        self._refresh()
        self.edit_log.insert("end", "🎨  Highlight cleared\n")

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
        view_row_idx = self.tree.index(row_id)
        row_idx  = self._map_view_row(view_row_idx)
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
        region = self.tree.identify("region", event.x, event.y)
        col_id = self.tree.identify_column(event.x)

        cols = self.cleaner.df.columns.tolist()
        col_idx = None
        col_name = None
        if col_id:
            try:
                col_idx = int(col_id.replace("#", "")) - 1
            except Exception:
                col_idx = None
        if col_idx is not None and 0 <= col_idx < len(cols):
            col_name = cols[col_idx]

        # Header right-click menu
        if region == "heading" and col_name is not None:
            menu = tk.Menu(
                self.root, tearoff=0,
                bg=T["panel"], fg=T["text"],
                activebackground=T["select"],
                activeforeground=T["gold"],
                font=("Segoe UI", 9),
            )
            menu.add_command(
                label="Sort A -> Z",
                command=lambda: self._sort_column(col_name, True),
            )
            menu.add_command(
                label="Sort Z -> A",
                command=lambda: self._sort_column(col_name, False),
            )
            menu.add_separator()
            menu.add_command(
                label="Filter...",
                command=lambda: self._filter_column_dialog(col_name),
            )
            menu.add_command(
                label="Clear Filter",
                command=lambda: (self.active_filters.pop(col_name, None), self._refresh()),
            )
            menu.add_command(
                label="Clear All Filters",
                command=lambda: self._clear_filters(),
            )
            menu.add_separator()
            menu.add_command(
                label="Rename Column...",
                command=lambda: self._rename_column_prompt(col_name),
            )
            menu.add_command(
                label="Delete Column",
                command=lambda: self._delete_column_prompt(col_name),
            )
            menu.add_separator()
            menu.add_command(
                label="Convert Column -> Number",
                command=lambda: self._convert_column(col_name, "number"),
            )
            menu.add_command(
                label="Convert Column -> Date (Keep Time)",
                command=lambda: self._convert_column(col_name, "datetime"),
            )
            menu.add_command(
                label="Convert Column -> Date Only",
                command=lambda: self._convert_column(col_name, "date", date_only=True),
            )
            menu.add_command(
                label="Convert Column -> Text",
                command=lambda: self._convert_column(col_name, "text"),
            )
            menu.add_separator()
            menu.add_command(
                label="Trim Whitespace In Column",
                command=lambda: self._trim_column(col_name),
            )
            menu.add_command(
                label="Clean Symbols In Column",
                command=lambda: self._log_edit(self.cleaner.clean_signs_column(col_name)) or self._refresh(),
            )
            menu.add_command(
                label="Replace In Column...",
                command=lambda: self._replace_in_column_prompt(col_name, None),
            )
            menu.add_command(
                label="Fill Missing In Column...",
                command=lambda: self._fill_missing_in_column_prompt(col_name),
            )
            menu.add_separator()
            menu.add_command(
                label="Fix All Headers",
                command=lambda: (self._log_edit(self.cleaner.standardise_col_names()), self._refresh()),
            )
            menu.post(event.x_root, event.y_root)
            return

        # Cell right-click menu
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        view_row_idx = self.tree.index(row_id)
        row_idx = self._map_view_row(view_row_idx)

        cur_val = None
        if col_name is not None:
            try:
                cur_val = self.cleaner.df.iat[row_idx, col_idx]
            except Exception:
                cur_val = None

        menu = tk.Menu(
            self.root, tearoff=0,
            bg=T["panel"], fg=T["text"],
            activebackground=T["select"],
            activeforeground=T["gold"],
            font=("Segoe UI", 9),
        )
        menu.add_command(
            label="Add Row Below",
            command=lambda: (
                self._log_edit(self.cleaner.add_row(idx=row_idx)),
                self._refresh(),
            ),
        )
        menu.add_command(
            label="Add Row With Values...",
            command=lambda: self._open_row_entry_dialog(idx=row_idx),
        )
        menu.add_command(
            label="Duplicate This Row",
            command=lambda: (
                self._log_edit(self.cleaner.duplicate_row(row_idx)),
                self._refresh(),
            ),
        )
        menu.add_command(
            label="Delete This Row",
            command=lambda: (
                self._log_edit(self.cleaner.delete_row(row_idx)),
                self._refresh(),
            ),
        )

        if col_name is not None:
            menu.add_separator()
            menu.add_command(
                label="Edit Cell...",
                command=lambda: self._edit_cell_dialog(row_idx, col_name),
            )
            menu.add_command(
                label="Replace In Column...",
                command=lambda: self._replace_in_column_prompt(col_name, cur_val),
            )
            menu.add_command(
                label="Fill Missing In Column...",
                command=lambda: self._fill_missing_in_column_prompt(col_name),
            )
            menu.add_command(
                label="Trim Whitespace In Column",
                command=lambda: self._trim_column(col_name),
            )
            menu.add_command(
                label="Clean Symbols In Column",
                command=lambda: self._log_edit(self.cleaner.clean_signs_column(col_name)) or self._refresh(),
            )
            menu.add_separator()
            menu.add_command(
                label="Convert Column -> Number",
                command=lambda: self._convert_column(col_name, "number"),
            )
            menu.add_command(
                label="Convert Column -> Date (Keep Time)",
                command=lambda: self._convert_column(col_name, "datetime"),
            )
            menu.add_command(
                label="Convert Column -> Date Only",
                command=lambda: self._convert_column(col_name, "date", date_only=True),
            )
            menu.add_command(
                label="Convert Column -> Text",
                command=lambda: self._convert_column(col_name, "text"),
            )
            menu.add_separator()
            menu.add_command(
                label="Rename Column...",
                command=lambda: self._rename_column_prompt(col_name),
            )
            menu.add_command(
                label="Auto Detect Column Type",
                command=lambda: (self._log_edit(self.cleaner.auto_detect_column_type(col_name)), self._refresh()),
            )
            menu.add_command(
                label="Delete Column",
                command=lambda: self._delete_column_prompt(col_name),
            )
            menu.add_separator()
            menu.add_command(label="Undo", command=self.undo)
            menu.add_command(label="Redo", command=self.redo)
            menu.add_separator()
            menu.add_command(label="Undo", command=self.undo)
            menu.add_command(label="Redo", command=self.redo)

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
        if self.current_sheet:
            self.var_status.set(f"↺  Reset sheet '{self.current_sheet}'")
        else:
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
        agg = self.cb_cht_agg.get()
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
            if ct in ("bar", "line", "pie"):
                if agg == "count":
                    series = df.groupby(xc, dropna=False).size()
                else:
                    if not yc:
                        raise ValueError("Select a Y column for this chart.")
                    series = df.groupby(xc, dropna=False)[yc].agg(agg)
                if ct == "bar":
                    series.plot(kind="bar", ax=ax, color=T["gold"], edgecolor=T["bg"])
                elif ct == "line":
                    series.plot(kind="line", ax=ax, color=T["gold"], marker="o")
                elif ct == "pie":
                    series.plot(kind="pie", ax=ax, autopct="%1.1f%%")
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

            title_y = "Count" if agg == "count" and ct in ("bar","line","pie") else yc
            ax.set_title(f"{ct.title()} Chart  —  {title_y} by {xc}",
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
            sheet = f"\nSheet: {self.current_sheet}" if self.current_sheet else ""
            messagebox.showinfo("Saved", f"CSV saved:\n{path}{sheet}")

    def export_excel(self):
        if not self.cleaner:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if path:
            if len(self.sheets) > 1:
                with pd.ExcelWriter(path, engine="openpyxl") as writer:
                    for name, cleaner in self.sheets.items():
                        cleaner.df.to_excel(writer, sheet_name=name[:31], index=False)
            else:
                self.cleaner.df.to_excel(path, index=False)
            messagebox.showinfo("Saved", f"Excel saved:\n{path}")

    def export_json(self):
        if not self.cleaner:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json")],
        )
        if not path:
            return
        try:
            data = self.cleaner.df.to_dict(orient="records")
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Saved", f"JSON saved:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


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
            f"SHEET: {self.current_sheet or 'Sheet1'}", "",
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
                self.cfg["api_key"] = key
                cfg_save(self.cfg)
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
