"""
╔══════════════════════════════════════════════════════════════╗
║            EXCEL AI CLEANER — by Sajeeb The Analyst         ║
║                      ai_engine.py                           ║
║     Groq AI brain — scan, stream, parse, suggest fixes      ║
╚══════════════════════════════════════════════════════════════╝

Features:
  • Builds a rich statistical profile of the data (no raw rows sent)
  • Streams Groq / Llama-3 response in real time to the UI
  • Parses AI output into structured FixCard objects
  • Each FixCard maps to a real cleaner.py function
  • Works on datasets of ANY size
  • 100% free — uses Groq free tier
"""

import re
import threading
from dataclasses import dataclass, field
from typing import Callable, Optional

import numpy as np
import pandas as pd


# ════════════════════════════════════════════════════════════
#  FIX CARD  — one parsed issue + its auto-fix action
# ════════════════════════════════════════════════════════════

@dataclass
class FixCard:
    """
    Represents one data quality issue found by the AI.
    The UI renders this as a clickable gold card.
    """
    issue_id    : int
    emoji       : str          # leading emoji from AI output
    title       : str          # short issue title
    detail      : str          # full description
    fix_label   : str          # what the Apply button will say
    fix_action  : str          # key that maps to a cleaner function
    fix_kwargs  : dict = field(default_factory=dict)
    applied     : bool = False


# ════════════════════════════════════════════════════════════
#  ACTION MAP  — maps fix_action keys → cleaner method names
# ════════════════════════════════════════════════════════════

# Each entry: fix_action_key → (cleaner_method_name, kwargs_override)
ACTION_MAP = {
    "fix_nulls"              : ("fix_nulls",               {}),
    "trim_whitespace"        : ("trim_whitespace",          {}),
    "drop_blank_rows"        : ("drop_blank_rows",          {}),
    "drop_duplicates"        : ("drop_duplicates",          {}),
    "remove_junk_rows"       : ("remove_junk_rows",         {}),
    "fix_numeric_columns"    : ("fix_numeric_columns",      {}),
    "fix_date_columns"       : ("fix_date_columns",         {}),
    "standardise_case"       : ("standardise_case",         {"mode": "title"}),
    "standardise_col_names"  : ("standardise_column_names", {}),
    "one_click_clean"        : ("one_click_clean",          {}),
}


# ════════════════════════════════════════════════════════════
#  DATA PROFILER  — builds summary sent to AI
# ════════════════════════════════════════════════════════════

def build_profile(df: pd.DataFrame) -> str:
    """
    Builds a rich text profile of the DataFrame.
    Only statistics are sent — never raw data rows.
    Works for any size dataset.
    """
    lines = [
        "═" * 60,
        "DATASET PROFILE",
        "═" * 60,
        f"Shape          : {len(df):,} rows × {len(df.columns)} columns",
        f"Total missing  : {int(df.isna().sum().sum()):,} cells",
        f"Duplicate rows : {int(df.duplicated().sum()):,}",
        f"Junk rows      : {int((df.isna().mean(axis=1) > 0.8).sum()):,}  (>80% empty)",
        "",
    ]

    for col in df.columns:
        s       = df[col]
        missing = int(s.isna().sum())
        pct     = f"{missing / len(df) * 100:.1f}%" if len(df) else "0%"
        dtype   = str(s.dtype)
        uniq    = s.nunique()

        lines.append(f"── COLUMN: `{col}`")
        lines.append(f"   dtype={dtype}  |  missing={missing} ({pct})  |  unique={uniq}")

        non_null = s.dropna().astype(str).str.strip()
        if non_null.empty:
            lines.append("   ⚠ All values are missing")
            lines.append("")
            continue

        # Sample values
        top_vals = non_null.value_counts().head(10).index.tolist()
        lines.append(f"   Top values    : {top_vals}")

        # Numeric stats
        if pd.api.types.is_numeric_dtype(s):
            num = pd.to_numeric(s, errors="coerce").dropna()
            if len(num):
                lines.append(
                    f"   Numeric stats : min={num.min():.4g}  max={num.max():.4g}"
                    f"  mean={num.mean():.4g}  std={num.std():.4g}"
                )
                inf_c = int(np.isinf(num).sum())
                if inf_c:
                    lines.append(f"   ⚠ Contains {inf_c} inf / -inf values")

                # IQR outliers
                if len(num) >= 4:
                    q1, q3 = num.quantile(0.25), num.quantile(0.75)
                    iqr    = q3 - q1
                    if iqr > 0:
                        out = int(((num < q1 - 1.5*iqr) | (num > q3 + 1.5*iqr)).sum())
                        if out:
                            lines.append(f"   ⚠ Outliers (IQR): {out} values")
        else:
            # Date parse rate
            try:
                dr = pd.to_datetime(
                    non_null, errors="coerce"
                ).notna().mean()
                if dr > 0.5:
                    lines.append(f"   Date parse rate: {dr*100:.0f}%  (stored as text)")
            except Exception:
                pass

            # Numeric-as-text rate
            nr = pd.to_numeric(
                non_null.str.replace(r"[$,\s%€£¥]", "", regex=True),
                errors="coerce",
            ).notna().mean()
            if nr > 0.7:
                lines.append(f"   Numeric-as-text: {nr*100:.0f}%")

            # Mixed date formats
            fmts = set()
            for v in non_null.head(80):
                if re.search(r"\d{4}-\d{2}-\d{2}", v):   fmts.add("YYYY-MM-DD")
                if re.search(r"\d{2}/\d{2}/\d{4}", v):   fmts.add("DD/MM/YYYY")
                if re.search(r"\d{2}-\d{2}-\d{4}", v):   fmts.add("DD-MM-YYYY")
                if re.search(r"\d{2} \w{3} \d{2,4}", v): fmts.add("DD Mon YYYY")
                if re.search(r"\d{2}/\d{2}-\d{2}", v):   fmts.add("mixed/malformed")
            if len(fmts) > 1:
                lines.append(f"   Mixed date fmts: {', '.join(sorted(fmts))}")

            # Case inconsistency
            vals = non_null.unique()
            lower_map: dict = {}
            for v in vals:
                lower_map.setdefault(v.lower().strip(), []).append(v)
            conflicts = {k: v for k, v in lower_map.items() if len(v) > 1}
            if conflicts:
                ex = "; ".join(
                    "/".join(repr(x) for x in v[:3])
                    for v in list(conflicts.values())[:3]
                )
                lines.append(f"   Case conflicts : {ex}")

            # Currency symbols
            money = non_null.str.contains(r"[$€£¥]", regex=True).sum()
            if money:
                lines.append(f"   Currency symbols in {money} values")

            # Multi-value cells
            if non_null.str.contains(r"[,/|]", regex=True).mean() > 0.5:
                lines.append("   Multi-value cells (comma/slash separated)")

        lines.append("")

    return "\n".join(lines)


# ════════════════════════════════════════════════════════════
#  EXPERT SYSTEM PROMPT
# ════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """You are a world-class data analyst and data quality engineer with 15 years of experience.
A user has uploaded a dataset. You are given a detailed statistical profile of it.

Your job: find EVERY data quality problem and suggest a specific fix for each one.

ISSUES TO LOOK FOR:
1. Missing / null values — which columns, severity, percentage
2. Duplicate rows
3. Wrong data types — numbers or dates stored as text
4. Inconsistent formats — mixed date formats, currency symbols, mixed separators
5. Inconsistent capitalisation — e.g. "usa" vs "USA" vs "Us"
6. Typos and misspellings — e.g. "New Zeland", "Pg-13" vs "PG-13"
7. Wrong decimal scale — e.g. score=86.0 when max should be 10.0
8. Impossible or suspicious values — inf, negative duration, 0 revenue
9. Ghost columns — Col_0, Col_1, Unnamed, completely empty columns
10. Truncated column names — very short names ending in consonant
11. Multi-value cells — genres, tags separated by commas/slashes
12. ID columns with duplicates
13. Nearly empty columns — more than 80% missing
14. Outliers — statistically detected via IQR

FORMAT YOUR RESPONSE EXACTLY LIKE THIS — one issue per block:

ISSUE: [short title of the problem]
DETAIL: [specific description — mention exact column names and values]
FIX: [exact action to take]
ACTION: [one of: fix_nulls | trim_whitespace | drop_blank_rows | drop_duplicates | remove_junk_rows | fix_numeric_columns | fix_date_columns | standardise_case | standardise_col_names | manual_review]
---

After all issues add:

SUMMARY:
[2-3 sentence overall assessment of the dataset quality]

Be specific. Reference exact column names. Be direct and concise.
Do NOT repeat the profile back. Do NOT add any preamble."""


# ════════════════════════════════════════════════════════════
#  FIX CARD PARSER
# ════════════════════════════════════════════════════════════

# Map common issue keywords to emojis
EMOJI_MAP = {
    "missing"     : "❓",
    "null"        : "❓",
    "blank"       : "❓",
    "duplicate"   : "⚠️",
    "numeric"     : "🔢",
    "number"      : "🔢",
    "date"        : "📅",
    "time"        : "📅",
    "case"        : "🔠",
    "capital"     : "🔠",
    "typo"        : "🔤",
    "spelling"    : "🔤",
    "outlier"     : "📊",
    "column"      : "👻",
    "ghost"       : "👻",
    "whitespace"  : "✂️",
    "space"       : "✂️",
    "currency"    : "💰",
    "symbol"      : "💰",
    "junk"        : "🗑️",
    "empty"       : "🗑️",
    "scale"       : "🔢",
    "decimal"     : "🔢",
    "format"      : "📅",
    "truncat"     : "✂️",
    "id"          : "🔑",
}

# Human-readable fix button labels per action
FIX_LABELS = {
    "fix_nulls"           : "Fix Null Values",
    "trim_whitespace"     : "Trim Whitespace",
    "drop_blank_rows"     : "Drop Blank Rows",
    "drop_duplicates"     : "Remove Duplicates",
    "remove_junk_rows"    : "Remove Junk Rows",
    "fix_numeric_columns" : "Convert to Numeric",
    "fix_date_columns"    : "Convert to Datetime",
    "standardise_case"    : "Standardise Case",
    "standardise_col_names": "Fix Column Names",
    "manual_review"       : "Needs Manual Review",
}


def _pick_emoji(text: str) -> str:
    t = text.lower()
    for keyword, emoji in EMOJI_MAP.items():
        if keyword in t:
            return emoji
    return "⚠️"


def parse_fix_cards(ai_response: str) -> list[FixCard]:
    """
    Parse AI response into a list of FixCard objects.
    Looks for ISSUE / DETAIL / FIX / ACTION blocks separated by ---.
    """
    cards   : list[FixCard] = []
    summary : str           = ""

    # Extract summary
    sum_match = re.search(r"SUMMARY:\s*(.+?)(?:\Z|ISSUE:)", ai_response,
                          re.DOTALL | re.IGNORECASE)
    if sum_match:
        summary = sum_match.group(1).strip()

    # Split on --- separator
    blocks = re.split(r"\n---+\n?", ai_response)

    issue_id = 0
    for block in blocks:
        block = block.strip()
        if not block or block.upper().startswith("SUMMARY"):
            continue

        # Extract fields
        issue_match  = re.search(r"ISSUE:\s*(.+)",  block, re.IGNORECASE)
        detail_match = re.search(r"DETAIL:\s*(.+)", block, re.IGNORECASE | re.DOTALL)
        fix_match    = re.search(r"FIX:\s*(.+)",    block, re.IGNORECASE)
        action_match = re.search(r"ACTION:\s*(.+)", block, re.IGNORECASE)

        if not issue_match:
            continue

        title  = issue_match.group(1).strip()
        detail = ""
        if detail_match:
            # Detail goes until FIX: line
            raw_detail = detail_match.group(1)
            detail = re.split(r"\nFIX:", raw_detail, flags=re.IGNORECASE)[0].strip()

        fix_desc   = fix_match.group(1).strip()  if fix_match    else ""
        action_key = action_match.group(1).strip().lower() if action_match else "manual_review"

        # Normalise action key
        action_key = action_key.replace(" ", "_")
        if action_key not in ACTION_MAP and action_key != "manual_review":
            action_key = "manual_review"

        emoji = _pick_emoji(title + " " + detail)
        label = FIX_LABELS.get(action_key, "Apply Fix")

        cards.append(FixCard(
            issue_id   = issue_id,
            emoji      = emoji,
            title      = title,
            detail     = detail or fix_desc,
            fix_label  = label,
            fix_action = action_key,
        ))
        issue_id += 1

    return cards, summary


# ════════════════════════════════════════════════════════════
#  GROQ AI ENGINE
# ════════════════════════════════════════════════════════════

class AIEngine:
    """
    Handles all communication with the Groq API.
    Streams response and calls callbacks for live UI updates.
    """

    MODEL = "llama-3.3-70b-versatile"

    def __init__(self, api_key: str = ""):
        self.api_key = api_key
        self._running = False

    def set_key(self, key: str):
        self.api_key = key.strip()

    def is_ready(self) -> bool:
        return bool(self.api_key)

    # ── test connection ──────────────────────────────────────
    def test_connection(self,
                        on_success: Optional[Callable] = None,
                        on_error:   Optional[Callable] = None):
        def _test():
            try:
                from groq import Groq
                client = Groq(api_key=self.api_key)
                client.chat.completions.create(
                    model    = self.MODEL,
                    messages = [{"role": "user", "content": "Reply OK"}],
                    max_tokens = 5,
                )
                if on_success:
                    on_success("✅  Groq API connected successfully!")
            except ImportError:
                if on_error:
                    on_error("❌  Package missing.\n\nRun:\n    pip install groq")
            except Exception as e:
                if on_error:
                    on_error(f"❌  Connection failed:\n\n{str(e)}")

        threading.Thread(target=_test, daemon=True).start()

    # ── full scan ────────────────────────────────────────────
    def scan(
        self,
        df          : pd.DataFrame,
        on_chunk    : Optional[Callable[[str], None]]              = None,
        on_cards    : Optional[Callable[[list, str], None]]        = None,
        on_done     : Optional[Callable[[str], None]]              = None,
        on_error    : Optional[Callable[[str], None]]              = None,
        on_status   : Optional[Callable[[str], None]]              = None,
    ):
        """
        Run AI scan in a background thread.

        Callbacks:
          on_chunk(text)         — called for each streamed token
          on_cards(cards, summary) — called with parsed FixCard list when done
          on_done(full_text)     — called when streaming finishes
          on_error(message)      — called on any error
          on_status(message)     — called for progress status updates
        """
        if not self.api_key:
            if on_error:
                on_error("❌  No API key set.\nGo to ⚙️ Settings and add your Groq key.")
            return

        self._running = True
        threading.Thread(
            target = self._scan_worker,
            args   = (df, on_chunk, on_cards, on_done, on_error, on_status),
            daemon = True,
        ).start()

    def stop(self):
        self._running = False

    # ── background worker ────────────────────────────────────
    def _scan_worker(self, df, on_chunk, on_cards, on_done, on_error, on_status):
        try:
            from groq import Groq
        except ImportError:
            if on_error:
                on_error(
                    "❌  'groq' package not installed.\n\n"
                    "Run this in your VS Code terminal:\n\n"
                    "    pip install groq"
                )
            return

        # 1 — build profile
        if on_status:
            on_status("🧠  Building data profile…")

        profile = build_profile(df)

        # 2 — stream from Groq
        if on_status:
            on_status("🤖  AI is thinking…")

        try:
            client = Groq(api_key=self.api_key)
            stream = client.chat.completions.create(
                model    = self.MODEL,
                messages = [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",   "content": f"Dataset profile:\n\n{profile}"},
                ],
                max_tokens = 2500,
                temperature= 0.2,    # low temp = more consistent, factual output
                stream     = True,
            )

            chunks = []
            for chunk in stream:
                if not self._running:
                    break
                piece = chunk.choices[0].delta.content or ""
                chunks.append(piece)
                if on_chunk:
                    on_chunk(piece)

            full_response = "".join(chunks)

        except Exception as e:
            if on_error:
                on_error(
                    f"❌  Groq API error:\n\n{str(e)}\n\n"
                    "Check your API key in ⚙️ Settings."
                )
            self._running = False
            return

        # 3 — parse fix cards
        if on_status:
            on_status("🃏  Parsing fix cards…")

        try:
            cards, summary = parse_fix_cards(full_response)
        except Exception:
            cards, summary = [], ""

        # 4 — callbacks
        if on_cards:
            on_cards(cards, summary)
        if on_done:
            on_done(full_response)

        self._running = False

    # ── apply a fix card ─────────────────────────────────────
    def apply_fix(self, card: FixCard, cleaner) -> str:
        """
        Apply the fix associated with a FixCard to the cleaner.
        Returns log message.
        """
        if card.applied:
            return f"ℹ️  Fix '{card.title}' was already applied."

        action_key = card.fix_action
        if action_key == "manual_review":
            return f"ℹ️  '{card.title}' requires manual review — no auto-fix available."

        if action_key not in ACTION_MAP:
            return f"❌  Unknown action: {action_key}"

        method_name, default_kwargs = ACTION_MAP[action_key]
        kwargs = {**default_kwargs, **card.fix_kwargs}

        method = getattr(cleaner, method_name, None)
        if method is None:
            return f"❌  Cleaner method '{method_name}' not found."

        try:
            result = method(**kwargs) if kwargs else method()
            card.applied = True
            return result if isinstance(result, str) else f"✅  Applied: {card.title}"
        except Exception as e:
            return f"❌  Error applying fix: {e}"


# ════════════════════════════════════════════════════════════
#  SINGLETON FACTORY
# ════════════════════════════════════════════════════════════

_engine_instance: Optional[AIEngine] = None

def get_engine(api_key: str = "") -> AIEngine:
    """Return the shared AIEngine instance."""
    global _engine_instance
    if _engine_instance is None:
        _engine_instance = AIEngine(api_key)
    elif api_key:
        _engine_instance.set_key(api_key)
    return _engine_instance