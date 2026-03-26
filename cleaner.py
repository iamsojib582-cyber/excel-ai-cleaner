"""
╔══════════════════════════════════════════════════════════════╗
║            EXCEL AI CLEANER — by Sajeeb The Analyst         ║
║                       cleaner.py                            ║
║      Cleaning engine + full Undo / Redo history stack       ║
╚══════════════════════════════════════════════════════════════╝

Features:
  • Every action is reversible via undo/redo
  • Tracks before/after stats for PDF report
  • One-click clean runs all fixes in smart order
  • Manual cell / row / column editing
  • Modified cells tracked for gold highlight in UI
  • Type-safe cell editing with validation
"""

import re
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Optional

import numpy as np
import pandas as pd


# ════════════════════════════════════════════════════════════
#  ACTION RECORD  — stored in undo/redo stack
# ════════════════════════════════════════════════════════════

@dataclass
class Action:
    """Represents one reversible cleaning or editing action."""
    name        : str                        # human-readable label
    df_before   : pd.DataFrame               # snapshot before
    df_after    : pd.DataFrame               # snapshot after
    timestamp   : str = field(
        default_factory=lambda: datetime.now().strftime("%H:%M:%S"))
    description : str = ""                   # detail for clean log


# ════════════════════════════════════════════════════════════
#  UNDO / REDO ENGINE
# ════════════════════════════════════════════════════════════

class UndoRedoStack:
    MAX_HISTORY = 50          # cap memory usage

    def __init__(self):
        self._undo: list[Action] = []
        self._redo: list[Action] = []

    # ── push a new action ────────────────────────────────────
    def push(self, action: Action):
        self._undo.append(action)
        if len(self._undo) > self.MAX_HISTORY:
            self._undo.pop(0)
        self._redo.clear()    # new action clears redo branch

    # ── undo ─────────────────────────────────────────────────
    def undo(self) -> Optional[Action]:
        if not self._undo:
            return None
        action = self._undo.pop()
        self._redo.append(action)
        return action

    # ── redo ─────────────────────────────────────────────────
    def redo(self) -> Optional[Action]:
        if not self._redo:
            return None
        action = self._redo.pop()
        self._undo.append(action)
        return action

    # ── state ────────────────────────────────────────────────
    def can_undo(self) -> bool: return bool(self._undo)
    def can_redo(self) -> bool: return bool(self._redo)

    def undo_label(self) -> str:
        return self._undo[-1].name if self._undo else ""

    def redo_label(self) -> str:
        return self._redo[-1].name if self._redo else ""

    def history(self) -> list[str]:
        """Return list of action names oldest→newest."""
        return [a.name for a in self._undo]

    def clear(self):
        self._undo.clear()
        self._redo.clear()


# ════════════════════════════════════════════════════════════
#  STATS SNAPSHOT  — for before/after comparison & PDF report
# ════════════════════════════════════════════════════════════

@dataclass
class DataStats:
    rows        : int
    columns     : int
    missing     : int
    duplicates  : int
    numeric_cols: int
    text_cols   : int
    date_cols   : int
    timestamp   : str = field(
        default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def summary(self) -> str:
        return (
            f"Rows: {self.rows:,}  |  Columns: {self.columns}  |  "
            f"Missing: {self.missing:,}  |  Duplicates: {self.duplicates:,}"
        )


def snapshot_stats(df: pd.DataFrame) -> DataStats:
    return DataStats(
        rows         = len(df),
        columns      = len(df.columns),
        missing      = int(df.isna().sum().sum()),
        duplicates   = int(df.duplicated().sum()),
        numeric_cols = len(df.select_dtypes(include="number").columns),
        text_cols    = len(df.select_dtypes(include=["object", "string"]).columns),
        date_cols    = len(df.select_dtypes(include=["datetime"]).columns),
    )


# ════════════════════════════════════════════════════════════
#  CLEANER  — main class
# ════════════════════════════════════════════════════════════

class Cleaner:
    """
    Wraps a DataFrame and exposes every cleaning/editing operation.
    All mutating methods:
      1. Copy df_before
      2. Apply the change
      3. Push an Action to the undo stack
      4. Return a description string for the clean log
    """

    NULL_TOKENS = [
        "NULL", "null", "N/A", "NA", "n/a", "na",
        "None", "none", "NaN", "nan", "#N/A", "#n/a",
        "#NA", "?", "-", "", "INF", "inf", "-inf",
        "N\\A", "N/a", "NIL", "nil", "NONE",
    ]

    def __init__(self, df: pd.DataFrame):
        self.df              = df.copy()
        self.stack           = UndoRedoStack()
        self.stats_before    : Optional[DataStats] = None
        self.stats_after     : Optional[DataStats] = None
        self.modified_cells  : set[tuple] = set()   # (row_index, col_name)
        self._record_before()

    # ── internal helpers ─────────────────────────────────────

    def _record_before(self):
        """Snapshot stats when a file is first loaded."""
        self.stats_before = snapshot_stats(self.df)

    def _commit(self, name: str, df_before: pd.DataFrame,
                description: str = "") -> str:
        """Push action + update after-stats + return description."""
        action = Action(
            name        = name,
            df_before   = df_before,
            df_after    = self.df.copy(),
            description = description,
        )
        self.stack.push(action)
        self.stats_after = snapshot_stats(self.df)
        return description or name

    # ── undo / redo ──────────────────────────────────────────

    def undo(self) -> Optional[str]:
        action = self.stack.undo()
        if action is None:
            return None
        self.df = action.df_before.copy()
        self.stats_after = snapshot_stats(self.df)
        return f"↩  Undone: {action.name}"

    def redo(self) -> Optional[str]:
        action = self.stack.redo()
        if action is None:
            return None
        self.df = action.df_after.copy()
        self.stats_after = snapshot_stats(self.df)
        return f"↪  Redone: {action.name}"

    def can_undo(self) -> bool: return self.stack.can_undo()
    def can_redo(self) -> bool: return self.stack.can_redo()

    # ════════════════════════════════════════════════════════
    #  CLEANING OPERATIONS
    # ════════════════════════════════════════════════════════

    # ── 1. Replace null-like tokens ──────────────────────────
    def fix_nulls(self) -> str:
        before  = self.df.copy()
        count   = int(self.df.isin(self.NULL_TOKENS).sum().sum())
        self.df = self.df.replace(self.NULL_TOKENS, pd.NA)
        desc    = f"🚫  Replaced {count} null-like tokens (NULL, #N/A, inf…) with blanks."
        return self._commit("Fix Nulls", before, desc)

    # ── 2. Trim whitespace ───────────────────────────────────
    def trim_whitespace(self) -> str:
        before = self.df.copy()
        for col in self.df.select_dtypes(include=["object", "string"]).columns:
            self.df[col] = self.df[col].astype("string").str.strip()
        desc = "✂️   Trimmed leading/trailing whitespace from all text columns."
        return self._commit("Trim Whitespace", before, desc)

    # ── 3. Drop blank rows ───────────────────────────────────
    def drop_blank_rows(self) -> str:
        before = self.df.copy()
        count  = int(self.df.isna().all(axis=1).sum())
        self.df = self.df.dropna(how="all").reset_index(drop=True)
        desc = f"🗑️   Dropped {count} fully blank rows."
        return self._commit("Drop Blank Rows", before, desc)

    # ── 4. Drop duplicate rows ───────────────────────────────
    def drop_duplicates(self) -> str:
        before = self.df.copy()
        count  = int(self.df.duplicated().sum())
        self.df = self.df.drop_duplicates().reset_index(drop=True)
        desc = f"🗑️   Dropped {count} duplicate rows."
        return self._commit("Drop Duplicates", before, desc)

    # ── 5. Remove junk rows (>80% empty) ─────────────────────
    def remove_junk_rows(self) -> str:
        before  = self.df.copy()
        mask    = self.df.isna().mean(axis=1) > 0.80
        count   = int(mask.sum())
        self.df = self.df[~mask].reset_index(drop=True)
        desc = f"🗑️   Removed {count} junk rows (more than 80% empty)."
        return self._commit("Remove Junk Rows", before, desc)

    # ── 6. Fix numeric columns stored as text ────────────────
    def fix_numeric_columns(self) -> str:
        before = self.df.copy()
        fixed  = []
        for col in self.df.select_dtypes(include=["object", "string"]).columns:
            s   = self.df[col].astype(str).str.replace(r"[$,\s%€£¥]", "", regex=True)
            num = pd.to_numeric(s, errors="coerce")
            if num.notna().mean() > 0.80:
                self.df[col] = pd.to_numeric(s, errors="coerce")
                fixed.append(col)
        desc = (
            f"🔢  Converted {len(fixed)} text column(s) to numeric: "
            f"{', '.join(f'`{c}`' for c in fixed)}."
            if fixed else "🔢  No text-as-numeric columns found."
        )
        return self._commit("Fix Numeric Columns", before, desc)

    # ── 7. Fix date columns stored as text ───────────────────
    def fix_date_columns(self) -> str:
        before = self.df.copy()
        fixed  = []
        for col in self.df.select_dtypes(include=["object", "string"]).columns:
            try:
                parsed = pd.to_datetime(
                    self.df[col].astype(str),
                    errors="coerce",
                )
                if parsed.notna().mean() > 0.70:
                    self.df[col] = parsed
                    fixed.append(col)
            except Exception:
                pass
        desc = (
            f"📅  Converted {len(fixed)} text column(s) to datetime: "
            f"{', '.join(f'`{c}`' for c in fixed)}."
            if fixed else "📅  No text-as-date columns found."
        )
        return self._commit("Fix Date Columns", before, desc)

    # ── 8. Standardise text case ─────────────────────────────
    def standardise_case(self, mode: str = "title") -> str:
        """mode: 'title' | 'upper' | 'lower'"""
        before = self.df.copy()
        for col in self.df.select_dtypes(include=["object", "string"]).columns:
            s = self.df[col].astype("string")
            if mode == "title":
                self.df[col] = s.str.title()
            elif mode == "upper":
                self.df[col] = s.str.upper()
            elif mode == "lower":
                self.df[col] = s.str.lower()
        label = {"title": "Title Case", "upper": "UPPER CASE", "lower": "lower case"}[mode]
        desc  = f"🔠  Standardised all text columns to {label}."
        return self._commit(f"Standardise Case ({label})", before, desc)

    # ── 9. Standardise column names ──────────────────────────
    def standardise_column_names(self) -> str:
        before = self.df.copy()
        self.df.columns = [
            re.sub(r"\s+", "_",
                   re.sub(r"[^\w\s]", "", str(c)).strip()).upper()
            for c in self.df.columns
        ]
        desc = "🏷️   Standardised column names to UPPER_SNAKE_CASE."
        return self._commit("Standardise Column Names", before, desc)

    # ── 10. Find & replace in a column ───────────────────────
    def replace_values(self, column: str, find: Any, replace_with: Any) -> str:
        if column not in self.df.columns:
            return f"❌  Column `{column}` not found."
        before     = self.df.copy()
        count      = int((self.df[column].astype(str) == str(find)).sum())
        self.df[column] = self.df[column].replace(find, replace_with)
        desc = f"🔄  Replaced '{find}' → '{replace_with}' in `{column}` ({count} cells)."
        return self._commit("Replace Values", before, desc)

    # ── 11. Fill missing values in a column ──────────────────
    def fill_missing(self, column: str,
                     method: str = "value", value: Any = "") -> str:
        """
        method:
          'value'   — fill with a fixed value
          'mean'    — fill with column mean (numeric only)
          'median'  — fill with column median (numeric only)
          'mode'    — fill with most frequent value
          'ffill'   — forward fill
          'bfill'   — backward fill
        """
        if column not in self.df.columns:
            return f"❌  Column `{column}` not found."
        before = self.df.copy()
        count  = int(self.df[column].isna().sum())

        if method == "value":
            self.df[column] = self.df[column].fillna(value)
        elif method == "mean":
            self.df[column] = self.df[column].fillna(self.df[column].mean())
        elif method == "median":
            self.df[column] = self.df[column].fillna(self.df[column].median())
        elif method == "mode":
            mode_val = self.df[column].mode()
            if not mode_val.empty:
                self.df[column] = self.df[column].fillna(mode_val[0])
        elif method == "ffill":
            self.df[column] = self.df[column].ffill()
        elif method == "bfill":
            self.df[column] = self.df[column].bfill()

        desc = f"❓  Filled {count} missing values in `{column}` using '{method}'."
        return self._commit("Fill Missing", before, desc)

    # ════════════════════════════════════════════════════════
    #  DATA ENTRY  — row / column / cell editing
    # ════════════════════════════════════════════════════════

    # ── Add row ──────────────────────────────────────────────
    def add_row(self, position: str = "bottom",
                after_index: Optional[int] = None) -> str:
        """
        position: 'bottom' | 'top' | 'after'
        If 'after', after_index must be given.
        """
        before   = self.df.copy()
        blank    = pd.DataFrame(
            [[pd.NA] * len(self.df.columns)],
            columns=self.df.columns)

        if position == "top":
            self.df = pd.concat([blank, self.df], ignore_index=True)
        elif position == "after" and after_index is not None:
            top    = self.df.iloc[:after_index + 1]
            bottom = self.df.iloc[after_index + 1:]
            self.df = pd.concat([top, blank, bottom], ignore_index=True)
        else:
            self.df = pd.concat([self.df, blank], ignore_index=True)

        desc = f"➕  Added 1 blank row ({position})."
        return self._commit("Add Row", before, desc)

    # ── Delete row ───────────────────────────────────────────
    def delete_row(self, row_index: int) -> str:
        if row_index < 0 or row_index >= len(self.df):
            return f"❌  Row index {row_index} out of range."
        before   = self.df.copy()
        self.df  = self.df.drop(index=row_index).reset_index(drop=True)
        desc     = f"🗑️   Deleted row {row_index}."
        return self._commit("Delete Row", before, desc)

    # ── Add column ───────────────────────────────────────────
    def add_column(self, name: str,
                   default_value: Any = pd.NA,
                   position: Optional[int] = None) -> str:
        if name in self.df.columns:
            return f"❌  Column `{name}` already exists."
        before = self.df.copy()
        if position is None or position >= len(self.df.columns):
            self.df[name] = default_value
        else:
            self.df.insert(position, name, default_value)
        desc = f"➕  Added column `{name}`."
        return self._commit("Add Column", before, desc)

    # ── Delete column ────────────────────────────────────────
    def delete_column(self, column: str) -> str:
        if column not in self.df.columns:
            return f"❌  Column `{column}` not found."
        before   = self.df.copy()
        self.df  = self.df.drop(columns=[column])
        desc     = f"🗑️   Deleted column `{column}`."
        return self._commit("Delete Column", before, desc)

    # ── Rename column ────────────────────────────────────────
    def rename_column(self, old_name: str, new_name: str) -> str:
        if old_name not in self.df.columns:
            return f"❌  Column `{old_name}` not found."
        if new_name in self.df.columns:
            return f"❌  Column `{new_name}` already exists."
        before   = self.df.copy()
        self.df  = self.df.rename(columns={old_name: new_name})
        # Update modified cells tracking
        self.modified_cells = {
            (r, new_name if c == old_name else c)
            for r, c in self.modified_cells
        }
        desc = f"✏️   Renamed column `{old_name}` → `{new_name}`."
        return self._commit("Rename Column", before, desc)

    # ── Edit cell ────────────────────────────────────────────
    def edit_cell(self, row_index: int, column: str, new_value: Any) -> str:
        """
        Update a single cell value.
        Validates type — if column is numeric, tries to cast.
        Marks cell as modified (gold highlight in UI).
        """
        if column not in self.df.columns:
            return f"❌  Column `{column}` not found."
        if row_index < 0 or row_index >= len(self.df):
            return f"❌  Row {row_index} out of range."

        before    = self.df.copy()
        old_value = self.df.at[row_index, column]

        # Type validation / coercion
        if pd.api.types.is_numeric_dtype(self.df[column]):
            try:
                new_value = float(new_value) if "." in str(new_value) else int(new_value)
            except (ValueError, TypeError):
                return (
                    f"⚠️  `{column}` is a numeric column. "
                    f"'{new_value}' is not a valid number."
                )

        self.df.at[row_index, column] = new_value
        self.modified_cells.add((row_index, column))

        desc = (
            f"✏️   Cell [{row_index}, `{column}`]:  "
            f"'{old_value}'  →  '{new_value}'."
        )
        return self._commit("Edit Cell", before, desc)

    # ── Clear cell ───────────────────────────────────────────
    def clear_cell(self, row_index: int, column: str) -> str:
        return self.edit_cell(row_index, column, pd.NA)

    # ════════════════════════════════════════════════════════
    #  ONE-CLICK CLEAN  — smart ordered pipeline
    # ════════════════════════════════════════════════════════

    def one_click_clean(self) -> list[str]:
        """
        Run all cleaning steps in the correct order.
        Returns list of log messages.
        """
        log = []

        # 1 — fix null tokens first (so later steps see clean NA)
        log.append(self.fix_nulls())

        # 2 — trim whitespace
        log.append(self.trim_whitespace())

        # 3 — remove junk rows
        log.append(self.remove_junk_rows())

        # 4 — drop fully blank rows
        log.append(self.drop_blank_rows())

        # 5 — drop duplicates
        log.append(self.drop_duplicates())

        # 6 — fix numeric columns
        log.append(self.fix_numeric_columns())

        # 7 — fix date columns
        log.append(self.fix_date_columns())

        # 8 — standardise column names
        log.append(self.standardise_column_names())

        # Filter out "nothing found" messages for cleaner log
        log = [l for l in log if l]
        return log

    # ════════════════════════════════════════════════════════
    #  STATS & COMPARISON
    # ════════════════════════════════════════════════════════

    def get_comparison(self) -> dict:
        """
        Returns before/after stats dict for the PDF report and UI.
        """
        if not self.stats_before:
            self._record_before()
        after = snapshot_stats(self.df)

        return {
            "before": self.stats_before or DataStats(0,0,0,0,0,0,0,""),
            "after":  after,
            "diff": {
                "rows_removed"    : (self.stats_before.rows if self.stats_before else 0) - after.rows,
                "missing_fixed"   : (self.stats_before.missing if self.stats_before else 0) - after.missing,
                "dupes_removed"   : (self.stats_before.duplicates if self.stats_before else 0) - after.duplicates,
                "cols_converted"  : (after.numeric_cols + after.date_cols)
                                  - ((self.stats_before.numeric_cols if self.stats_before else 0)
                                     + (self.stats_before.date_cols if self.stats_before else 0)),
            },
        }

    # ════════════════════════════════════════════════════════
    #  UTILITY
    # ════════════════════════════════════════════════════════

    def reset(self, original_df: pd.DataFrame):
        """Restore to original loaded data and clear history."""
        self.df             = original_df.copy()
        self.modified_cells = set()
        self.stack.clear()
        self._record_before()
        self.stats_after    = None

    def is_modified(self, row_index: int, column: str) -> bool:
        """Return True if this cell was manually edited (gold highlight)."""
        return (row_index, column) in self.modified_cells

    def column_summary(self, column: str) -> dict:
        """Quick stats for a single column — used in UI tooltips."""
        if column not in self.df.columns:
            return {}
        s = self.df[column]
        info = {
            "dtype"  : str(s.dtype),
            "missing": int(s.isna().sum()),
            "unique" : int(s.nunique()),
            "count"  : int(s.notna().sum()),
        }
        if pd.api.types.is_numeric_dtype(s):
            num = pd.to_numeric(s, errors="coerce").dropna()
            if len(num):
                info.update({
                    "min" : float(num.min()),
                    "max" : float(num.max()),
                    "mean": round(float(num.mean()), 4),
                    "std" : round(float(num.std()), 4),
                })
        else:
            top = s.value_counts().head(3)
            info["top_values"] = top.index.tolist()
        return info