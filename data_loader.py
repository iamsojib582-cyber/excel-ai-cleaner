"""
╔══════════════════════════════════════════════════════════════╗
║            EXCEL AI CLEANER — by Sajeeb The Analyst         ║
║                     data_loader.py                          ║
║         Smart file loader — never crashes on messy files    ║
╚══════════════════════════════════════════════════════════════╝

Supports  : CSV, XLSX, XLS, JSON
Features  :
  • Auto-detects encoding  (utf-8, latin-1, cp1252 …)
  • Auto-detects separator (; , tab |)
  • Handles inconsistent column counts per row
  • Expands semicolon-packed Excel cells into proper columns
  • Replaces all known null-like tokens (NULL, #N/A, inf …)
  • Auto-converts obvious numeric / date columns
  • Returns a clean pd.DataFrame every time
"""

import json
import os
import re

import numpy as np
import pandas as pd


# ════════════════════════════════════════════════════════════
#  NULL-LIKE TOKENS
# ════════════════════════════════════════════════════════════

NULL_TOKENS = [
    "NULL", "null", "N/A", "NA", "n/a", "na",
    "None", "none", "NaN", "nan", "#N/A", "#n/a",
    "#NA", "?", "-", "", "INF", "inf", "-inf",
    "N\\A", "N/a", "na", "NIL", "nil", "NONE",
]


# ════════════════════════════════════════════════════════════
#  SEPARATOR DETECTION
# ════════════════════════════════════════════════════════════

def detect_separator(raw_text: str) -> str:
    """
    Scan first 20 lines and return the most frequent separator.
    Candidates: semicolon, comma, tab, pipe.
    """
    sample = "\n".join(raw_text.splitlines()[:20])
    scores = {
        ";":  sample.count(";"),
        ",":  sample.count(","),
        "\t": sample.count("\t"),
        "|":  sample.count("|"),
    }
    return max(scores, key=scores.get)


# ════════════════════════════════════════════════════════════
#  ENCODING DETECTION
# ════════════════════════════════════════════════════════════

def read_raw_text(path: str) -> str:
    """Try common encodings until one works. Returns raw text."""
    for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1252", "iso-8859-1"]:
        try:
            with open(path, "r", encoding=enc, errors="replace") as f:
                return f.read()
        except Exception:
            continue
    raise ValueError(f"Cannot read '{os.path.basename(path)}' — unsupported encoding.")


# ════════════════════════════════════════════════════════════
#  NULL + TYPE NORMALISATION
# ════════════════════════════════════════════════════════════

def normalise_nulls(df: pd.DataFrame) -> pd.DataFrame:
    """Replace all known null-like strings with pd.NA."""
    df = df.replace(NULL_TOKENS, pd.NA)
    return df


def auto_convert_types(df: pd.DataFrame) -> pd.DataFrame:
    """
    For each object column:
      • If >80% of values parse as numbers → convert to numeric
      • If >70% of values parse as dates   → convert to datetime
    """
    for col in df.select_dtypes(include=["object"]).columns:
        series = df[col].dropna().astype(str).str.strip()
        if series.empty:
            continue

        # Try numeric
        cleaned = series.str.replace(r"[$,\s%€£¥]", "", regex=True)
        num     = pd.to_numeric(cleaned, errors="coerce")
        if num.notna().mean() > 0.80:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(r"[$,\s%€£¥]", "", regex=True),
                errors="coerce",
            )
            continue

        # Try datetime
        try:
            parsed = pd.to_datetime(series, errors="coerce")
            if parsed.notna().mean() > 0.70:
                df[col] = pd.to_datetime(
                    df[col].astype(str), errors="coerce"
                )
        except Exception:
            pass

    return df


# ════════════════════════════════════════════════════════════
#  HEADER CLEANER
# ════════════════════════════════════════════════════════════

def clean_header(raw_header: list) -> list:
    """
    • Strip non-ASCII characters (encoding artifacts)
    • Strip whitespace
    • Replace spaces with underscores
    • Fill empty names as Col_0, Col_1 …
    • Deduplicate (Name, Name → Name, Name_1)
    """
    cleaned = []
    for i, h in enumerate(raw_header):
        h = re.sub(r"[^\x00-\x7F]", "", str(h)).strip().replace(" ", "_")
        if not h:
            h = f"Col_{i}"
        cleaned.append(h)

    # Deduplicate
    seen, final = {}, []
    for name in cleaned:
        if name in seen:
            seen[name] += 1
            final.append(f"{name}_{seen[name]}")
        else:
            seen[name] = 0
            final.append(name)
    return final


# ════════════════════════════════════════════════════════════
#  CSV LOADER
# ════════════════════════════════════════════════════════════

def load_csv(path: str) -> pd.DataFrame:
    """
    Robust CSV loader:
      1. Detect encoding
      2. Detect separator
      3. Strip junk-only lines (e.g. ;;;;;;;;)
      4. Normalise row widths (pad short / trim long rows)
      5. Build DataFrame
    """
    raw  = read_raw_text(path)
    sep  = detect_separator(raw)

    # Filter out junk lines
    lines = [
        line.strip()
        for line in raw.splitlines()
        if line.strip() and not re.fullmatch(r"[;,|\t\s]+", line.strip())
    ]

    if not lines:
        raise ValueError("File appears to be completely empty.")

    # Split rows
    rows  = [line.split(sep) for line in lines]

    # Find the most common row width → use as target width
    from collections import Counter
    width_counts = Counter(len(r) for r in rows)
    target_width = width_counts.most_common(1)[0][0]

    # Normalise row widths
    normalised = []
    for row in rows:
        if len(row) < target_width:
            row = row + [""] * (target_width - len(row))
        else:
            row = row[:target_width]
        normalised.append(row)

    # Build DataFrame
    header = clean_header(normalised[0])
    df     = pd.DataFrame(normalised[1:], columns=header)

    return df


# ════════════════════════════════════════════════════════════
#  EXCEL LOADER
# ════════════════════════════════════════════════════════════

def load_excel(path: str) -> pd.DataFrame:
    """
    Load XLSX / XLS.
    If the file has data crammed into 1–3 columns with semicolons,
    automatically expand into proper columns.
    """
    df = pd.read_excel(path)

    # Detect semicolon-packed layout
    if df.shape[1] <= 3:
        first_col = df.iloc[:, 0].astype(str)
        if first_col.str.contains(";").mean() > 0.50:
            # Reconstruct from packed cells
            combined = df.apply(
                lambda row: ";".join(
                    str(v) for v in row if str(v) not in ["nan", ""]
                ),
                axis=1,
            )
            rows  = [r.split(";") for r in combined]
            width = max(len(r) for r in rows)
            rows  = [(r + [""] * (width - len(r)))[:width] for r in rows]
            header = clean_header(rows[0])
            df     = pd.DataFrame(rows[1:], columns=header)

    return df


# ════════════════════════════════════════════════════════════
#  JSON LOADER
# ════════════════════════════════════════════════════════════

def load_json(path: str) -> pd.DataFrame:
    """Load JSON — supports nested objects via json_normalize."""
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)

    if isinstance(raw, list):
        df = pd.json_normalize(raw)
    elif isinstance(raw, dict):
        # Try to find the first list value (common API response pattern)
        for val in raw.values():
            if isinstance(val, list):
                df = pd.json_normalize(val)
                break
        else:
            df = pd.json_normalize([raw])
    else:
        raise ValueError("JSON structure not recognised.")

    return df


# ════════════════════════════════════════════════════════════
#  MAIN PUBLIC FUNCTION
# ════════════════════════════════════════════════════════════

def load_file(path: str) -> pd.DataFrame:
    """
    Load ANY supported file and return a clean DataFrame.

    Steps:
      1. Route to the correct loader by extension
      2. Normalise nulls
      3. Auto-convert types
      4. Return

    Raises:
      ValueError  — unsupported format or empty file
      Exception   — any other loading error (with clear message)
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")

    ext = os.path.splitext(path)[1].lower()

    try:
        if ext == ".csv":
            df = load_csv(path)
        elif ext in (".xlsx", ".xls"):
            df = load_excel(path)
        elif ext == ".json":
            df = load_json(path)
        else:
            raise ValueError(
                f"Unsupported file type '{ext}'.\n"
                "Please open a CSV, XLSX, XLS, or JSON file."
            )
    except (ValueError, FileNotFoundError):
        raise
    except Exception as e:
        raise RuntimeError(
            f"Could not open '{os.path.basename(path)}'.\n\nDetail: {e}"
        ) from e

    df = normalise_nulls(df)
    df = auto_convert_types(df)

    if df.empty:
        raise ValueError("File loaded but contains no data.")

    return df


# ════════════════════════════════════════════════════════════
#  FILE INFO HELPER
# ════════════════════════════════════════════════════════════

def file_info(path: str) -> dict:
    """
    Return a dict with basic file metadata — used in the UI status bar.
    {
        name     : str,
        ext      : str,
        size_kb  : float,
        full_path: str,
    }
    """
    stat = os.stat(path)
    return {
        "name":      os.path.basename(path),
        "ext":       os.path.splitext(path)[1].lower().lstrip(".").upper(),
        "size_kb":   round(stat.st_size / 1024, 1),
        "full_path": path,
    }


# ════════════════════════════════════════════════════════════
#  QUICK TEST  (run this file directly to test the loader)
# ════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage:  python data_loader.py  path/to/your/file.csv")
        sys.exit(0)

    test_path = sys.argv[1]
    print(f"\nLoading: {test_path}")
    print("─" * 50)

    try:
        df   = load_file(test_path)
        info = file_info(test_path)
        print(f"✅  File     : {info['name']}  ({info['size_kb']} KB)")
        print(f"✅  Shape    : {len(df):,} rows × {len(df.columns)} columns")
        print(f"✅  Missing  : {int(df.isna().sum().sum()):,} cells")
        print(f"✅  Dupes    : {int(df.duplicated().sum()):,} rows")
        print(f"\nColumns: {list(df.columns)}")
        print(f"\nFirst 3 rows:\n{df.head(3)}")
    except Exception as e:
        print(f"❌  Error: {e}")