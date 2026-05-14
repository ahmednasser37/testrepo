"""
xer_parser.py — Robust XER file parser for p6-compare.

Ported from anti-gravity project (battle-tested V2).

Handles:
- BOM-aware encoding detection (UTF-8, UTF-16 LE/BE, Windows-1252 fallback)
- ERMHDR header capture
- %T/%F/%R/%E table grammar with malformed-line warnings
- Multiline continuation detection
- Returns: ParseResult(tables, header, warnings)

Adapter layer converts list-of-dicts tables → pandas DataFrames for
downstream compatibility with data_model.py and comparison_engine.py.
"""
from __future__ import annotations

import codecs
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import pandas as pd


# ── Public types ──────────────────────────────────────────────────────────────

@dataclass
class ParseResult:
    tables: Dict[str, List[Dict[str, str]]]
    header: Optional[str]           # Raw ERMHDR line, if present
    warnings: List[str]


# ── Encoding detection ────────────────────────────────────────────────────────

def _decode_bytes(raw: bytes) -> tuple[str, list[str]]:
    warnings: List[str] = []

    # Check BOMs
    if raw[:2] == b"\xff\xfe":
        encoding, payload = "utf-16-le", raw[2:]
    elif raw[:2] == b"\xfe\xff":
        encoding, payload = "utf-16-be", raw[2:]
    elif raw[:3] == b"\xef\xbb\xbf":
        encoding, payload = "utf-8-sig", raw[3:]
    else:
        # Heuristic: if first few bytes are ASCII-safe, try UTF-8, else fallback
        try:
            raw[:512].decode("utf-8")
            encoding, payload = "utf-8", raw
        except UnicodeDecodeError:
            encoding, payload = "windows-1252", raw
            warnings.append("UTF-8 decode failed during heuristic detection; retrying as windows-1252.")

    try:
        text = payload.decode(encoding)
    except (UnicodeDecodeError, LookupError) as exc:
        warnings.append(
            f"Encoding '{encoding}' decode failed ({exc}); "
            "retrying as windows-1252 with replacement."
        )
        text = payload.decode("windows-1252", errors="replace")
    return text, warnings


# ── Core parser ───────────────────────────────────────────────────────────────

def parse_xer_bytes(raw: bytes) -> ParseResult:
    """Parse raw XER bytes into a ParseResult (list-of-dicts tables)."""
    text, warnings = _decode_bytes(raw)
    return _parse_text(text, warnings)


def _parse_text(text: str, warnings: List[str]) -> ParseResult:
    tables: Dict[str, List[Dict[str, str]]] = {}
    header: Optional[str] = None
    current_table: Optional[str] = None
    current_fields: List[str] = []
    line_no = 0

    for raw_line in text.splitlines():
        line_no += 1
        line = raw_line.rstrip("\r\n")
        if not line:
            continue

        parts = line.split("\t")
        marker = parts[0]

        if marker == "ERMHDR":
            header = line
            continue

        if marker == "%T":
            if len(parts) < 2 or not parts[1].strip():
                warnings.append(f"Line {line_no}: malformed %T (no table name): {line!r}")
                current_table = None
                current_fields = []
                continue
            current_table = parts[1].strip()
            tables.setdefault(current_table, [])
            current_fields = []

        elif marker == "%F":
            if current_table is None:
                warnings.append(f"Line {line_no}: %F without preceding %T — ignored.")
                continue
            current_fields = parts[1:]

        elif marker == "%R":
            if current_table is None:
                warnings.append(f"Line {line_no}: %R without preceding %T — ignored.")
                continue
            if not current_fields:
                warnings.append(f"Line {line_no}: %R without preceding %F in table '{current_table}' — ignored.")
                continue
            row_values = parts[1:]
            if len(row_values) < len(current_fields):
                row_values = row_values + [""] * (len(current_fields) - len(row_values))
            row_values = row_values[: len(current_fields)]
            tables[current_table].append(dict(zip(current_fields, row_values)))

        elif marker == "%E":
            break
        else:
            # Multiline continuation: append to last field of last row
            if current_table and tables.get(current_table) and current_fields:
                warnings.append(f"Line {line_no}: Possible multiline continuation in table {current_table}.")
                last_row = tables[current_table][-1]
                last_field = current_fields[-1]
                last_row[last_field] = str(last_row.get(last_field, "")) + "\n" + line
            else:
                warnings.append(f"Line {line_no}: Unknown marker '{marker}' ignored.")

    return ParseResult(tables=tables, header=header, warnings=warnings)


# ── DataFrame adapter ─────────────────────────────────────────────────────────
#
# data_model.py and comparison_engine.py expect:
#   dict[str, pd.DataFrame]
#
# This adapter converts ParseResult.tables (list-of-dicts) → DataFrames,
# then applies typed coercions identical to the old _clean_tables logic.

_DATE_COLS = {
    "PROJECT":  ["plan_start_date", "plan_end_date", "last_recalc_date",
                 "add_date", "data_date"],
    "TASK":     ["target_start_date", "target_end_date",
                 "act_start_date", "act_end_date",
                 "early_start_date", "early_end_date",
                 "late_start_date", "late_end_date",
                 "restart_date", "reend_date",
                 "expect_end_date", "suspend_date", "resume_date", "create_date"],
    "WBS":      ["anticip_start_date", "anticip_end_date"],
    "TASKPRED": [],
}

_FLOAT_COLS = {
    "TASK":     ["target_drtn_hr_cnt", "remain_drtn_hr_cnt",
                 "act_work_qty", "target_work_qty",
                 "phys_complete_pct", "total_cost_baseline",
                 "target_cost", "act_cost",
                 "total_float_hr_cnt", "free_float_hr_cnt"],
    "TASKRSRC": ["target_qty", "act_reg_qty", "act_ot_qty",
                 "remain_qty", "target_cost", "act_reg_cost",
                 "act_ot_cost", "remain_cost", "cost_per_qty"],
    "RSRC":     ["cost_per_qty"],
    "TASKPRED": ["lag_hr_cnt"],
}


def to_dataframes(result: ParseResult) -> dict[str, pd.DataFrame]:
    """
    Convert a ParseResult into the dict[str, pd.DataFrame] format
    expected by data_model.py and comparison_engine.py.
    """
    dfs: dict[str, pd.DataFrame] = {}

    for table_name, rows in result.tables.items():
        if not rows:
            dfs[table_name] = pd.DataFrame()
            continue

        df = pd.DataFrame(rows)

        # Cast date columns
        for col in _DATE_COLS.get(table_name, []):
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")

        # Cast float columns
        for col in _FLOAT_COLS.get(table_name, []):
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

        dfs[table_name] = df

    return dfs


# ── Legacy compatibility shims ────────────────────────────────────────────────
# These functions maintain the API that app.py and comparison_engine.py call.

def parse_xer(filepath: str) -> dict[str, pd.DataFrame]:
    """Parse a .XER file path → DataFrames (used by app.py)."""
    with open(filepath, "rb") as f:
        raw = f.read()
    result = parse_xer_bytes(raw)
    return to_dataframes(result)


def parse_xer_string(content: str) -> dict[str, pd.DataFrame]:
    """Parse XER content from a string → DataFrames (used by app.py)."""
    result = _parse_text(content, [])
    return to_dataframes(result)


def parse_xer_bytes_to_df(raw: bytes) -> tuple[dict[str, pd.DataFrame], list[str]]:
    """
    Primary entry point for app.py:
    Returns (dataframes_dict, warnings_list).
    Use this instead of parse_xer() when you need the warnings too.
    """
    result = parse_xer_bytes(raw)
    return to_dataframes(result), result.warnings


def extract_data_date(tables: dict[str, pd.DataFrame]) -> pd.Timestamp | None:
    """Return the data_date from the PROJECT table, or None if unavailable."""
    project = tables.get("PROJECT", pd.DataFrame())
    if project.empty or "data_date" not in project.columns:
        return None
    val = project.iloc[0].get("data_date")
    if pd.isna(val):
        return None
    return pd.Timestamp(val) if not isinstance(val, pd.Timestamp) else val


def extract_project_meta(tables: dict[str, pd.DataFrame]) -> dict:
    """Return basic project metadata from the PROJECT table."""
    project = tables.get("PROJECT", pd.DataFrame())
    if project.empty:
        return {}
    row = project.iloc[0]
    return {
        "proj_id":    row.get("proj_id", ""),
        "proj_name":  row.get("proj_name", row.get("proj_short_name", "")),
        "data_date":  row.get("data_date"),
        "plan_start": row.get("plan_start_date"),
        "plan_end":   row.get("plan_end_date"),
    }


def get_relationships(tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    Return a normalized relationship DataFrame from TASKPRED.
    Columns: task_pred_id, task_id (successor), pred_task_id (predecessor),
             pred_type, lag_hr_cnt
    """
    taskpred = tables.get("TASKPRED", pd.DataFrame())
    if taskpred.empty:
        return pd.DataFrame(columns=[
            "task_pred_id", "task_id", "pred_task_id", "pred_type", "lag_hr_cnt"
        ])

    col_map = {
        "task_pred_id": ["task_pred_id", "taskpred_id"],
        "task_id":      ["task_id"],
        "pred_task_id": ["pred_task_id"],
        "pred_type":    ["pred_type"],
        "lag_hr_cnt":   ["lag_hr_cnt", "lag_drtn_hr_cnt"],
    }

    result: dict[str, pd.Series] = {}
    for canonical, candidates in col_map.items():
        for cand in candidates:
            if cand in taskpred.columns:
                result[canonical] = taskpred[cand]
                break
        if canonical not in result:
            result[canonical] = pd.Series([""] * len(taskpred))

    df = pd.DataFrame(result)
    df["lag_hr_cnt"] = pd.to_numeric(df["lag_hr_cnt"], errors="coerce").fillna(0.0)
    return df
