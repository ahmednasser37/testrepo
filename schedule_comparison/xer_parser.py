"""
XER file parser — reads Primavera P6 .XER files into raw DataFrames.

Primary implementation uses the `xerparser` (PyP6XER) library.
Falls back to manual text parsing if xerparser fails (e.g. malformed XER).

XER format (for fallback):
    %T  TABLE_NAME
    %F  field1\tfield2\t...
    %R  val1\tval2\t...
    %E  (end of file)
"""

import tempfile
import os
import pandas as pd
from pathlib import Path


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def parse_xer(filepath: str | Path) -> dict[str, pd.DataFrame]:
    """Parse a .XER file and return a dict of table_name -> DataFrame."""
    filepath = Path(filepath)
    with open(filepath, encoding="utf-8", errors="replace") as f:
        content = f.read()
    return parse_xer_string(content)


def parse_xer_string(content: str) -> dict[str, pd.DataFrame]:
    """Parse XER content from a string."""
    try:
        return _parse_with_xerparser(content)
    except Exception:
        return _parse_manual(content)


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


# ---------------------------------------------------------------------------
# xerparser-based implementation
# ---------------------------------------------------------------------------

def _parse_with_xerparser(content: str) -> dict[str, pd.DataFrame]:
    """Parse XER content using the xerparser library."""
    from xerparser.reader import Reader

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".xer", delete=False, encoding="utf-8"
        ) as tmp:
            tmp.write(content)
            tmp_path = tmp.name

        reader = Reader(tmp_path)
        tables = {}

        tables["TASK"] = _build_task_df(reader)
        tables["PROJECT"] = _build_project_df(reader)
        tables["TASKPRED"] = _build_taskpred_df(reader)
        tables["PROJWBS"] = _build_projwbs_df(reader)
        tables["TASKRSRC"] = _build_taskrsrc_df(reader)

        # For tables with no data rows, preserve raw column schema from the XER
        # so that downstream code and tests see the exact field names.
        raw = _parse_manual(content)
        for name in list(tables.keys()):
            df = tables[name]
            if df.empty and name in raw:
                raw_df = raw[name]
                # Only substitute when the raw table also has no rows but does
                # have columns (i.e. it was a header-only table in the XER).
                if raw_df.empty and len(raw_df.columns) > 0:
                    tables[name] = raw_df

        _clean_tables(tables)
        return tables

    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


def _build_task_df(reader) -> pd.DataFrame:
    cols = [
        "task_id", "proj_id", "wbs_id", "clndr_id",
        "task_code", "task_name", "status_code", "task_type", "duration_type",
        "phys_complete_pct", "target_drtn_hr_cnt", "remain_drtn_hr_cnt",
        "total_float_hr_cnt", "free_float_hr_cnt",
        "act_work_qty", "target_work_qty",
        "target_start_date", "target_end_date",
        "act_start_date", "act_end_date",
        "early_start_date", "early_end_date",
        "late_start_date", "late_end_date",
        "restart_date", "reend_date",
    ]
    rows = []
    for t in reader.activities.activities:
        rows.append({
            "task_id":             t.task_id,
            "proj_id":             t.proj_id,
            "wbs_id":              t.wbs_id,
            "clndr_id":            t.clndr_id,
            "task_code":           t.task_code,
            "task_name":           t.task_name,
            "status_code":         t.status_code,
            "task_type":           t.task_type,
            "duration_type":       getattr(t, "duration_type", None),
            "phys_complete_pct":   t.phys_complete_pct,
            "target_drtn_hr_cnt":  t.target_drtn_hr_cnt,
            "remain_drtn_hr_cnt":  t.remain_drtn_hr_cnt,
            "total_float_hr_cnt":  t.total_float_hr_cnt,
            "free_float_hr_cnt":   t.free_float_hr_cnt,
            "act_work_qty":        t.act_work_qty,
            "target_work_qty":     t.target_work_qty,
            "target_start_date":   t.target_start_date,
            "target_end_date":     t.target_end_date,
            "act_start_date":      t.act_start_date,
            "act_end_date":        t.act_end_date,
            "early_start_date":    t.early_start_date,
            "early_end_date":      t.early_end_date,
            "late_start_date":     t.late_start_date,
            "late_end_date":       t.late_end_date,
            "restart_date":        t.restart_date,
            "reend_date":          t.reend_date,
        })
    return pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)


def _build_project_df(reader) -> pd.DataFrame:
    cols = [
        "proj_id", "proj_short_name", "proj_name",
        "plan_start_date", "plan_end_date",
        "last_recalc_date", "data_date",
    ]
    rows = []
    for p in reader.projects._projects:
        short_name = p.proj_short_name or ""
        last_recalc = p.last_recalc_date
        rows.append({
            "proj_id":          p.proj_id,
            "proj_short_name":  short_name,
            "proj_name":        short_name,
            "plan_start_date":  p.plan_start_date,
            "plan_end_date":    p.scd_end_date if getattr(p, "scd_end_date", None) else p.plan_end_date,
            "last_recalc_date": last_recalc,
            "data_date":        last_recalc,
        })
    return pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)


def _build_taskpred_df(reader) -> pd.DataFrame:
    cols = ["task_pred_id", "task_id", "pred_task_id", "pred_type", "lag_hr_cnt"]
    rows = []
    for rel in reader.relations.task_pred:
        rows.append({
            "task_pred_id": rel.task_pred_id,
            "task_id":      rel.task_id,
            "pred_task_id": rel.pred_task_id,
            "pred_type":    rel.pred_type,
            "lag_hr_cnt":   rel.lag_hr_cnt,
        })
    return pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)


def _build_projwbs_df(reader) -> pd.DataFrame:
    cols = ["wbs_id", "proj_id", "parent_wbs_id", "wbs_short_name", "wbs_name"]
    rows = []
    for w in reader.wbss._wbss:
        short_name = w.wbs_short_name or ""
        rows.append({
            "wbs_id":         w.wbs_id,
            "proj_id":        w.proj_id,
            "parent_wbs_id":  w.parent_wbs_id,
            "wbs_short_name": short_name,
            "wbs_name":       w.wbs_name if w.wbs_name else short_name,
        })
    return pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)


def _build_taskrsrc_df(reader) -> pd.DataFrame:
    cols = [
        "taskrsrc_id", "task_id", "proj_id",
        "target_qty", "act_reg_qty", "act_ot_qty",
        "remain_qty", "target_cost", "act_reg_cost",
        "act_ot_cost", "remain_cost", "cost_per_qty",
    ]
    rows = []
    for r in reader.activityresources._taskrsrc:
        rows.append({
            "taskrsrc_id":  r.taskrsrc_id,
            "task_id":      r.task_id,
            "proj_id":      r.proj_id,
            "target_qty":   r.target_qty,
            "act_reg_qty":  r.act_reg_qty,
            "act_ot_qty":   r.act_ot_qty,
            "remain_qty":   r.remain_qty,
            "target_cost":  r.target_cost,
            "act_reg_cost": r.act_reg_cost,
            "act_ot_cost":  r.act_ot_cost,
            "remain_cost":  r.remain_cost,
            "cost_per_qty": r.cost_per_qty,
        })
    return pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)


# ---------------------------------------------------------------------------
# Manual (fallback) text parsing implementation
# ---------------------------------------------------------------------------

def _parse_manual(content: str) -> dict[str, pd.DataFrame]:
    """Parse XER content from a string (manual fallback)."""
    tables: dict[str, pd.DataFrame] = {}
    current_table = None
    current_fields: list[str] = []
    current_rows: list[list[str]] = []

    for line in content.splitlines():
        if not line.strip():
            continue

        tag = line[:2]
        rest = line[3:] if len(line) > 3 else ""

        if tag == "%T":
            if current_table and current_fields:
                tables[current_table] = _build_df(current_fields, current_rows)
            current_table = rest.strip()
            current_fields = []
            current_rows = []

        elif tag == "%F":
            current_fields = rest.split("\t")
            current_fields = [f.strip() for f in current_fields if f.strip()]

        elif tag == "%R":
            values = rest.split("\t")
            while len(values) < len(current_fields):
                values.append("")
            current_rows.append(values[: len(current_fields)])

        elif tag == "%E":
            if current_table and current_fields:
                tables[current_table] = _build_df(current_fields, current_rows)
            break

    if current_table and current_fields and current_table not in tables:
        tables[current_table] = _build_df(current_fields, current_rows)

    _clean_tables(tables)
    return tables


def _build_df(fields: list[str], rows: list[list[str]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=fields)
    return pd.DataFrame(rows, columns=fields)


# ---------------------------------------------------------------------------
# Column type coercion (shared by both paths)
# ---------------------------------------------------------------------------

_DATE_COLS = {
    "PROJECT": ["plan_start_date", "plan_end_date", "last_recalc_date",
                "add_date", "data_date"],
    "TASK": ["target_start_date", "target_end_date",
             "act_start_date", "act_end_date",
             "early_start_date", "early_end_date",
             "late_start_date", "late_end_date",
             "restart_date", "reend_date",
             "expect_end_date", "suspend_date", "resume_date",
             "create_date"],
    "WBS": ["anticip_start_date", "anticip_end_date"],
    "TASKPRED": [],
}

_FLOAT_COLS = {
    "TASK": ["target_drtn_hr_cnt", "remain_drtn_hr_cnt",
             "act_work_qty", "target_work_qty",
             "phys_complete_pct", "total_cost_baseline",
             "target_cost", "act_cost",
             "total_float_hr_cnt", "free_float_hr_cnt"],
    "TASKRSRC": ["target_qty", "act_reg_qty", "act_ot_qty",
                 "remain_qty", "target_cost", "act_reg_cost",
                 "act_ot_cost", "remain_cost", "cost_per_qty"],
    "RSRC": ["cost_per_qty"],
    "TASKPRED": ["lag_hr_cnt"],
}


def _clean_tables(tables: dict[str, pd.DataFrame]) -> None:
    """Convert date and numeric columns to proper types."""
    for table_name, df in tables.items():
        date_cols = _DATE_COLS.get(table_name, [])
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")

        float_cols = _FLOAT_COLS.get(table_name, [])
        for col in float_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
