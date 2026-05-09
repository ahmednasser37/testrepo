"""
XER file parser — reads Primavera P6 .XER files into raw DataFrames.

XER format:
    %T  TABLE_NAME
    %F  field1\tfield2\t...
    %R  val1\tval2\t...
    %E  (end of file)
"""

import pandas as pd
from io import StringIO
from pathlib import Path


def parse_xer(filepath: str | Path) -> dict[str, pd.DataFrame]:
    """Parse a .XER file and return a dict of table_name -> DataFrame."""
    filepath = Path(filepath)
    with open(filepath, encoding="utf-8", errors="replace") as f:
        content = f.read()
    return parse_xer_string(content)


def parse_xer_string(content: str) -> dict[str, pd.DataFrame]:
    """Parse XER content from a string."""
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


def _build_df(fields: list[str], rows: list[list[str]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=fields)
    return pd.DataFrame(rows, columns=fields)


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

    keep = []
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
