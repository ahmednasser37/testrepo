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
            # Save previous table if any
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
            # Pad or trim to match field count
            while len(values) < len(current_fields):
                values.append("")
            current_rows.append(values[: len(current_fields)])

        elif tag == "%E":
            if current_table and current_fields:
                tables[current_table] = _build_df(current_fields, current_rows)
            break

    # Flush last table if %E was missing
    if current_table and current_fields and current_table not in tables:
        tables[current_table] = _build_df(current_fields, current_rows)

    _clean_tables(tables)
    return tables


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
}

_FLOAT_COLS = {
    "TASK": ["target_drtn_hr_cnt", "remain_drtn_hr_cnt",
             "act_work_qty", "target_work_qty",
             "phys_complete_pct", "total_cost_baseline",
             "target_cost", "act_cost"],
    "TASKRSRC": ["target_qty", "act_reg_qty", "act_ot_qty",
                 "remain_qty", "target_cost", "act_reg_cost",
                 "act_ot_cost", "remain_cost", "cost_per_qty"],
    "RSRC": ["cost_per_qty"],
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
