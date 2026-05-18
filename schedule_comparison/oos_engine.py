"""
Out-of-Sequence (OOS) detection engine.
OOS: activity started before its FS predecessor finished.
"""
from __future__ import annotations
import pandas as pd


def detect_oos(updated_tables: dict, data_date) -> dict:
    """
    Detect OOS activities in the updated schedule.

    Returns:
        {
          "count": int,
          "items": [
            {
              "task_code": str,
              "task_name": str,
              "pred_code": str,
              "pred_name": str,
              "pred_type": str,
              "succ_actual_start": str,   # ISO date
              "pred_constrained_finish": str,  # ISO date (pred actual end + lag)
              "days_oos": float,
            }, ...
          ]
        }
    """
    task_df = updated_tables.get("TASK", pd.DataFrame())
    pred_df = updated_tables.get("TASKPRED", pd.DataFrame())

    if task_df.empty or pred_df.empty:
        return {"count": 0, "items": []}

    # Build task index by task_id
    task_idx = {}
    for _, row in task_df.iterrows():
        task_idx[str(row.get("task_id", ""))] = row

    dd = pd.Timestamp(data_date) if data_date else pd.Timestamp.now()
    oos_items = []

    for _, rel in pred_df.iterrows():
        # Determine pred_type column name (P6 uses PR_FS / FS)
        ptype = str(rel.get("pred_type", ""))
        # Only check FS relationships
        if ptype not in ("PR_FS", "FS", "pr_fs"):
            continue

        succ = task_idx.get(str(rel.get("task_id", "")))
        pred = task_idx.get(str(rel.get("pred_task_id", "")))
        if succ is None or pred is None:
            continue

        succ_act_start = pd.to_datetime(succ.get("act_start_date"), errors="coerce")
        if pd.isna(succ_act_start):
            continue  # successor hasn't started yet, not OOS

        pred_act_end = pd.to_datetime(pred.get("act_end_date"), errors="coerce")
        if pd.isna(pred_act_end):
            # Predecessor not finished — use data_date as its effective finish
            pred_effective_end = dd
        else:
            pred_effective_end = pred_act_end

        lag_days = float(rel.get("lag_hr_cnt", 0) or 0) / 8.0
        constrained_finish = pred_effective_end + pd.Timedelta(days=lag_days)

        if succ_act_start < constrained_finish:
            days_oos = round((constrained_finish - succ_act_start).total_seconds() / 86400, 1)
            oos_items.append({
                "task_code":  str(succ.get("task_code", "")),
                "task_name":  str(succ.get("task_name", "")),
                "pred_code":  str(pred.get("task_code", "")),
                "pred_name":  str(pred.get("task_name", "")),
                "pred_type":  ptype,
                "succ_actual_start":        succ_act_start.strftime("%Y-%m-%d"),
                "pred_constrained_finish":  constrained_finish.strftime("%Y-%m-%d"),
                "days_oos":   days_oos,
            })

    oos_items.sort(key=lambda x: x["days_oos"], reverse=True)
    return {"count": len(oos_items), "items": oos_items}
