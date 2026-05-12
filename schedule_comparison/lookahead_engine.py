"""
Lookahead engine — activities due in the next 2/4/6 weeks, grouped by WBS.
"""
from __future__ import annotations
from datetime import timedelta
import pandas as pd


def _fmt(dt) -> str:
    if dt is None:
        return ""
    try:
        ts = pd.Timestamp(dt)
        return "" if pd.isna(ts) else ts.strftime("%Y-%m-%d")
    except Exception:
        return str(dt) if dt else ""


def compute_lookahead(updated_tables: dict, updated_data: dict, data_date) -> dict:
    """
    Return JSON-serializable lookahead dict with 2/4/6-week windows and overdue items.
    data_date: pd.Timestamp or None
    """
    if data_date is None:
        data_date = pd.Timestamp.now().normalize()
    else:
        data_date = pd.Timestamp(data_date)

    activities = updated_data.get("activities", pd.DataFrame())

    # Float and criticality from TASK table
    task = updated_tables.get("TASK", pd.DataFrame())
    float_map: dict[str, float] = {}
    critical_map: dict[str, bool] = {}
    if not task.empty and "task_id" in task.columns:
        for _, r in task.iterrows():
            tid = str(r.get("task_id", ""))
            flt = float(r.get("total_float_hr_cnt", 0) or 0)
            float_map[tid] = round(flt / 8.0, 1)
            critical_map[tid] = flt <= 0

    if activities.empty:
        return _empty(data_date)

    wk2 = data_date + timedelta(weeks=2)
    wk4 = data_date + timedelta(weeks=4)
    wk6 = data_date + timedelta(weeks=6)

    overdue, two_week, four_week, six_week = [], [], [], []

    for _, act in activities.iterrows():
        status = str(act.get("status", "Not Started"))
        if status == "Completed":
            continue

        try:
            ps = pd.Timestamp(act.get("planned_start"))
            pf = pd.Timestamp(act.get("planned_finish"))
        except Exception:
            continue
        if pd.isna(ps) or pd.isna(pf):
            continue

        tid = str(act.get("task_id", ""))
        item = {
            "task_code":     str(act.get("task_code", "")),
            "task_name":     str(act.get("task_name", "")),
            "wbs_id":        str(act.get("wbs_id", "")),
            "wbs_name":      str(act.get("wbs_name", "")),
            "planned_start": _fmt(ps),
            "planned_finish": _fmt(pf),
            "duration_days": round(float(act.get("original_duration", 0)), 1),
            "pct_complete":  round(float(act.get("phys_complete_pct", 0)), 1),
            "is_critical":   critical_map.get(tid, False),
            "float_days":    float_map.get(tid, 0.0),
            "status":        status,
        }

        if pf < data_date:
            item["status"] = "Overdue"
            overdue.append(item)
        elif ps <= wk2:
            two_week.append(item)
        elif ps <= wk4:
            four_week.append(item)
        elif ps <= wk6:
            six_week.append(item)

    def _sort(lst):
        lst.sort(key=lambda x: (not x["is_critical"], x["planned_start"]))

    for lst in [overdue, two_week, four_week, six_week]:
        _sort(lst)

    all_items = overdue + two_week + four_week + six_week
    wbs_nodes = sorted({i["wbs_name"] for i in all_items if i["wbs_name"]})

    return {
        "data_date": _fmt(data_date),
        "overdue":   overdue,
        "two_week":  two_week,
        "four_week": four_week,
        "six_week":  six_week,
        "wbs_nodes": wbs_nodes,
        "counts": {
            "overdue":   len(overdue),
            "two_week":  len(two_week),
            "four_week": len(four_week),
            "six_week":  len(six_week),
        },
    }


def _empty(data_date) -> dict:
    return {
        "data_date": _fmt(data_date),
        "overdue": [], "two_week": [], "four_week": [], "six_week": [],
        "wbs_nodes": [],
        "counts": {"overdue": 0, "two_week": 0, "four_week": 0, "six_week": 0},
    }
