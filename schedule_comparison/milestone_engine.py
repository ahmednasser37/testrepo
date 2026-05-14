"""
Milestone engine — extracts and compares milestones from baseline vs updated XER.
P6 milestones: task_type in ('TT_Mile','TT_FinMile','TT_StartMile') or duration==0.
"""
from __future__ import annotations
import pandas as pd


_MILESTONE_TYPES = {"TT_Mile", "TT_FinMile", "TT_StartMile"}


def _fmt(dt) -> str:
    if dt is None:
        return ""
    try:
        ts = pd.Timestamp(dt)
        return "" if pd.isna(ts) else ts.strftime("%Y-%m-%d")
    except Exception:
        return str(dt) if dt else ""


def _ms_status(pct: float, actual_finish: str, planned_finish: str, data_date) -> str:
    if pct >= 100 or actual_finish:
        return "complete"
    if not planned_finish:
        return "unknown"
    try:
        pf = pd.Timestamp(planned_finish)
        dd = pd.Timestamp(data_date) if data_date else pd.Timestamp.now()
        days = (pf - dd).days
        if days < 0:
            return "late"
        elif days <= 14:
            return "at_risk"
        else:
            return "on_track"
    except Exception:
        return "unknown"


def _extract_milestones(tables: dict) -> dict[str, dict]:
    task = tables.get("TASK", pd.DataFrame())
    wbs  = tables.get("WBS",  pd.DataFrame())

    wbs_names: dict[str, str] = {}
    if not wbs.empty and "wbs_id" in wbs.columns:
        wbs_names = dict(zip(wbs["wbs_id"].astype(str), wbs["wbs_name"].astype(str)))

    result: dict[str, dict] = {}
    if task.empty:
        return result

    for _, t in task.iterrows():
        ttype = str(t.get("task_type", ""))
        dur   = float(t.get("target_drtn_hr_cnt", 1) or 1)
        if ttype not in _MILESTONE_TYPES and dur > 0:
            continue

        code = str(t.get("task_code", ""))
        pct  = float(t.get("phys_complete_pct", 0) or 0)
        if pct <= 1.0:
            pct *= 100

        result[code] = {
            "task_code":      code,
            "task_name":      str(t.get("task_name", "")),
            "wbs_name":       wbs_names.get(str(t.get("wbs_id", "")), ""),
            "planned_finish": _fmt(t.get("target_end_date")),
            "actual_finish":  _fmt(t.get("act_end_date")),
            "pct_complete":   round(pct, 1),
        }
    return result


def compare_milestones(
    baseline_tables: dict,
    updated_tables: dict,
    data_date,
) -> list[dict]:
    """Return list of milestone comparison dicts."""
    b_ms = _extract_milestones(baseline_tables)
    u_ms = _extract_milestones(updated_tables)

    rows: list[dict] = []
    for code in sorted(set(b_ms) | set(u_ms)):
        b = b_ms.get(code, {})
        u = u_ms.get(code, {})

        bl_fin = b.get("planned_finish", "")
        up_fin = u.get("planned_finish", "")

        var_days = 0.0
        if bl_fin and up_fin:
            try:
                var_days = float((pd.Timestamp(up_fin) - pd.Timestamp(bl_fin)).days)
            except Exception:
                pass

        pct     = u.get("pct_complete", b.get("pct_complete", 0.0))
        act_fin = u.get("actual_finish", "")
        cur_fin = up_fin or bl_fin

        change = ("added" if not b else "deleted" if not u
                  else "changed" if var_days != 0 else "unchanged")

        rows.append({
            "task_code":              code,
            "task_name":              u.get("task_name", b.get("task_name", "")),
            "wbs_name":               u.get("wbs_name", b.get("wbs_name", "")),
            "baseline_finish":        bl_fin,
            "updated_finish":         up_fin,
            "actual_finish":          act_fin,
            "finish_variance_days":   var_days,
            "pct_complete":           round(pct, 1),
            "status":                 _ms_status(pct, act_fin, cur_fin, data_date),
            "change_type":            change,
        })
    return rows
