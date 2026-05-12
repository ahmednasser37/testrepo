"""
Procurement engine — auto-detects procurement/supply WBS nodes
by name-matching and compares activities within them.
"""
from __future__ import annotations
import pandas as pd


_KEYWORDS = [
    "procur", "supply", "long lead", "short lead", "vendor",
    "purchase", "delivery", "logistic", "submittal",
    "long-lead", "short-lead", "material supply", "equipment supply",
]


def _fmt(dt) -> str:
    if dt is None:
        return ""
    try:
        ts = pd.Timestamp(dt)
        return "" if pd.isna(ts) else ts.strftime("%Y-%m-%d")
    except Exception:
        return str(dt) if dt else ""


def _is_proc(name: str) -> bool:
    nl = name.lower()
    return any(kw in nl for kw in _KEYWORDS)


def compute_procurement(
    baseline_tables: dict,
    updated_tables: dict,
    data_date,
) -> dict:
    """Return procurement analysis dict."""
    u_wbs  = updated_tables.get("WBS",  pd.DataFrame())
    u_task = updated_tables.get("TASK", pd.DataFrame())
    b_task = baseline_tables.get("TASK", pd.DataFrame())

    # Detect procurement WBS
    proc_wbs_ids: set[str] = set()
    detected: list[str] = []
    if not u_wbs.empty and "wbs_id" in u_wbs.columns:
        for _, row in u_wbs.iterrows():
            name = str(row.get("wbs_name", ""))
            if _is_proc(name):
                proc_wbs_ids.add(str(row["wbs_id"]))
                detected.append(name)

    if not proc_wbs_ids or u_task.empty:
        return {"detected_wbs_nodes": detected, "items": [],
                "summary": {"total": 0, "complete": 0, "in_progress": 0, "not_started": 0, "late": 0}}

    wbs_names = dict(zip(u_wbs["wbs_id"].astype(str), u_wbs["wbs_name"].astype(str)))

    # Baseline finish by task_code
    b_finish: dict[str, str] = {}
    if not b_task.empty:
        for _, t in b_task.iterrows():
            b_finish[str(t.get("task_code", ""))] = _fmt(t.get("target_end_date"))

    STATUS_MAP = {"TK_Complete": "Complete", "TK_Active": "In Progress", "TK_NotStart": "Not Started"}
    counts = {"complete": 0, "in_progress": 0, "not_started": 0, "late": 0}
    items: list[dict] = []

    for _, t in u_task.iterrows():
        wbs_id = str(t.get("wbs_id", ""))
        if wbs_id not in proc_wbs_ids:
            continue

        code    = str(t.get("task_code", ""))
        wbs_nm  = wbs_names.get(wbs_id, "")
        status  = STATUS_MAP.get(str(t.get("status_code", "")), "Not Started")
        pct     = float(t.get("phys_complete_pct", 0) or 0)
        if pct <= 1.0:
            pct *= 100

        bl_fin  = b_finish.get(code, "")
        up_fin  = _fmt(t.get("target_end_date"))
        act_fin = _fmt(t.get("act_end_date"))
        var     = 0.0
        if bl_fin and up_fin:
            try:
                var = float((pd.Timestamp(up_fin) - pd.Timestamp(bl_fin)).days)
            except Exception:
                pass

        if data_date and up_fin and not act_fin:
            try:
                if pd.Timestamp(up_fin) < pd.Timestamp(data_date) and pct < 100:
                    status = "Late"
            except Exception:
                pass

        wl = wbs_nm.lower()
        items.append({
            "task_code":          code,
            "task_name":          str(t.get("task_name", "")),
            "wbs_name":           wbs_nm,
            "baseline_finish":    bl_fin,
            "updated_finish":     up_fin,
            "actual_finish":      act_fin,
            "finish_variance_days": var,
            "pct_complete":       round(pct, 1),
            "status":             status,
            "is_long_lead":       "long" in wl,
            "is_short_lead":      "short" in wl,
            "lead_time_days":     round(float(t.get("target_drtn_hr_cnt", 0) or 0) / 8.0),
        })
        k = {"Complete": "complete", "In Progress": "in_progress",
              "Not Started": "not_started", "Late": "late"}.get(status, "not_started")
        counts[k] = counts.get(k, 0) + 1

    items.sort(key=lambda x: (x["status"] != "Late", x["updated_finish"] or ""))

    return {
        "detected_wbs_nodes": detected,
        "items": items,
        "summary": {"total": len(items), **counts},
    }
