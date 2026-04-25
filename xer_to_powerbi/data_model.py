"""
Business logic, calculations, and data transformations.
Converts raw XER tables (DataFrames) into the 5 output datasets.
"""

import re
import numpy as np
import pandas as pd
from datetime import datetime


# ── Helpers ───────────────────────────────────────────────────────────────────

def _to_dt(val) -> pd.Timestamp | None:
    if pd.isna(val):
        return None
    if isinstance(val, pd.Timestamp):
        return val
    try:
        return pd.Timestamp(str(val))
    except Exception:
        return None


def _clamp(val: float, lo: float = 0.0, hi: float = 100.0) -> float:
    return max(lo, min(hi, val))


def _extract_zone(task_code: str, wbs_name: str) -> str:
    """Extract zone label from task_code or wbs_name."""
    # task_code patterns: Z01-..., ZONE01-..., Z1-...
    m = re.search(r"Z(?:ONE)?[_\-]?(\d{1,2})", task_code, re.IGNORECASE)
    if m:
        return f"Zone {int(m.group(1)):02d}"

    # wbs_name patterns: "Zone 01", "Zone01", "منطقة 1"
    m = re.search(r"(?:Zone|منطقة)[_\s\-]?(\d{1,2})", wbs_name, re.IGNORECASE)
    if m:
        return f"Zone {int(m.group(1)):02d}"

    return "Unzoned"


def _status_label(code: str) -> str:
    mapping = {
        "TK_Complete":  "Completed",
        "TK_Active":    "In Progress",
        "TK_NotStart":  "Not Started",
    }
    return mapping.get(code, "Not Started")


def _planned_pct(row: pd.Series, data_date: pd.Timestamp) -> float:
    status = row["status"]
    ps = _to_dt(row["planned_start"])
    pf = _to_dt(row["planned_finish"])

    if status == "Completed":
        return 100.0
    if status == "Not Started":
        if ps is None or ps > data_date:
            return 0.0
        # Started but not marked active — use date-based calc
    if ps is None or pf is None:
        return 0.0
    span = (pf - ps).total_seconds()
    if span <= 0:
        return 100.0
    elapsed = (data_date - ps).total_seconds()
    return _clamp(elapsed / span * 100)


# ── WBS helpers ───────────────────────────────────────────────────────────────

def _build_wbs_level(wbs_df: pd.DataFrame) -> pd.DataFrame:
    """Assign wbs_level (1/2/3) based on parent chain depth."""
    parent_map = dict(zip(wbs_df["wbs_id"], wbs_df["parent_wbs_id"]))

    def depth(wid, memo={}):
        if wid in memo:
            return memo[wid]
        parent = parent_map.get(wid, "")
        if not parent or parent == wid:
            memo[wid] = 1
        else:
            memo[wid] = 1 + depth(parent, memo)
        return memo[wid]

    wbs_df = wbs_df.copy()
    wbs_df["wbs_level"] = wbs_df["wbs_id"].apply(depth)
    return wbs_df


def _wbs_name_map(wbs_df: pd.DataFrame) -> dict[str, str]:
    return dict(zip(wbs_df["wbs_id"], wbs_df["wbs_name"]))


# ── Sheet builders ────────────────────────────────────────────────────────────

def build_activities(tables: dict, data_date: pd.Timestamp) -> pd.DataFrame:
    task = tables.get("TASK", pd.DataFrame())
    wbs  = tables.get("WBS",  pd.DataFrame())

    if task.empty:
        return pd.DataFrame()

    wbs_names = _wbs_name_map(wbs) if not wbs.empty else {}

    # Normalise phys_complete_pct (XER stores as 0.65 = 65%)
    pcp_col = "phys_complete_pct"
    task = task.copy()
    if pcp_col in task.columns:
        task[pcp_col] = pd.to_numeric(task[pcp_col], errors="coerce").fillna(0.0)
        # Values > 1 are already in 0-100 range; <= 1 need ×100
        mask = task[pcp_col] <= 1.0
        task.loc[mask, pcp_col] = task.loc[mask, pcp_col] * 100
        task[pcp_col] = task[pcp_col].clip(0, 100)

    rows = []
    for _, t in task.iterrows():
        wbs_id   = t.get("wbs_id", "")
        wbs_name = wbs_names.get(str(wbs_id), "")
        t_code   = str(t.get("task_code", ""))

        status_raw = str(t.get("status_code", "TK_NotStart"))
        status = _status_label(status_raw)

        phys_pct = float(t.get(pcp_col, 0.0))

        row = {
            "task_id":             t.get("task_id", ""),
            "task_code":           t_code,
            "task_name":           t.get("task_name", ""),
            "wbs_id":              wbs_id,
            "wbs_name":            wbs_name,
            "status":              status,
            "planned_start":       _to_dt(t.get("target_start_date")),
            "planned_finish":      _to_dt(t.get("target_end_date")),
            "actual_start":        _to_dt(t.get("act_start_date")),
            "actual_finish":       _to_dt(t.get("act_end_date")),
            "original_duration":   float(t.get("target_drtn_hr_cnt", 0)) / 8,
            "remaining_duration":  float(t.get("remain_drtn_hr_cnt", 0)) / 8,
            "phys_complete_pct":   phys_pct,
            "zone":                _extract_zone(t_code, wbs_name),
        }
        rows.append(row)

    df = pd.DataFrame(rows)

    # planned_pct per activity
    df["planned_pct"] = df.apply(lambda r: _planned_pct(r, data_date), axis=1)
    df["variance_pct"] = df["phys_complete_pct"] - df["planned_pct"]

    # Weight — filled after cost data is merged; placeholder = equal weight
    n = len(df)
    df["weight"] = 1.0 / n if n > 0 else 0.0

    return df


def build_wbs_sheet(tables: dict, activities_df: pd.DataFrame,
                    resources_df: pd.DataFrame) -> pd.DataFrame:
    wbs = tables.get("WBS", pd.DataFrame())
    if wbs.empty:
        return pd.DataFrame()

    wbs = _build_wbs_level(wbs.copy())
    parent_map = dict(zip(wbs["wbs_id"], wbs["parent_wbs_id"]))

    # Aggregate costs from resource sheet
    cost_by_wbs: dict[str, tuple[float, float]] = {}
    if not resources_df.empty and not activities_df.empty:
        act_wbs = dict(zip(activities_df["task_id"], activities_df["wbs_id"]))
        for _, r in resources_df.iterrows():
            wid = str(act_wbs.get(r["task_id"], ""))
            tc  = float(r.get("target_cost", 0))
            ac  = float(r.get("act_cost", 0))
            if wid:
                existing = cost_by_wbs.get(wid, (0.0, 0.0))
                cost_by_wbs[wid] = (existing[0] + tc, existing[1] + ac)

    rows = []
    for _, w in wbs.iterrows():
        wid = str(w["wbs_id"])
        tc, ac = cost_by_wbs.get(wid, (0.0, 0.0))
        rows.append({
            "wbs_id":             wid,
            "wbs_name":           w.get("wbs_name", ""),
            "parent_wbs_id":      str(w.get("parent_wbs_id", "")),
            "wbs_level":          int(w.get("wbs_level", 1)),
            "total_planned_cost": tc,
            "total_actual_cost":  ac,
        })

    return pd.DataFrame(rows)


def build_resources(tables: dict, activities_df: pd.DataFrame) -> pd.DataFrame:
    taskrsrc = tables.get("TASKRSRC", pd.DataFrame())
    rsrc     = tables.get("RSRC",     pd.DataFrame())

    if taskrsrc.empty:
        return pd.DataFrame()

    # Resource master lookup
    rsrc_info: dict[str, dict] = {}
    if not rsrc.empty:
        for _, r in rsrc.iterrows():
            rsrc_info[str(r["rsrc_id"])] = {
                "rsrc_name":  r.get("rsrc_name", ""),
                "rsrc_type":  r.get("rsrc_type", ""),
                "unit_id":    r.get("unit_id", ""),
                "unit_price": float(r.get("cost_per_qty", 0)),
            }

    # Activity lookup
    act_info: dict[str, dict] = {}
    if not activities_df.empty:
        for _, a in activities_df.iterrows():
            act_info[str(a["task_id"])] = {
                "task_name":       a.get("task_name", ""),
                "wbs_id":          a.get("wbs_id", ""),
                "phys_complete_pct": float(a.get("phys_complete_pct", 0)),
            }

    rows = []
    for _, tr in taskrsrc.iterrows():
        tid    = str(tr.get("task_id", ""))
        rid    = str(tr.get("rsrc_id", ""))
        rinfo  = rsrc_info.get(rid, {})
        ainfo  = act_info.get(tid, {})

        unit_price  = rinfo.get("unit_price", float(tr.get("cost_per_qty", 0)))
        target_qty  = float(tr.get("target_qty", 0))
        act_qty     = float(tr.get("act_reg_qty", 0)) + float(tr.get("act_ot_qty", 0))
        remain_qty  = float(tr.get("remain_qty", 0))
        phys_pct    = ainfo.get("phys_complete_pct", 0.0)

        target_cost = target_qty * unit_price
        act_cost    = act_qty    * unit_price
        ev_cost     = target_cost * phys_pct / 100.0
        cv          = ev_cost - act_cost

        rows.append({
            "task_id":       tid,
            "task_name":     ainfo.get("task_name", ""),
            "wbs_id":        ainfo.get("wbs_id", ""),
            "rsrc_id":       rid,
            "rsrc_name":     rinfo.get("rsrc_name", ""),
            "rsrc_type":     rinfo.get("rsrc_type", ""),
            "unit_of_measure": rinfo.get("unit_id", ""),
            "unit_price":    unit_price,
            "target_qty":    target_qty,
            "act_qty":       act_qty,
            "remain_qty":    remain_qty,
            "target_cost":   target_cost,
            "act_cost":      act_cost,
            "ev_cost":       ev_cost,
            "cv":            cv,
            "is_material":   rinfo.get("rsrc_type", "") == "MT",
        })

    return pd.DataFrame(rows)


def build_project_info(tables: dict, activities_df: pd.DataFrame,
                       resources_df: pd.DataFrame,
                       data_date: pd.Timestamp) -> pd.DataFrame:
    project = tables.get("PROJECT", pd.DataFrame())

    if project.empty:
        proj_row = {}
    else:
        proj_row = project.iloc[0].to_dict()

    proj_id   = proj_row.get("proj_id",   "")
    proj_name = proj_row.get("proj_name", proj_row.get("proj_short_name", ""))
    plan_start = _to_dt(proj_row.get("plan_start_date"))
    plan_end   = _to_dt(proj_row.get("plan_end_date"))

    n_total     = len(activities_df)
    n_completed = int((activities_df["status"] == "Completed").sum())
    n_inprog    = int((activities_df["status"] == "In Progress").sum())

    # Costs
    if not resources_df.empty:
        bac        = resources_df["target_cost"].sum()
        ac_total   = resources_df["act_cost"].sum()
        ev_total   = resources_df["ev_cost"].sum()
    else:
        bac = ac_total = ev_total = 0.0

    # Weights based on cost — always normalised so Σweight == 1
    activities_df = activities_df.copy()
    if bac > 0 and not activities_df.empty:
        cost_by_task = (
            resources_df.groupby("task_id")["target_cost"].sum()
            if not resources_df.empty else pd.Series(dtype=float)
        )
        activities_df["_tc"] = activities_df["task_id"].map(cost_by_task).fillna(0)
        # Tasks with no cost get a proxy weight = average task cost
        zero_mask = activities_df["_tc"] == 0
        if zero_mask.any() and n_total > 0:
            activities_df.loc[zero_mask, "_tc"] = bac / n_total
        total_tc = activities_df["_tc"].sum()
        activities_df["weight"] = (
            activities_df["_tc"] / total_tc if total_tc > 0 else 1.0 / n_total
        )
    else:
        activities_df["weight"] = 1.0 / n_total if n_total > 0 else 0.0

    overall_planned = (activities_df["planned_pct"] * activities_df["weight"]).sum()
    overall_actual  = (activities_df["phys_complete_pct"] * activities_df["weight"]).sum()
    overall_variance = overall_actual - overall_planned

    # PV = planned value to data_date (weighted planned cost)
    pv = (activities_df["planned_pct"] / 100.0 * activities_df["weight"] * bac).sum()

    spi = ev_total / pv      if pv      > 0 else 0.0
    cpi = ev_total / ac_total if ac_total > 0 else 0.0
    eac = bac / cpi           if cpi     > 0 else bac
    vac = bac - eac

    row = {
        "project_id":              proj_id,
        "project_name":            proj_name,
        "data_date":               data_date,
        "planned_start":           plan_start,
        "planned_finish":          plan_end,
        "total_activities":        n_total,
        "completed_activities":    n_completed,
        "inprogress_activities":   n_inprog,
        "overall_planned_pct":     round(overall_planned, 2),
        "overall_actual_pct":      round(overall_actual,  2),
        "overall_variance_pct":    round(overall_variance, 2),
        "BAC":                     round(bac,        2),
        "PV":                      round(pv,         2),
        "EV":                      round(ev_total,   2),
        "AC":                      round(ac_total,   2),
        "SPI":                     round(spi,        6),
        "CPI":                     round(cpi,        6),
        "EAC":                     round(eac,        2),
        "VAC":                     round(vac,        2),
    }

    return pd.DataFrame([row])


def build_scurve(activities_df: pd.DataFrame,
                 resources_df: pd.DataFrame,
                 plan_start: pd.Timestamp,
                 plan_end: pd.Timestamp,
                 bac: float) -> pd.DataFrame:
    """Generate monthly S-Curve data."""
    if activities_df.empty or plan_start is None or plan_end is None:
        return pd.DataFrame()

    # Monthly period ends
    periods = pd.date_range(start=plan_start, end=plan_end, freq="ME")
    if len(periods) == 0:
        periods = pd.date_range(start=plan_start, end=plan_end, freq="MS")

    # Cost per task
    if not resources_df.empty:
        cost_by_task = resources_df.groupby("task_id").agg(
            target_cost=("target_cost", "sum"),
            act_cost=("act_cost", "sum"),
        ).reset_index()
    else:
        cost_by_task = pd.DataFrame(columns=["task_id", "target_cost", "act_cost"])

    acts = activities_df.merge(cost_by_task, on="task_id", how="left")
    acts["target_cost"] = acts["target_cost"].fillna(0)
    acts["act_cost"]    = acts["act_cost"].fillna(0)

    total_planned = acts["target_cost"].sum()
    total_actual  = acts["act_cost"].sum()

    rows = []
    cum_planned_cost = 0.0
    cum_actual_cost  = 0.0

    for period_end in periods:
        period_planned_cost = 0.0
        period_actual_cost  = 0.0

        for _, a in acts.iterrows():
            ps = a.get("planned_start")
            pf = a.get("planned_finish")
            tc = float(a.get("target_cost", 0))
            ac = float(a.get("act_cost", 0))
            phys = float(a.get("phys_complete_pct", 0)) / 100.0

            if ps is None or pf is None or tc == 0:
                continue

            span = (pf - ps).total_seconds()
            if span <= 0:
                if ps <= period_end:
                    period_planned_cost += tc
                continue

            # Planned cost earned up to period_end
            elapsed = min((period_end - ps).total_seconds(), span)
            elapsed = max(0, elapsed)
            frac    = elapsed / span
            period_planned_cost += tc * frac

            # Actual cost: distribute proportionally if finished, else partial
            astart = a.get("actual_start")
            aend   = a.get("actual_finish")
            if aend is not None and pd.notna(aend):
                if aend <= period_end:
                    period_actual_cost += ac
            elif astart is not None and pd.notna(astart):
                if astart <= period_end:
                    # Estimate actual cost proportionally
                    afrac = phys
                    period_actual_cost += ac * (
                        min((period_end - astart).total_seconds(), span) / span
                    )

        cum_planned_cost = period_planned_cost
        cum_actual_cost  = period_actual_cost

        rows.append({
            "period_date":       period_end.replace(day=1),
            "planned_cum_pct":   round(cum_planned_cost / total_planned * 100, 2)
                                 if total_planned > 0 else 0.0,
            "actual_cum_pct":    round(cum_actual_cost / total_actual * 100, 2)
                                 if total_actual  > 0 else 0.0,
            "planned_cum_cost":  round(cum_planned_cost,  2),
            "actual_cum_cost":   round(cum_actual_cost,   2),
        })

    return pd.DataFrame(rows)


# ── Main entry ────────────────────────────────────────────────────────────────

def process(tables: dict) -> dict[str, pd.DataFrame]:
    """
    Given raw XER tables, return a dict with keys:
    activities, wbs, resources, project_info, scurve
    """
    project = tables.get("PROJECT", pd.DataFrame())
    if not project.empty:
        data_date = pd.to_datetime(
            project.iloc[0].get("data_date"), errors="coerce"
        )
        if pd.isna(data_date):
            data_date = pd.Timestamp.now().normalize()
    else:
        data_date = pd.Timestamp.now().normalize()

    activities_df = build_activities(tables, data_date)
    resources_df  = build_resources(tables, activities_df)
    wbs_df        = build_wbs_sheet(tables, activities_df, resources_df)
    project_df    = build_project_info(tables, activities_df,
                                       resources_df, data_date)

    # Sync final normalised weights from project_df into activities for scurve
    if not project_df.empty:
        bac = float(project_df.iloc[0].get("BAC", 0))
        if bac > 0 and not resources_df.empty:
            cost_by_task = resources_df.groupby("task_id")["target_cost"].sum()
            activities_df = activities_df.copy()
            activities_df["_tc"] = activities_df["task_id"].map(cost_by_task).fillna(0)
            n = len(activities_df)
            zero_mask = activities_df["_tc"] == 0
            if zero_mask.any() and n > 0:
                activities_df.loc[zero_mask, "_tc"] = bac / n
            total_tc = activities_df["_tc"].sum()
            activities_df["weight"] = (
                activities_df["_tc"] / total_tc if total_tc > 0 else 1.0 / n
            )
    else:
        bac = 0.0

    plan_start = _to_dt(project.iloc[0].get("plan_start_date")) if not project.empty else None
    plan_end   = _to_dt(project.iloc[0].get("plan_end_date"))   if not project.empty else None

    scurve_df = build_scurve(activities_df, resources_df,
                             plan_start, plan_end, bac)

    # Drop any internal helper columns before export
    activities_df = activities_df.drop(
        columns=[c for c in activities_df.columns if c.startswith("_")],
        errors="ignore",
    )

    return {
        "activities":   activities_df,
        "wbs":          wbs_df,
        "resources":    resources_df,
        "project_info": project_df,
        "scurve":       scurve_df,
    }
