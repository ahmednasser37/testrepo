"""
Business logic, calculations, and data transformations.
Converts raw XER tables (DataFrames) into the 5 output datasets.
"""

import re
import numpy as np
import pandas as pd
from datetime import datetime


# Canonical schema for the activities DataFrame — used for schema'd empty returns
ACTIVITIES_COLUMNS = [
    "task_id", "task_code", "task_name", "wbs_id", "wbs_name", "status",
    "planned_start", "planned_finish", "actual_start", "actual_finish",
    "original_duration", "remaining_duration", "phys_complete_pct", "zone",
    "planned_pct", "variance_pct", "weight",
]


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


def _is_procurement(wbs_name: str) -> bool:
    """Detect if WBS node is related to procurement/supply."""
    keywords = ["procur", "supply", "long lead", "short lead", "vendor", "purchase", "delivery", "logistic", "material", "equip"]
    w = wbs_name.lower()
    return any(kw in w for kw in keywords)


def _status_label(code: str) -> str:
    mapping = {
        "TK_Complete":  "Completed",
        "TK_Active":    "In Progress",
        "TK_NotStart":  "Not Started",
    }
    return mapping.get(code, "Not Started")


def _planned_pct(row: pd.Series, data_date: pd.Timestamp) -> float:
    """
    Planned % complete at data_date, derived purely from baseline dates.
    Status is intentionally ignored — PV is about the PLAN, not actual progress.
    Early-completed activities (pf > data_date) correctly get < 100%.
    """
    ps = _to_dt(row["planned_start"])
    pf = _to_dt(row["planned_finish"])
    if ps is None or pf is None:
        return 0.0
    if data_date >= pf:
        return 100.0
    if data_date <= ps:
        return 0.0
    span = (pf - ps).total_seconds()
    if span <= 0:
        return 100.0
    return _clamp((data_date - ps).total_seconds() / span * 100)


# ── WBS helpers ───────────────────────────────────────────────────────────────

def _build_wbs_level(wbs_df: pd.DataFrame) -> pd.DataFrame:
    """Assign wbs_level (1/2/3+) based on parent chain depth."""
    parent_map = dict(zip(wbs_df["wbs_id"], wbs_df["parent_wbs_id"]))

    memo: dict = {}

    def depth(wid):
        if wid in memo:
            return memo[wid]
        parent = parent_map.get(wid, "")
        if not parent or parent == wid:
            memo[wid] = 1
        else:
            memo[wid] = 1 + depth(parent)
        return memo[wid]

    wbs_df = wbs_df.copy()
    wbs_df["wbs_level"] = wbs_df["wbs_id"].apply(depth)
    return wbs_df


def _wbs_name_map(wbs_df: pd.DataFrame) -> dict[str, str]:
    return dict(zip(wbs_df["wbs_id"], wbs_df["wbs_name"]))


def _rollup_wbs_costs(
    cost_by_wbs: dict[str, tuple[float, float]],
    wbs_df: pd.DataFrame,
) -> dict[str, tuple[float, float]]:
    """
    Propagate leaf-level costs up the full WBS subtree.
    Nodes are processed deepest-first so each parent accumulates
    its complete subtree total in a single pass.
    """
    parent_map = dict(zip(wbs_df["wbs_id"], wbs_df["parent_wbs_id"]))
    level_map  = dict(zip(wbs_df["wbs_id"], wbs_df["wbs_level"]))

    # Sort descending by level → leaves processed before parents
    ordered = sorted(wbs_df["wbs_id"].tolist(),
                     key=lambda w: level_map.get(w, 0), reverse=True)

    result = dict(cost_by_wbs)  # copy so we don't mutate the input

    for wid in ordered:
        parent = str(parent_map.get(wid, ""))
        if not parent or parent == wid or parent == "nan":
            continue
        child_tc, child_ac = result.get(wid, (0.0, 0.0))
        if child_tc == 0.0 and child_ac == 0.0:
            continue
        parent_tc, parent_ac = result.get(parent, (0.0, 0.0))
        result[parent] = (parent_tc + child_tc, parent_ac + child_ac)

    return result


# ── Sheet builders ────────────────────────────────────────────────────────────

def build_activities(tables: dict, data_date: pd.Timestamp) -> pd.DataFrame:
    task = tables.get("TASK", pd.DataFrame())
    wbs  = tables.get("PROJWBS", tables.get("WBS", pd.DataFrame()))

    # Return schema'd empty DataFrame so downstream functions can safely check .empty
    if task.empty:
        return pd.DataFrame(columns=ACTIVITIES_COLUMNS)

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
            "planned_start":       _to_dt(t.get("target_start_date") or t.get("early_start_date") or t.get("start_date")),
            "planned_finish":      _to_dt(t.get("target_end_date") or t.get("early_end_date") or t.get("finish_date")),
            "actual_start":        _to_dt(t.get("act_start_date")),
            "actual_finish":       _to_dt(t.get("act_end_date")),
            "original_duration":   float(t.get("target_drtn_hr_cnt", 0)) / 8,
            "remaining_duration":  float(t.get("remain_drtn_hr_cnt", 0)) / 8,
            "phys_complete_pct":   phys_pct,
            "zone":                _extract_zone(t_code, wbs_name),
            "is_procurement":      _is_procurement(wbs_name),
        }
        rows.append(row)

    df = pd.DataFrame(rows)

    # planned_pct per activity
    df["planned_pct"] = df.apply(lambda r: _planned_pct(r, data_date), axis=1)
    df["variance_pct"] = df["phys_complete_pct"] - df["planned_pct"]

    # Weight — placeholder = equal weight; recalculated after cost data is merged
    n = len(df)
    df["weight"] = 1.0 / n if n > 0 else 0.0

    return df


def build_wbs_sheet(tables: dict, activities_df: pd.DataFrame,
                    resources_df: pd.DataFrame) -> pd.DataFrame:
    wbs = tables.get("PROJWBS", tables.get("WBS", pd.DataFrame()))
    if wbs.empty:
        return pd.DataFrame()

    wbs = _build_wbs_level(wbs.copy())

    # Aggregate direct task costs at leaf WBS level
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

    # FIX (Codex P2): propagate child costs up to all ancestor WBS nodes
    cost_by_wbs = _rollup_wbs_costs(cost_by_wbs, wbs)

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
                "task_name":         a.get("task_name", ""),
                "wbs_id":            a.get("wbs_id", ""),
                "phys_complete_pct": float(a.get("phys_complete_pct", 0)),
                "is_procurement":    a.get("is_procurement", False),
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

        # FIX: Prioritize explicit cost fields from TASKRSRC if they exist
        target_cost = float(tr.get("target_cost") or tr.get("target_cost_amt") or 0)
        if target_cost == 0:
            target_cost = target_qty * unit_price
        
        act_cost = float(tr.get("act_reg_cost") or tr.get("act_reg_cost_amt") or 0) + \
                   float(tr.get("act_ot_cost") or tr.get("act_ot_cost_amt") or 0)
        if act_cost == 0:
            act_cost = act_qty * unit_price
            
        ev_cost     = target_cost * phys_pct / 100.0
        cv          = ev_cost - act_cost

        rows.append({
            "task_id":         tid,
            "task_name":       ainfo.get("task_name", ""),
            "wbs_id":          ainfo.get("wbs_id", ""),
            "rsrc_id":         rid,
            "rsrc_name":       rinfo.get("rsrc_name", ""),
            "rsrc_type":       rinfo.get("rsrc_type", ""),
            "unit_of_measure": rinfo.get("unit_id", ""),
            "unit_price":      unit_price,
            "target_qty":      target_qty,
            "act_qty":         act_qty,
            "remain_qty":      remain_qty,
            "target_cost":     target_cost,
            "act_cost":        act_cost,
            "ev_cost":         ev_cost,
            "cv":              cv,
            "is_material":     rinfo.get("rsrc_type", "") == "MT",
            "is_procurement":  ainfo.get("is_procurement", False),
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

    proj_id    = proj_row.get("proj_id",   "")
    proj_name  = proj_row.get("proj_name", proj_row.get("proj_short_name", ""))
    plan_start = _to_dt(proj_row.get("plan_start_date"))
    plan_end   = _to_dt(proj_row.get("plan_end_date"))

    # FIX (Codex P1): guard against empty/missing TASK table
    if activities_df.empty or "status" not in activities_df.columns:
        n_total = n_completed = n_inprog = 0
        overall_planned = overall_actual = overall_variance = 0.0
        pv = bac = ac_total = ev_total = 0.0
        spi = cpi = 0.0
        eac = vac = 0.0
        if not resources_df.empty:
            bac      = resources_df["target_cost"].sum()
            ac_total = resources_df["act_cost"].sum()
            ev_total = resources_df["ev_cost"].sum()
            eac      = bac
            vac      = 0.0
    else:
        n_total     = len(activities_df)
        n_completed = int((activities_df["status"] == "Completed").sum())
        n_inprog    = int((activities_df["status"] == "In Progress").sum())

        # Costs
        if not resources_df.empty:
            bac      = resources_df["target_cost"].sum()
            ac_total = resources_df["act_cost"].sum()
            ev_total = resources_df["ev_cost"].sum()
        else:
            bac = ac_total = ev_total = 0.0

        activities_df = activities_df.copy()
        cost_by_task = (
            resources_df.groupby("task_id")["target_cost"].sum()
            if not resources_df.empty else pd.Series(dtype=float)
        )
        activities_df["_tc"] = activities_df["task_id"].map(cost_by_task).fillna(0.0)

        # PV = Σ(budget_i × planned_pct_i / 100) — direct, no phantom weights
        pv = (activities_df["_tc"] * activities_df["planned_pct"] / 100.0).sum()

        # Cost-proportional weights for overall_planned/actual display metrics only
        total_tc = activities_df["_tc"].sum()
        if total_tc > 0:
            activities_df["weight"] = activities_df["_tc"] / total_tc
        else:
            activities_df["weight"] = 1.0 / n_total if n_total > 0 else 0.0

        overall_planned  = (activities_df["planned_pct"]       * activities_df["weight"]).sum()
        overall_actual   = (activities_df["phys_complete_pct"] * activities_df["weight"]).sum()
        overall_variance = overall_actual - overall_planned

        spi = ev_total / pv      if pv      > 0 else 0.0
        cpi = ev_total / ac_total if ac_total > 0 else 0.0
        eac = bac / cpi           if cpi     > 0 else bac
        vac = bac - eac

    row = {
        "project_id":             proj_id,
        "project_name":           proj_name,
        "data_date":              data_date,
        "planned_start":          plan_start,
        "planned_finish":         plan_end,
        "total_activities":       n_total,
        "completed_activities":   n_completed,
        "inprogress_activities":  n_inprog,
        "overall_planned_pct":    round(overall_planned,  2),
        "overall_actual_pct":     round(overall_actual,   2),
        "overall_variance_pct":   round(overall_variance, 2),
        "BAC":                    round(bac,        2),
        "PV":                     round(pv,         2),
        "EV":                     round(ev_total,   2),
        "AC":                     round(ac_total,   2),
        "SPI":                    round(spi,        6),
        "CPI":                    round(cpi,        6),
        "EAC":                    round(eac,        2),
        "VAC":                    round(vac,        2),
    }

    return pd.DataFrame([row])


def build_scurve(activities_df: pd.DataFrame,
                 resources_df: pd.DataFrame,
                 plan_start: pd.Timestamp,
                 plan_end: pd.Timestamp,
                 bac: float) -> pd.DataFrame:
    """Generate monthly S-Curve data."""
    if activities_df.empty:
        return pd.DataFrame()

    # FIX: Use activity date range if project dates are missing or narrow
    acts_start = activities_df["planned_start"].min()
    acts_finish = activities_df["planned_finish"].max()
    
    start_dt = plan_start if plan_start and plan_start < acts_start else acts_start
    end_dt = plan_end if plan_end and plan_end > acts_finish else acts_finish

    if pd.isna(start_dt) or pd.isna(end_dt):
        return pd.DataFrame()

    # Monthly period ends
    periods = pd.date_range(start=start_dt, end=end_dt, freq="ME")
    if len(periods) == 0:
        periods = pd.date_range(start=start_dt, end=end_dt + pd.Timedelta(days=31), freq="ME")

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
    if total_planned <= 0:
        # Fallback to sum of weights (which are 1/n or cost-based)
        total_planned = acts["weight"].sum()
        use_weight_as_cost = True
    else:
        use_weight_as_cost = False

    rows = []

    prev_planned_cum = 0.0
    prev_actual_cum = 0.0

    for period_end in periods:
        period_planned_cum = 0.0
        period_actual_cum  = 0.0

        for _, a in acts.iterrows():
            ps = a.get("planned_start")
            pf = a.get("planned_finish")
            
            if use_weight_as_cost:
                tc = float(a.get("weight", 0))
                # Approximate actual weight by physical %
                ac = tc * (float(a.get("phys_complete_pct", 0)) / 100.0)
            else:
                tc = float(a.get("target_cost", 0))
                ac = float(a.get("act_cost", 0))

            if ps is None or pf is None or tc <= 0:
                continue

            span = (pf - ps).total_seconds()
            if span <= 0:
                if ps <= period_end:
                    period_planned_cum += tc
                continue

            # Cumulative Planned cost earned up to period_end
            elapsed_p = min((period_end - ps).total_seconds(), span)
            elapsed_p = max(0, elapsed_p)
            period_planned_cum += tc * (elapsed_p / span)

            # Cumulative Actual cost: 
            astart = a.get("actual_start")
            aend   = a.get("actual_finish")
            if aend is not None and pd.notna(aend):
                if aend <= period_end:
                    period_actual_cum += ac
            elif astart is not None and pd.notna(astart):
                if astart <= period_end:
                    # Distribute AC based on elapsed time vs total span (or use phys %)
                    phys = float(a.get("phys_complete_pct", 0)) / 100.0
                    period_actual_cum += ac * phys
            
        # Periodic (this month) = Cumulative - Previous Month's Cumulative
        periodic_planned = period_planned_cum - prev_planned_cum
        periodic_actual  = period_actual_cum - prev_actual_cum
        
        # Guard against small floating point noise
        periodic_planned = max(0.0, periodic_planned)
        periodic_actual  = max(0.0, periodic_actual)

        rows.append({
            "period_date":       period_end.replace(day=1),
            "planned_cum_pct":   round(period_planned_cum / total_planned * 100, 2)
                                 if total_planned > 0 else 0.0,
            "actual_cum_pct":    round(period_actual_cum  / total_planned * 100, 2)
                                 if total_planned > 0 else 0.0,
            "planned_cum_cost":  round(period_planned_cum, 2),
            "actual_cum_cost":   round(period_actual_cum,  2),
            "planned_periodic_cost": round(periodic_planned, 2),
            "actual_periodic_cost":  round(periodic_actual, 2),
        })
        
        prev_planned_cum = period_planned_cum
        prev_actual_cum = period_actual_cum

    return pd.DataFrame(rows)


def build_resource_loading(tables: dict, activities_df: pd.DataFrame) -> dict:
    """Compute per-period (monthly) manpower and equipment loading."""
    taskrsrc = tables.get("TASKRSRC", pd.DataFrame())
    rsrc     = tables.get("RSRC",     pd.DataFrame())

    empty = {
        "manpower": [], "equipment": [],
        "manpower_peak": 0, "manpower_avg": 0,
        "equipment_peak": 0, "equipment_avg": 0,
    }

    if taskrsrc.empty or activities_df.empty:
        return empty

    # rsrc_type lookup: RT_Labor → manpower, RT_Equip → equipment
    rsrc_type_map: dict[str, str] = {}
    if not rsrc.empty and "rsrc_id" in rsrc.columns and "rsrc_type" in rsrc.columns:
        for _, r in rsrc.iterrows():
            rsrc_type_map[str(r["rsrc_id"])] = str(r.get("rsrc_type", ""))

    # Activity date lookup
    act_dates: dict[str, tuple] = {}
    for _, a in activities_df.iterrows():
        ps = a.get("planned_start")
        pf = a.get("planned_finish")
        if ps is not None and pf is not None and pd.notna(ps) and pd.notna(pf):
            act_dates[str(a["task_id"])] = (pd.Timestamp(ps), pd.Timestamp(pf))

    # Date range
    all_starts = [v[0] for v in act_dates.values()]
    all_ends   = [v[1] for v in act_dates.values()]
    if not all_starts:
        return empty

    range_start = min(all_starts)
    range_end   = max(all_ends)
    periods = pd.date_range(start=range_start, end=range_end, freq="ME")
    if len(periods) == 0:
        periods = pd.date_range(start=range_start,
                                end=range_end + pd.Timedelta(days=31), freq="ME")

    man_by_period: dict = {p: 0.0 for p in periods}
    eq_by_period:  dict = {p: 0.0 for p in periods}

    for _, tr in taskrsrc.iterrows():
        tid   = str(tr.get("task_id", ""))
        rid   = str(tr.get("rsrc_id", ""))
        rtype = rsrc_type_map.get(rid, "RT_Labor")
        qty   = float(tr.get("target_qty", 0) or 0)

        dates = act_dates.get(tid)
        if not dates or qty <= 0:
            continue

        ps, pf = dates
        span_days = max(1, (pf - ps).days)

        for p in periods:
            p_start = p.replace(day=1)
            p_end   = p

            overlap_start = max(ps, p_start)
            overlap_end   = min(pf, p_end)
            if overlap_start >= overlap_end:
                continue

            overlap_days = (overlap_end - overlap_start).days
            period_qty   = qty * (overlap_days / span_days)

            if rtype == "RT_Equip":
                eq_by_period[p]  = eq_by_period.get(p, 0)  + period_qty
            else:
                man_by_period[p] = man_by_period.get(p, 0) + period_qty

    def _to_list(by_period: dict) -> list[dict]:
        return [
            {"period": p.strftime("%Y-%m"), "qty": round(v, 0)}
            for p, v in sorted(by_period.items())
        ]

    def _stats(lst: list[dict]):
        qtys = [x["qty"] for x in lst if x["qty"] > 0]
        if not qtys:
            return 0.0, 0.0
        avg = sum(qtys) / len(qtys)
        pk  = max(qtys)
        return round(pk, 0), round(avg, 0)

    man_list = _to_list(man_by_period)
    eq_list  = _to_list(eq_by_period)
    man_peak, man_avg = _stats(man_list)
    eq_peak,  eq_avg  = _stats(eq_list)

    # Annotate each point
    for item in man_list:
        item["is_peak"]    = item["qty"] == man_peak and man_peak > 0
        item["above_avg"]  = item["qty"] > man_avg
    for item in eq_list:
        item["is_peak"]    = item["qty"] == eq_peak and eq_peak > 0
        item["above_avg"]  = item["qty"] > eq_avg

    return {
        "manpower":       man_list,
        "equipment":      eq_list,
        "manpower_peak":  man_peak,
        "manpower_avg":   man_avg,
        "equipment_peak": eq_peak,
        "equipment_avg":  eq_avg,
    }


# ── Main entry ────────────────────────────────────────────────────────────────

def process(tables: dict) -> dict[str, pd.DataFrame]:
    """
    Given raw XER tables, return a dict with keys:
    activities, wbs, resources, project_info, scurve
    """
    project = tables.get("PROJECT", pd.DataFrame())
    data_date = None
    if not project.empty:
        row = project.iloc[0]
        for col in ("last_recalc_date", "data_date"):
            val = row.get(col)
            if val is not None:
                try:
                    ts = pd.to_datetime(val, errors="coerce")
                    if not pd.isna(ts):
                        data_date = ts
                        break
                except Exception:
                    pass
    if data_date is None:
        data_date = pd.Timestamp.now().normalize()

    activities_df = build_activities(tables, data_date)
    resources_df  = build_resources(tables, activities_df)
    wbs_df        = build_wbs_sheet(tables, activities_df, resources_df)
    project_df    = build_project_info(tables, activities_df,
                                       resources_df, data_date)

    # Sync final cost-proportional weights into activities for downstream use
    if not project_df.empty and not activities_df.empty:
        bac = float(project_df.iloc[0].get("BAC", 0))
        if not resources_df.empty:
            cost_by_task = resources_df.groupby("task_id")["target_cost"].sum()
            activities_df = activities_df.copy()
            activities_df["_tc"] = activities_df["task_id"].map(cost_by_task).fillna(0.0)
            total_tc = activities_df["_tc"].sum()
            n = len(activities_df)
            if total_tc > 0:
                activities_df["weight"] = activities_df["_tc"] / total_tc
            else:
                activities_df["weight"] = 1.0 / n if n > 0 else 0.0
    else:
        bac = 0.0

    plan_start = _to_dt(project.iloc[0].get("plan_start_date")) if not project.empty else None
    plan_end   = _to_dt(project.iloc[0].get("plan_end_date"))   if not project.empty else None

    scurve_df = build_scurve(activities_df, resources_df,
                             plan_start, plan_end, bac)

    # Drop internal helper columns before export
    activities_df = activities_df.drop(
        columns=[c for c in activities_df.columns if c.startswith("_")],
        errors="ignore",
    )

    return {
        "activities":        activities_df,
        "wbs":               wbs_df,
        "resources":         resources_df,
        "project_info":      project_df,
        "scurve":            scurve_df,
        "resource_loading":  build_resource_loading(tables, activities_df),
    }
