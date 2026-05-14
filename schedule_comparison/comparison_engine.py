"""
Schedule comparison engine — pure deterministic logic, no AI, no web.
Compares a baseline XER schedule against an updated XER schedule.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

import pandas as pd

from xer_parser import extract_data_date, extract_project_meta, get_relationships
from data_model import process


# ── Dataclasses ───────────────────────────────────────────────────────────────

@dataclass
class ScheduleSnapshot:
    project_name: str
    project_id: str
    data_date: pd.Timestamp | None
    plan_start: pd.Timestamp | None
    plan_end: pd.Timestamp | None
    activities: pd.DataFrame
    relationships: pd.DataFrame
    wbs: pd.DataFrame
    resources: pd.DataFrame


@dataclass
class ActivityVariance:
    task_code: str
    task_name: str
    wbs_name: str
    change_type: str          # "added" | "deleted" | "changed" | "unchanged"
    start_variance_days: float
    finish_variance_days: float
    duration_variance_days: float
    pct_complete_change: float
    float_change_hours: float
    old_status: str
    new_status: str
    is_critical: bool         # total_float_hr_cnt <= 0 in updated schedule
    baseline_start: str       # ISO date string or ""
    baseline_finish: str      # ISO date string or ""
    updated_start: str
    updated_finish: str


@dataclass
class RelationshipVariance:
    pred_code: str
    succ_code: str
    change_type: str          # "added" | "deleted" | "changed"
    old_pred_type: str
    new_pred_type: str
    lag_change_hours: float


@dataclass
class ComparisonResult:
    baseline_project_name: str
    updated_project_name: str
    baseline_data_date: pd.Timestamp | None
    updated_data_date: pd.Timestamp | None
    plan_start: pd.Timestamp | None
    plan_end: pd.Timestamp | None
    data_date_warning: bool
    activity_variances: list[ActivityVariance] = field(default_factory=list)
    relationship_variances: list[RelationshipVariance] = field(default_factory=list)
    summary: dict = field(default_factory=dict)


# ── Snapshot builder ──────────────────────────────────────────────────────────

def build_snapshot(tables: dict) -> ScheduleSnapshot:
    """Parse raw XER tables into a ScheduleSnapshot."""
    meta = extract_project_meta(tables)
    data_date = extract_data_date(tables)

    datasets = process(tables)
    activities = datasets.get("activities", pd.DataFrame())
    wbs = datasets.get("wbs", pd.DataFrame())
    resources = datasets.get("resources", pd.DataFrame())

    relationships = get_relationships(tables)

    # Enrich activities with task_code → id mapping for relationship lookup
    task = tables.get("TASK", pd.DataFrame())
    if not task.empty and not activities.empty:
        id_to_code = {}
        if "task_id" in task.columns and "task_code" in task.columns:
            id_to_code = dict(zip(task["task_id"].astype(str), task["task_code"].astype(str)))
        if not relationships.empty:
            relationships = relationships.copy()
            relationships["pred_code"] = relationships["pred_task_id"].astype(str).map(id_to_code).fillna(relationships["pred_task_id"].astype(str))
            relationships["succ_code"] = relationships["task_id"].astype(str).map(id_to_code).fillna(relationships["task_id"].astype(str))

        # Merge float data from raw TASK table if available
        if "task_code" in task.columns and "total_float_hr_cnt" in task.columns:
            float_map = dict(zip(task["task_code"].astype(str), pd.to_numeric(task["total_float_hr_cnt"], errors="coerce").fillna(0.0)))
            if "task_code" in activities.columns:
                activities = activities.copy()
                activities["total_float_hr_cnt"] = activities["task_code"].astype(str).map(float_map).fillna(0.0)

    return ScheduleSnapshot(
        project_name=str(meta.get("proj_name", "")),
        project_id=str(meta.get("proj_id", "")),
        data_date=data_date,
        plan_start=meta.get("plan_start"),
        plan_end=meta.get("plan_end"),
        activities=activities,
        relationships=relationships,
        wbs=wbs,
        resources=resources,
    )


# ── Activity matching ─────────────────────────────────────────────────────────

def match_activities(
    baseline_acts: pd.DataFrame,
    updated_acts: pd.DataFrame,
) -> dict[str, str]:
    """
    Match baseline activities to updated activities.
    Returns dict: baseline_task_code -> updated_task_code (same value when matched).
    Primary key: task_code. Fallback: task_name when task_code absent.
    """
    if baseline_acts.empty or "task_code" not in baseline_acts.columns:
        return {}
    if updated_acts.empty or "task_code" not in updated_acts.columns:
        return {}

    baseline_codes = set(baseline_acts["task_code"].astype(str))
    updated_codes = set(updated_acts["task_code"].astype(str))

    matched: dict[str, str] = {}

    # Direct code match
    for code in baseline_codes & updated_codes:
        matched[code] = code

    # Fallback: name-based match for unmatched baseline codes
    unmatched_baseline = baseline_codes - set(matched)
    if unmatched_baseline and "task_name" in baseline_acts.columns and "task_name" in updated_acts.columns:
        updated_name_to_code = dict(zip(
            updated_acts["task_name"].astype(str),
            updated_acts["task_code"].astype(str),
        ))
        for code in unmatched_baseline:
            row = baseline_acts[baseline_acts["task_code"].astype(str) == code]
            if row.empty:
                continue
            name = str(row.iloc[0].get("task_name", ""))
            if name and name in updated_name_to_code:
                matched[code] = updated_name_to_code[name]

    return matched


# ── Helpers ───────────────────────────────────────────────────────────────────

def _days_between(a: object, b: object) -> float:
    """Return (b - a) in calendar days; positive means b is later than a."""
    try:
        ta = pd.Timestamp(a)
        tb = pd.Timestamp(b)
        if pd.isna(ta) or pd.isna(tb):
            return 0.0
        return (tb - ta).total_seconds() / 86400.0
    except Exception:
        return 0.0


def _fmt_date(val: object) -> str:
    try:
        t = pd.Timestamp(val)
        if pd.isna(t):
            return ""
        return t.strftime("%Y-%m-%d")
    except Exception:
        return ""


def _row_by_code(df: pd.DataFrame, code: str) -> pd.Series | None:
    if df.empty or "task_code" not in df.columns:
        return None
    mask = df["task_code"].astype(str) == code
    if not mask.any():
        return None
    return df[mask].iloc[0]


# ── Comparison engine ─────────────────────────────────────────────────────────

def _compare_activities(
    baseline: ScheduleSnapshot,
    updated: ScheduleSnapshot,
    match_map: dict[str, str],
) -> list[ActivityVariance]:
    """Build the full list of ActivityVariance records."""
    variances: list[ActivityVariance] = []

    baseline_codes = set(baseline.activities["task_code"].astype(str)) if not baseline.activities.empty else set()
    updated_codes = set(updated.activities["task_code"].astype(str)) if not updated.activities.empty else set()

    processed_updated: set[str] = set()

    # Matched and deleted
    for b_code in baseline_codes:
        u_code = match_map.get(b_code)
        b_row = _row_by_code(baseline.activities, b_code)
        if b_row is None:
            continue

        if u_code is None:
            # Deleted
            variances.append(ActivityVariance(
                task_code=b_code,
                task_name=str(b_row.get("task_name", "")),
                wbs_name=str(b_row.get("wbs_name", "")),
                change_type="deleted",
                start_variance_days=0.0,
                finish_variance_days=0.0,
                duration_variance_days=0.0,
                pct_complete_change=0.0,
                float_change_hours=0.0,
                old_status=str(b_row.get("status", "")),
                new_status="",
                is_critical=False,
                baseline_start=_fmt_date(b_row.get("planned_start")),
                baseline_finish=_fmt_date(b_row.get("planned_finish")),
                updated_start="",
                updated_finish="",
            ))
            continue

        u_row = _row_by_code(updated.activities, u_code)
        if u_row is None:
            continue
        processed_updated.add(u_code)

        start_var = _days_between(b_row.get("planned_start"), u_row.get("planned_start"))
        finish_var = _days_between(b_row.get("planned_finish"), u_row.get("planned_finish"))
        dur_b = float(b_row.get("original_duration", 0) or 0)
        dur_u = float(u_row.get("original_duration", 0) or 0)
        dur_var = dur_u - dur_b
        pct_b = float(b_row.get("phys_complete_pct", 0) or 0)
        pct_u = float(u_row.get("phys_complete_pct", 0) or 0)
        float_b = float(b_row.get("total_float_hr_cnt", 0) or 0)
        float_u = float(u_row.get("total_float_hr_cnt", 0) or 0)
        float_change = float_u - float_b

        threshold = 0.01
        is_changed = (
            abs(start_var) >= threshold
            or abs(finish_var) >= threshold
            or abs(dur_var) >= threshold
            or abs(pct_u - pct_b) >= threshold
            or abs(float_change) >= 1.0
            or str(b_row.get("status", "")) != str(u_row.get("status", ""))
        )

        variances.append(ActivityVariance(
            task_code=b_code,
            task_name=str(b_row.get("task_name", "")),
            wbs_name=str(b_row.get("wbs_name", "")),
            change_type="changed" if is_changed else "unchanged",
            start_variance_days=round(start_var, 2),
            finish_variance_days=round(finish_var, 2),
            duration_variance_days=round(dur_var, 2),
            pct_complete_change=round(pct_u - pct_b, 2),
            float_change_hours=round(float_change, 1),
            old_status=str(b_row.get("status", "")),
            new_status=str(u_row.get("status", "")),
            is_critical=float_u <= 0,
            baseline_start=_fmt_date(b_row.get("planned_start")),
            baseline_finish=_fmt_date(b_row.get("planned_finish")),
            updated_start=_fmt_date(u_row.get("planned_start")),
            updated_finish=_fmt_date(u_row.get("planned_finish")),
        ))

    # Added
    added_codes = updated_codes - processed_updated
    for u_code in added_codes:
        u_row = _row_by_code(updated.activities, u_code)
        if u_row is None:
            continue
        variances.append(ActivityVariance(
            task_code=u_code,
            task_name=str(u_row.get("task_name", "")),
            wbs_name=str(u_row.get("wbs_name", "")),
            change_type="added",
            start_variance_days=0.0,
            finish_variance_days=0.0,
            duration_variance_days=float(u_row.get("original_duration", 0) or 0),
            pct_complete_change=float(u_row.get("phys_complete_pct", 0) or 0),
            float_change_hours=float(u_row.get("total_float_hr_cnt", 0) or 0),
            old_status="",
            new_status=str(u_row.get("status", "")),
            is_critical=float(u_row.get("total_float_hr_cnt", 0) or 0) <= 0,
            baseline_start="",
            baseline_finish="",
            updated_start=_fmt_date(u_row.get("planned_start")),
            updated_finish=_fmt_date(u_row.get("planned_finish")),
        ))

    return variances


def _compare_relationships(
    baseline: ScheduleSnapshot,
    updated: ScheduleSnapshot,
) -> list[RelationshipVariance]:
    """Identify added, deleted, and changed logic ties."""
    variances: list[RelationshipVariance] = []

    if baseline.relationships.empty and updated.relationships.empty:
        return variances

    def rel_key(row: pd.Series) -> tuple[str, str]:
        return (str(row.get("pred_code", row.get("pred_task_id", ""))),
                str(row.get("succ_code", row.get("task_id", ""))))

    b_rels: dict[tuple, pd.Series] = {}
    if not baseline.relationships.empty:
        for _, r in baseline.relationships.iterrows():
            b_rels[rel_key(r)] = r

    u_rels: dict[tuple, pd.Series] = {}
    if not updated.relationships.empty:
        for _, r in updated.relationships.iterrows():
            u_rels[rel_key(r)] = r

    all_keys = set(b_rels) | set(u_rels)
    for key in all_keys:
        pred_code, succ_code = key
        if key in b_rels and key not in u_rels:
            r = b_rels[key]
            variances.append(RelationshipVariance(
                pred_code=pred_code, succ_code=succ_code,
                change_type="deleted",
                old_pred_type=str(r.get("pred_type", "")),
                new_pred_type="",
                lag_change_hours=0.0,
            ))
        elif key not in b_rels and key in u_rels:
            r = u_rels[key]
            variances.append(RelationshipVariance(
                pred_code=pred_code, succ_code=succ_code,
                change_type="added",
                old_pred_type="",
                new_pred_type=str(r.get("pred_type", "")),
                lag_change_hours=float(r.get("lag_hr_cnt", 0) or 0),
            ))
        else:
            br = b_rels[key]
            ur = u_rels[key]
            b_type = str(br.get("pred_type", ""))
            u_type = str(ur.get("pred_type", ""))
            b_lag = float(br.get("lag_hr_cnt", 0) or 0)
            u_lag = float(ur.get("lag_hr_cnt", 0) or 0)
            if b_type != u_type or abs(b_lag - u_lag) >= 0.5:
                variances.append(RelationshipVariance(
                    pred_code=pred_code, succ_code=succ_code,
                    change_type="changed",
                    old_pred_type=b_type,
                    new_pred_type=u_type,
                    lag_change_hours=round(u_lag - b_lag, 1),
                ))

    return variances


def _build_summary(
    av: list[ActivityVariance],
    rv: list[RelationshipVariance],
    baseline: ScheduleSnapshot,
    updated: ScheduleSnapshot,
) -> dict:
    added = [a for a in av if a.change_type == "added"]
    deleted = [a for a in av if a.change_type == "deleted"]
    changed = [a for a in av if a.change_type == "changed"]
    unchanged = [a for a in av if a.change_type == "unchanged"]

    delayed = [a for a in changed if a.finish_variance_days > 5]
    improved = [a for a in changed if a.finish_variance_days < -5]

    finish_vars = [a.finish_variance_days for a in changed if a.finish_variance_days != 0]
    max_delay = max((a.finish_variance_days for a in av if a.change_type != "deleted"), default=0.0)
    avg_finish_var = (sum(finish_vars) / len(finish_vars)) if finish_vars else 0.0

    critical_updated = [a for a in av if a.is_critical and a.change_type != "deleted"]

    top_delayed = sorted(
        [a for a in av if a.finish_variance_days > 0 and a.change_type != "deleted"],
        key=lambda x: x.finish_variance_days,
        reverse=True,
    )[:10]

    return {
        "total_baseline": len(baseline.activities),
        "total_updated": len(updated.activities),
        "added": len(added),
        "deleted": len(deleted),
        "changed": len(changed),
        "unchanged": len(unchanged),
        "delayed_activities": len(delayed),
        "improved_activities": len(improved),
        "critical_activities_updated": len(critical_updated),
        "max_delay_days": round(max_delay, 1),
        "avg_finish_variance_days": round(avg_finish_var, 1),
        "relationships_added": sum(1 for r in rv if r.change_type == "added"),
        "relationships_deleted": sum(1 for r in rv if r.change_type == "deleted"),
        "relationships_changed": sum(1 for r in rv if r.change_type == "changed"),
        "top_delayed": [
            {
                "task_code": a.task_code,
                "task_name": a.task_name,
                "finish_variance_days": a.finish_variance_days,
                "updated_finish": a.updated_finish,
            }
            for a in top_delayed
        ],
    }


# ── Main entry ────────────────────────────────────────────────────────────────

def compare_schedules(
    baseline_tables: dict,
    updated_tables: dict,
) -> ComparisonResult:
    """
    Full comparison pipeline.
    Returns a ComparisonResult with all variances and summary statistics.
    """
    baseline = build_snapshot(baseline_tables)
    updated = build_snapshot(updated_tables)

    b_date = baseline.data_date
    u_date = updated.data_date
    date_warning = (
        b_date is not None
        and u_date is not None
        and abs((b_date - u_date).total_seconds()) > 86400
    )

    match_map = match_activities(baseline.activities, updated.activities)
    av = _compare_activities(baseline, updated, match_map)
    rv = _compare_relationships(baseline, updated)
    summary = _build_summary(av, rv, baseline, updated)

    plan_start = baseline.plan_start or updated.plan_start
    plan_end = baseline.plan_end or updated.plan_end

    return ComparisonResult(
        baseline_project_name=baseline.project_name,
        updated_project_name=updated.project_name,
        baseline_data_date=b_date,
        updated_data_date=u_date,
        plan_start=plan_start,
        plan_end=plan_end,
        data_date_warning=date_warning,
        activity_variances=av,
        relationship_variances=rv,
        summary=summary,
    )
