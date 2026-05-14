"""
KPI engine — computes schedule performance indicators.
"""
from __future__ import annotations
import pandas as pd
from comparison_engine import ComparisonResult


def _row_by_code(df: pd.DataFrame, code: str) -> pd.Series | None:
    if df.empty or "task_code" not in df.columns:
        return None
    mask = df["task_code"].astype(str) == code
    if not mask.any():
        return None
    return df[mask].iloc[0]


def compute_kpis(
    result: ComparisonResult,
    baseline_data: dict,
    updated_data: dict,
    updated_tables: dict,
) -> dict:
    """Return JSON-serializable KPI dict. All values safe for JSON."""
    s = result.summary
    variances = result.activity_variances

    u_acts = updated_data.get("activities", pd.DataFrame())
    u_info = updated_data.get("project_info", pd.DataFrame())
    b_info = baseline_data.get("project_info", pd.DataFrame())

    def _pval(df, col):
        if df.empty or col not in df.columns:
            return None
        v = df.iloc[0].get(col)
        return None if (v is None or (isinstance(v, float) and pd.isna(v))) else v

    # Schedule delay: updated plan_end - baseline plan_end
    schedule_delay_days = 0.0
    b_end = _pval(b_info, "planned_finish")
    u_end = _pval(u_info, "planned_finish")
    if b_end and u_end:
        try:
            schedule_delay_days = round(
                (pd.Timestamp(str(u_end)) - pd.Timestamp(str(b_end))).days, 1
            )
        except Exception:
            pass

    # Weighted % complete
    pct_complete = 0.0
    if not u_acts.empty and "weight" in u_acts.columns:
        pct_complete = round(
            float((u_acts["phys_complete_pct"] * u_acts["weight"]).sum()), 1
        )
    elif not u_acts.empty and "original_duration" in u_acts.columns:
        total_dur = u_acts["original_duration"].sum()
        if total_dur > 0:
            pct_complete = round(
                float((u_acts["phys_complete_pct"] * u_acts["original_duration"]).sum() / total_dur), 1
            )

    # SPI calculation (Baseline-based)
    # Denominator = Planned progress based on original BASELINE dates
    # Numerator   = Actual progress in updated schedule
    spi_duration = None
    if not u_acts.empty:
        total_numerator = 0.0
        total_denominator = 0.0
        dd = pd.Timestamp(result.updated_data_date) if result.updated_data_date else pd.Timestamp.now()
        
        # Map task codes to weights/durations for fast lookup
        weight_map = dict(zip(u_acts["task_code"], u_acts["weight"]))
        dur_map = dict(zip(u_acts["task_code"], u_acts["original_duration"]))
        
        for v in variances:
            code = v.task_code
            w = weight_map.get(code, 1.0 / len(variances) if variances else 1.0)
            
            # Baseline planned %
            b_start = pd.to_datetime(v.baseline_start, errors="coerce")
            b_finish = pd.to_datetime(v.baseline_finish, errors="coerce")
            b_planned_pct = 0.0
            if pd.notna(b_start) and pd.notna(b_finish):
                span = (b_finish - b_start).total_seconds()
                if span > 0:
                    elapsed = (dd - b_start).total_seconds()
                    b_planned_pct = max(0.0, min(100.0, (elapsed / span) * 100))
                elif b_start <= dd:
                    b_planned_pct = 100.0
            
            # Actual %
            u_row = _row_by_code(u_acts, code)
            actual_pct = float(u_row.get("phys_complete_pct", 0)) if u_row is not None else 0.0
            
            total_numerator += (actual_pct / 100.0) * w
            total_denominator += (b_planned_pct / 100.0) * w
            
        if total_denominator > 0:
            spi_duration = round(total_numerator / total_denominator, 3)

    # Float stats from updated TASK table
    task = updated_tables.get("TASK", pd.DataFrame())
    float_days: list[float] = []
    if not task.empty and "total_float_hr_cnt" in task.columns:
        float_days = [float(v) / 8.0 for v in task["total_float_hr_cnt"]
                      if pd.notna(v) and v != ""]
    avg_float_days = round(sum(float_days) / len(float_days), 1) if float_days else 0.0
    near_critical_count = sum(1 for f in float_days if 0 < f <= 10)

    # Float consumption rate: avg change in float across all variances
    float_changes = [v.float_change_hours / 8.0 for v in variances
                     if v.float_change_hours != 0]
    avg_float_consumed = round(
        sum(float_changes) / len(float_changes), 2
    ) if float_changes else 0.0

    n = len(variances)
    behind   = sum(1 for v in variances if v.finish_variance_days > 5)
    improved = sum(1 for v in variances if v.finish_variance_days < -5)
    critical_total    = sum(1 for v in variances if v.is_critical)
    critical_at_risk  = sum(1 for v in variances if v.is_critical and v.change_type != "unchanged")

    on_time_start  = sum(1 for v in variances if abs(v.start_variance_days) <= 2 and v.change_type != "added")
    on_time_finish = sum(1 for v in variances if abs(v.finish_variance_days) <= 2 and v.change_type not in ("added", "deleted"))

    return {
        # 4 priority KPIs
        "float_consumption_days":   avg_float_consumed,
        "pct_complete_weighted":    pct_complete,
        "spi_duration":             spi_duration,
        "schedule_delay_days":      schedule_delay_days,
        # Float
        "avg_float_days":           avg_float_days,
        "near_critical_count":      near_critical_count,
        # Activity counts
        "critical_total":           critical_total,
        "critical_at_risk":         critical_at_risk,
        "activities_behind":        behind,
        "activities_improved":      improved,
        "activities_on_track":      max(0, n - behind - improved),
        "on_time_start_rate":       round(on_time_start  / n * 100, 1) if n > 0 else None,
        "on_time_finish_rate":      round(on_time_finish / n * 100, 1) if n > 0 else None,
        # From summary
        "total_activities":         s.get("total_updated", 0),
        "added":                    s.get("added", 0),
        "deleted":                  s.get("deleted", 0),
        "changed":                  s.get("changed", 0),
        "unchanged":                s.get("unchanged", 0),
        "max_delay_days":           round(float(s.get("max_delay_days", 0)), 1),
        "avg_finish_variance_days": round(float(s.get("avg_finish_variance_days", 0)), 1),
    }
