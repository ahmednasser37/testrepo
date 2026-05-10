"""
Dashboard renderer — prepares chart data and renders the Jinja2 dashboard template.
"""

from __future__ import annotations

import json
import os
from pathlib import Path

from jinja2 import Environment, FileSystemLoader

from comparison_engine import ComparisonResult


_TEMPLATES_DIR = Path(__file__).parent / "templates"

_CHANGE_COLORS = {
    "added": "#22c55e",
    "deleted": "#ef4444",
    "changed": "#f59e0b",
    "unchanged": "#94a3b8",
}


def _fmt_date(val) -> str:
    if val is None:
        return "N/A"
    try:
        import pandas as pd
        t = pd.Timestamp(val)
        if pd.isna(t):
            return "N/A"
        return t.strftime("%Y-%m-%d")
    except Exception:
        return str(val)


def prepare_chart_data(result: ComparisonResult) -> dict:
    """Convert ComparisonResult into Plotly-ready JSON-serialisable dicts."""
    s = result.summary

    # Distribution bar chart
    labels = ["Added", "Deleted", "Changed", "Unchanged"]
    keys   = ["added", "deleted", "changed", "unchanged"]
    dist_values = [s.get(k, 0) for k in keys]
    dist_colors = [_CHANGE_COLORS[k] for k in keys]

    # Top delayed horizontal bar
    top_delayed = s.get("top_delayed", [])
    delay_codes = [d["task_code"] for d in top_delayed]
    delay_days  = [round(d["finish_variance_days"], 1) for d in top_delayed]

    # Scatter: baseline finish vs updated finish
    scatter_x: list[str] = []
    scatter_y: list[str] = []
    scatter_labels: list[str] = []
    scatter_types: list[str] = []

    for av in result.activity_variances:
        if av.baseline_finish and av.updated_finish:
            scatter_x.append(av.baseline_finish)
            scatter_y.append(av.updated_finish)
            scatter_labels.append(f"{av.task_code}: {av.task_name[:40]}")
            scatter_types.append(av.change_type)
        elif av.change_type == "added" and av.updated_finish:
            scatter_x.append(av.updated_finish)
            scatter_y.append(av.updated_finish)
            scatter_labels.append(f"{av.task_code}: {av.task_name[:40]} [ADDED]")
            scatter_types.append("added")

    return {
        "dist_labels": labels,
        "dist_values": dist_values,
        "dist_colors": dist_colors,
        "delay_codes": delay_codes,
        "delay_days":  delay_days,
        "scatter_x":      scatter_x,
        "scatter_y":      scatter_y,
        "scatter_labels": scatter_labels,
        "scatter_types":  scatter_types,
    }


def render_dashboard(
    result: ComparisonResult,
    ai_summary: dict,
    ck: str,
) -> str:
    """Render the full dashboard HTML string."""
    env = Environment(
        loader=FileSystemLoader(str(_TEMPLATES_DIR)),
        autoescape=True,
    )
    template = env.get_template("dashboard.html")

    chart_data = prepare_chart_data(result)

    return template.render(
        result=result,
        ai_summary=ai_summary,
        chart_data=chart_data,
        cache_key=ck,
        baseline_date=_fmt_date(result.baseline_data_date),
        updated_date=_fmt_date(result.updated_data_date),
    )
