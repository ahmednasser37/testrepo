"""
Flask web application — upload two XER files, compare, render 10-tab dashboard.
"""
from __future__ import annotations

import csv
import io
import json
import os
import threading
import uuid
from datetime import datetime
from pathlib import Path

import pandas as pd

from flask import Flask, request, render_template, Response, redirect, url_for, stream_with_context, jsonify

from xer_parser import parse_xer_bytes_to_df, extract_data_date
from comparison_engine import compare_schedules
from data_model import process as dm_process
from ai_service import cache_key, get_ai_summary, stream_chat_response
from dashboard_renderer import prepare_chart_data
from lookahead_engine import compute_lookahead
from kpi_engine import compute_kpis
from milestone_engine import compare_milestones
from procurement_engine import compute_procurement
from oos_engine import detect_oos
from path_engine import compute_longest_path

app = Flask(__name__)
app.secret_key = os.environ.get("SCE_SECRET_KEY", "dev-secret-change-in-production")

_UPLOAD_MAX_MB = int(os.environ.get("SCE_UPLOAD_MAX_MB", 200))
app.config["MAX_CONTENT_LENGTH"] = _UPLOAD_MAX_MB * 1024 * 1024

_CACHE_DIR         = Path(os.environ.get("SCE_CACHE_DIR", "/tmp/sce_cache"))
_OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY")
_OPENROUTER_MODEL  = os.environ.get("OPENROUTER_MODEL", "deepseek/deepseek-chat-v3-0324:free")

# ── Background job store ───────────────────────────────────────────────────────
# Keyed by job_id (uuid hex). Each entry: {status, pct, step, redirect?, error?}
_jobs: dict[str, dict] = {}
_jobs_lock = threading.Lock()


def _job_update(job_id: str, **kwargs):
    with _jobs_lock:
        _jobs[job_id].update(kwargs)


def _run_job(job_id: str, baseline_bytes: bytes, updated_bytes: bytes,
             ck: str, key16: str, dest: str):
    """Background thread: parse, compare, build dashboard, save to cache."""
    try:
        _job_update(job_id, pct=5,  step="Parsing baseline XER…")
        baseline_tables, bl_warnings = parse_xer_bytes_to_df(baseline_bytes)
        del baseline_bytes

        _job_update(job_id, pct=20, step="Parsing updated XER…")
        updated_tables, up_warnings = parse_xer_bytes_to_df(updated_bytes)
        del updated_bytes

        for name, tables, label in [
            ("PROJECT", baseline_tables, "Baseline"),
            ("TASK",    baseline_tables, "Baseline"),
            ("PROJECT", updated_tables,  "Updated"),
            ("TASK",    updated_tables,  "Updated"),
        ]:
            if name not in tables or tables[name].empty:
                _job_update(job_id, status="error",
                            error=f"{label} file is missing the {name} table.")
                return

        _job_update(job_id, pct=35, step="Comparing schedules…")
        _job_update(job_id, pct=50, step="Computing EVM & KPIs…")
        _job_update(job_id, pct=65, step="Detecting out-of-sequence activities…")
        _job_update(job_id, pct=75, step="Computing longest path…")

        full_data = _build_full_data(ck, baseline_tables, updated_tables)

        _job_update(job_id, pct=92, step="Saving dashboard…")
        _save_full(key16, full_data)

        _job_update(job_id, status="done", pct=100,
                    step="Done! Redirecting…", redirect=dest)

    except Exception as exc:
        _job_update(job_id, status="error", error=str(exc))


# ── Helpers ───────────────────────────────────────────────────────────────────

def _allowed(filename: str) -> bool:
    return Path(filename).suffix.lower() == ".xer"


def _full_path(key16: str) -> Path:
    _CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return _CACHE_DIR / f"{key16}_full.json"


def _load_full(key16: str) -> dict | None:
    p = _full_path(key16)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


def _save_full(key16: str, data: dict) -> None:
    try:
        _full_path(key16).write_text(
            json.dumps(data, ensure_ascii=False, default=str), encoding="utf-8"
        )
    except Exception:
        pass


def _df_to_rows(df) -> list[dict]:
    if df is None or df.empty:
        return []
    return json.loads(df.to_json(orient="records", date_format="iso", default_handler=str))


def _build_full_data(
    ck: str,
    baseline_tables: dict,
    updated_tables: dict,
) -> dict:
    """Run all engines and return a single JSON-serializable dict."""
    key16 = ck[:16]

    # Core comparison
    result      = compare_schedules(baseline_tables, updated_tables)
    data_date   = extract_data_date(updated_tables)
    b_date      = extract_data_date(baseline_tables)

    # data_model processing for both schedules
    baseline_dm = dm_process(baseline_tables)
    updated_dm  = dm_process(updated_tables)

    # EV / project info
    b_proj_info = _df_to_rows(baseline_dm.get("project_info"))
    u_proj_info = _df_to_rows(updated_dm.get("project_info"))
    b_pi = b_proj_info[0] if b_proj_info else {}
    u_pi = u_proj_info[0] if u_proj_info else {}

    has_cost_data = (
        not updated_dm.get("resources", None) is None
        and not updated_dm["resources"].empty
        and float(u_pi.get("BAC", 0)) > 0
    )

    # S-curves
    scurve_b = _df_to_rows(baseline_dm.get("scurve"))
    scurve_u = _df_to_rows(updated_dm.get("scurve"))

    # WBS summary
    wbs_df = updated_dm.get("wbs")
    wbs_summary = _df_to_rows(wbs_df)

    # Engine outputs
    kpis        = compute_kpis(result, baseline_dm, updated_dm, updated_tables)
    lookahead   = compute_lookahead(updated_tables, updated_dm, data_date)
    milestones  = compare_milestones(baseline_tables, updated_tables, data_date)
    procurement = compute_procurement(baseline_tables, updated_tables, data_date)

    # OOS detection
    try:
        oos = detect_oos(updated_tables, data_date)
    except Exception as _e:
        app.logger.warning("detect_oos failed: %s", _e)
        oos = {"count": 0, "items": []}

    # Longest path / bottleneck
    try:
        bottleneck = compute_longest_path(updated_tables, data_date)
    except Exception as _e:
        app.logger.warning("compute_longest_path failed: %s", _e)
        bottleneck = {"bottleneck": None, "driving_chain": []}

    # Full relationships list (all updated relationships for milestone trace + gantt lines)
    all_relationships = []
    try:
        from xer_parser import get_relationships as _get_rels
        _rels_df = _get_rels(updated_tables)
        _task_raw = updated_tables.get("TASK", pd.DataFrame())
        if not _rels_df.empty and not _task_raw.empty and "task_id" in _task_raw.columns:
            _id2code = dict(zip(_task_raw["task_id"].astype(str), _task_raw["task_code"].astype(str)))
            _rels_df = _rels_df.copy()
            if "pred_task_id" in _rels_df.columns:
                _rels_df["pred_code"] = _rels_df["pred_task_id"].astype(str).map(_id2code).fillna("")
            if "task_id" in _rels_df.columns:
                _rels_df["succ_code"] = _rels_df["task_id"].astype(str).map(_id2code).fillna("")
        if not _rels_df.empty:
            for _, _r in _rels_df.iterrows():
                all_relationships.append({
                    "pred_code": str(_r.get("pred_code", _r.get("pred_task_id", ""))),
                    "succ_code": str(_r.get("succ_code", _r.get("task_id", ""))),
                    "pred_type": str(_r.get("pred_type", "")),
                    "lag_days":  round(float(_r.get("lag_hr_cnt", 0) or 0) / 8.0, 1),
                })
    except Exception as _e:
        app.logger.warning("all_relationships build failed: %s", _e)

    # Resource loading (manpower + equipment per period)
    resource_loading = updated_dm.get("resource_loading", {})

    # Gantt activities (sorted by WBS for gantt chart rendering)
    _acts_df = updated_dm.get("activities", None)
    _task_df = updated_tables.get("TASK", None)
    gantt = []
    if _acts_df is not None and not _acts_df.empty:
        # Float and critical flag from raw TASK table
        _float_map: dict[str, float] = {}
        _critical_map: dict[str, bool] = {}
        if _task_df is not None and not _task_df.empty:
            for _, _t in _task_df.iterrows():
                _tid = str(_t.get("task_id", ""))
                _float_map[_tid]    = float(_t.get("total_float_hr_cnt", 0) or 0) / 8.0
                _critical_map[_tid] = str(_t.get("driving_path_flag", "")) == "Y"

        _task_type_map: dict[str, str] = {}
        if _task_df is not None and not _task_df.empty and "task_type" in _task_df.columns:
            for _, _t in _task_df.iterrows():
                _task_type_map[str(_t.get("task_id", ""))] = str(_t.get("task_type", ""))

        for _, _a in _acts_df.iterrows():
            _tid = str(_a.get("task_id", ""))
            def _s(v): return v.isoformat() if isinstance(v, (pd.Timestamp,)) and pd.notna(v) else (str(v)[:19] if v and str(v) not in ("None","NaT","nan","") else "")
            gantt.append({
                "task_code":        str(_a.get("task_code", "")),
                "task_name":        str(_a.get("task_name", "")),
                "wbs_name":         str(_a.get("wbs_name", "")),
                "status":           str(_a.get("status", "")),
                "planned_start":    _s(_a.get("planned_start")),
                "planned_finish":   _s(_a.get("planned_finish")),
                "actual_start":     _s(_a.get("actual_start")),
                "actual_finish":    _s(_a.get("actual_finish")),
                "phys_complete_pct": round(float(_a.get("phys_complete_pct", 0)), 1),
                "original_duration": round(float(_a.get("original_duration", 0)), 1),
                "total_float_days":  round(_float_map.get(_tid, 0), 1),
                "is_critical":       _critical_map.get(_tid, False),
                "wbs_id":            str(_a.get("wbs_id", "")),
                "task_type":         _task_type_map.get(_tid, ""),
            })
        # Sort by WBS then planned_start for sensible gantt ordering
        gantt.sort(key=lambda x: (x["wbs_name"], x["planned_start"] or ""))

    # Chart data (for schedule tab)
    chart_data  = prepare_chart_data(result)

    # AI summary (cached; stub if no key)
    ai_summary  = get_ai_summary(result, ck, _OPENROUTER_API_KEY, _OPENROUTER_MODEL)

    # Variance rows
    variances = [
        {
            "task_code":              v.task_code,
            "task_name":              v.task_name,
            "wbs_name":               v.wbs_name,
            "change_type":            v.change_type,
            "start_variance_days":    round(v.start_variance_days, 1),
            "finish_variance_days":   round(v.finish_variance_days, 1),
            "duration_variance_days": round(v.duration_variance_days, 1),
            "pct_complete_change":    round(v.pct_complete_change, 1),
            "float_change_hours":     round(v.float_change_hours, 1),
            "old_status":             v.old_status,
            "new_status":             v.new_status,
            "is_critical":            v.is_critical,
            "baseline_finish":        v.baseline_finish,
            "updated_finish":         v.updated_finish,
        }
        for v in result.activity_variances
    ]

    rel_variances = [
        {
            "pred_code":        r.pred_code,
            "succ_code":        r.succ_code,
            "change_type":      r.change_type,
            "old_pred_type":    r.old_pred_type,
            "new_pred_type":    r.new_pred_type,
            "lag_change_hours": round(r.lag_change_hours, 1),
        }
        for r in result.relationship_variances
    ]

    return {
        "key":          key16,
        "generated_at": datetime.utcnow().isoformat(),
        # Project
        "project": {
            "baseline_name":       result.baseline_project_name,
            "updated_name":        result.updated_project_name,
            "baseline_data_date":  str(b_date) if b_date else "",
            "updated_data_date":   str(data_date) if data_date else "",
            "data_date_warning":   result.data_date_warning,
            "plan_start":          str(result.plan_start) if result.plan_start else "",
            "plan_end":            str(result.plan_end)   if result.plan_end   else "",
        },
        # EV
        "ev": {
            "has_cost_data": has_cost_data,
            "baseline":      b_pi,
            "updated":       u_pi,
        },
        "scurve": {
            "baseline": scurve_b,
            "updated":  scurve_u,
        },
        # Comparison
        "summary":              result.summary,
        "chart_data":           chart_data,
        "activity_variances":   variances,
        "relationship_variances": rel_variances,
        # Supporting data
        "wbs_summary":          wbs_summary,
        "kpis":                 kpis,
        "lookahead":            lookahead,
        "milestones":           milestones,
        "procurement":          procurement,
        "oos":              oos,
        "bottleneck":       bottleneck,
        "all_relationships": all_relationships,
        "resource_loading":     resource_loading,
        "gantt":                gantt,
        # AI
        "ai_summary":           ai_summary,
    }


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", error=None)


@app.route("/compare", methods=["POST"])
def compare():
    def _err(msg, status=400):
        return jsonify({"error": msg}), status

    baseline_file = request.files.get("baseline")
    updated_file  = request.files.get("updated")

    if not baseline_file or not baseline_file.filename:
        return _err("Please upload the baseline XER file.")
    if not updated_file or not updated_file.filename:
        return _err("Please upload the updated XER file.")
    if not _allowed(baseline_file.filename):
        return _err("Baseline file must be a .xer file.")
    if not _allowed(updated_file.filename):
        return _err("Updated file must be a .xer file.")

    baseline_bytes = baseline_file.read()
    updated_bytes  = updated_file.read()

    if not baseline_bytes:
        return _err("Baseline file is empty.")
    if not updated_bytes:
        return _err("Updated file is empty.")

    try:
        ck    = cache_key(baseline_bytes, updated_bytes)
        key16 = ck[:16]
    except Exception as exc:
        return _err(f"Failed to hash XER files: {exc}")

    dest = url_for("dashboard", key=key16)

    # If already cached, return immediately
    if _full_path(key16).exists():
        return jsonify({"redirect": dest})

    # Start background job
    job_id = uuid.uuid4().hex
    with _jobs_lock:
        _jobs[job_id] = {"status": "processing", "pct": 0, "step": "Starting…"}

    t = threading.Thread(
        target=_run_job,
        args=(job_id, baseline_bytes, updated_bytes, ck, key16, dest),
        daemon=True,
    )
    t.start()

    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def job_status(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if job is None:
        return jsonify({"error": "Job not found"}), 404
    return jsonify(job)


@app.route("/r/<key>")
def dashboard(key: str):
    data = _load_full(key)
    if not data:
        return render_template("index.html",
            error="Dashboard not found. Please re-upload your files.")
    return render_template("dashboard.html", data=data, key=key)


@app.route("/export/csv")
def export_csv():
    key  = request.args.get("key", "")
    data = _load_full(key)
    if not data:
        return Response("Not found — please re-upload.", status=404)

    out = io.StringIO()
    fields = [
        "task_code", "task_name", "wbs_name", "change_type",
        "start_variance_days", "finish_variance_days", "duration_variance_days",
        "pct_complete_change", "float_change_hours",
        "old_status", "new_status", "is_critical",
        "baseline_finish", "updated_finish",
    ]
    w = csv.DictWriter(out, fieldnames=fields, extrasaction="ignore")
    w.writeheader()
    w.writerows(data.get("activity_variances", []))
    return Response(out.getvalue(), mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=schedule_comparison.csv"})


@app.route("/export/json")
def export_json():
    key  = request.args.get("key", "")
    data = _load_full(key)
    if not data:
        return Response("Not found — please re-upload.", status=404)
    return Response(
        json.dumps({"activity_variances": data.get("activity_variances", []),
                    "summary": data.get("summary", {})},
                   indent=2, ensure_ascii=False),
        mimetype="application/json",
        headers={"Content-Disposition": "attachment; filename=schedule_comparison.json"},
    )


@app.route("/chat", methods=["POST"])
def chat():
    body  = request.get_json(silent=True) or {}
    key   = body.get("key", "")
    message = body.get("message", "").strip()

    if not message:
        return Response("Missing message", status=400)

    data = _load_full(key)
    if not data:
        return Response("Dashboard not found", status=404)

    if not _OPENROUTER_API_KEY:
        def _stub():
            yield "data: AI chat requires OPENROUTER_API_KEY to be set.\n\n"
            yield "data: [DONE]\n\n"
        return Response(stream_with_context(_stub()), mimetype="text/event-stream")

    def _generate():
        try:
            for chunk in stream_chat_response(message, data, _OPENROUTER_API_KEY, _OPENROUTER_MODEL):
                yield f"data: {json.dumps(chunk)}\n\n"
        except Exception as exc:
            yield f"data: {json.dumps({'error': str(exc)})}\n\n"
        yield "data: [DONE]\n\n"

    return Response(stream_with_context(_generate()), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.errorhandler(413)
def too_large(_):
    return render_template("index.html",
        error=f"File too large. Maximum {_UPLOAD_MAX_MB} MB per file."), 413


if __name__ == "__main__":
    port  = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
