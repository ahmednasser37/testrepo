"""
Flask web application — upload two XER files, compare them, render dashboard.
"""

from __future__ import annotations

import csv
import io
import json
import os
import tempfile
from pathlib import Path

from flask import Flask, request, render_template, Response, redirect, url_for

from xer_parser import parse_xer_string
from comparison_engine import compare_schedules, ComparisonResult
from ai_service import cache_key, get_ai_summary
from dashboard_renderer import render_dashboard

app = Flask(__name__)
app.secret_key = os.environ.get("SCE_SECRET_KEY", "dev-secret-change-in-production")

_UPLOAD_MAX_MB = int(os.environ.get("SCE_UPLOAD_MAX_MB", 50))
app.config["MAX_CONTENT_LENGTH"] = _UPLOAD_MAX_MB * 1024 * 1024

_CACHE_DIR = Path(os.environ.get("SCE_CACHE_DIR", "/tmp/sce_cache"))
_OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY")
_OPENROUTER_MODEL = os.environ.get("OPENROUTER_MODEL", "deepseek/deepseek-chat-v3-0324:free")

_ALLOWED_EXTENSIONS = {".xer"}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _allowed(filename: str) -> bool:
    return Path(filename).suffix.lower() in _ALLOWED_EXTENSIONS


def _save_result(ck: str, result: ComparisonResult) -> None:
    """Persist comparison result as JSON for export endpoints."""
    _CACHE_DIR.mkdir(parents=True, exist_ok=True)
    data = _result_to_dict(result)
    (_CACHE_DIR / f"{ck[:16]}_result.json").write_text(
        json.dumps(data, ensure_ascii=False, default=str), encoding="utf-8"
    )


def _load_result_dict(ck: str) -> dict | None:
    p = _CACHE_DIR / f"{ck[:16]}_result.json"
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


def _result_to_dict(result: ComparisonResult) -> dict:
    return {
        "baseline_project":  result.baseline_project_name,
        "updated_project":   result.updated_project_name,
        "baseline_data_date": str(result.baseline_data_date),
        "updated_data_date":  str(result.updated_data_date),
        "data_date_warning":  result.data_date_warning,
        "summary": result.summary,
        "activity_variances": [
            {
                "task_code":              av.task_code,
                "task_name":              av.task_name,
                "wbs_name":               av.wbs_name,
                "change_type":            av.change_type,
                "start_variance_days":    av.start_variance_days,
                "finish_variance_days":   av.finish_variance_days,
                "duration_variance_days": av.duration_variance_days,
                "pct_complete_change":    av.pct_complete_change,
                "float_change_hours":     av.float_change_hours,
                "old_status":             av.old_status,
                "new_status":             av.new_status,
                "is_critical":            av.is_critical,
                "baseline_finish":        av.baseline_finish,
                "updated_finish":         av.updated_finish,
            }
            for av in result.activity_variances
        ],
        "relationship_variances": [
            {
                "pred_code":       rv.pred_code,
                "succ_code":       rv.succ_code,
                "change_type":     rv.change_type,
                "old_pred_type":   rv.old_pred_type,
                "new_pred_type":   rv.new_pred_type,
                "lag_change_hours": rv.lag_change_hours,
            }
            for rv in result.relationship_variances
        ],
    }


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", error=None)


@app.route("/compare", methods=["POST"])
def compare():
    baseline_file = request.files.get("baseline")
    updated_file  = request.files.get("updated")

    # Validate uploads
    if not baseline_file or not baseline_file.filename:
        return render_template("index.html", error="Please upload the baseline XER file.")
    if not updated_file or not updated_file.filename:
        return render_template("index.html", error="Please upload the updated XER file.")

    if not _allowed(baseline_file.filename):
        return render_template("index.html", error="Baseline file must be a .xer file.")
    if not _allowed(updated_file.filename):
        return render_template("index.html", error="Updated file must be a .xer file.")

    baseline_bytes = baseline_file.read()
    updated_bytes  = updated_file.read()

    if not baseline_bytes:
        return render_template("index.html", error="Baseline file is empty.")
    if not updated_bytes:
        return render_template("index.html", error="Updated file is empty.")

    try:
        baseline_content = baseline_bytes.decode("utf-8", errors="replace")
        updated_content  = updated_bytes.decode("utf-8", errors="replace")

        baseline_tables = parse_xer_string(baseline_content)
        updated_tables  = parse_xer_string(updated_content)

        if "PROJECT" not in baseline_tables or baseline_tables["PROJECT"].empty:
            return render_template("index.html", error="Baseline file does not contain a valid PROJECT table.")
        if "PROJECT" not in updated_tables or updated_tables["PROJECT"].empty:
            return render_template("index.html", error="Updated file does not contain a valid PROJECT table.")
        if "TASK" not in baseline_tables or baseline_tables["TASK"].empty:
            return render_template("index.html", error="Baseline file contains no activities (TASK table missing or empty).")
        if "TASK" not in updated_tables or updated_tables["TASK"].empty:
            return render_template("index.html", error="Updated file contains no activities (TASK table missing or empty).")

    except Exception as exc:
        return render_template("index.html", error=f"Failed to parse XER files: {exc}")

    try:
        ck = cache_key(baseline_bytes, updated_bytes)
        result = compare_schedules(baseline_tables, updated_tables)
        _save_result(ck, result)

        ai_summary = get_ai_summary(result, ck, _OPENROUTER_API_KEY, _OPENROUTER_MODEL)

        html = render_dashboard(result, ai_summary, ck)
        return Response(html, mimetype="text/html")

    except Exception as exc:
        return render_template("index.html", error=f"Comparison failed: {exc}")


@app.route("/export/csv")
def export_csv():
    ck = request.args.get("key", "")
    data = _load_result_dict(ck)
    if not data:
        return Response("Comparison result not found. Please re-upload the files.", status=404)

    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=[
        "task_code", "task_name", "wbs_name", "change_type",
        "start_variance_days", "finish_variance_days", "duration_variance_days",
        "pct_complete_change", "float_change_hours",
        "old_status", "new_status", "is_critical",
        "baseline_finish", "updated_finish",
    ])
    writer.writeheader()
    writer.writerows(data.get("activity_variances", []))

    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=schedule_comparison.csv"},
    )


@app.route("/export/json")
def export_json():
    ck = request.args.get("key", "")
    data = _load_result_dict(ck)
    if not data:
        return Response("Comparison result not found. Please re-upload the files.", status=404)

    return Response(
        json.dumps(data, indent=2, ensure_ascii=False),
        mimetype="application/json",
        headers={"Content-Disposition": "attachment; filename=schedule_comparison.json"},
    )


@app.errorhandler(413)
def too_large(_):
    return render_template("index.html",
                           error=f"File too large. Maximum size is {_UPLOAD_MAX_MB} MB per file."), 413


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
