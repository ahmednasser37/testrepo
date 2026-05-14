"""
AI summary service — wraps OpenRouter API with local file caching.
Sends only compact structured context; never sends raw XER content.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Iterator

import requests

from comparison_engine import ComparisonResult

_CACHE_DIR = Path(os.environ.get("SCE_CACHE_DIR", "/tmp/sce_cache"))
_DEFAULT_MODEL = "deepseek/deepseek-chat-v3-0324:free"
_OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

_STUB: dict = {
    "executive_summary": "AI summary unavailable — set OPENROUTER_API_KEY to enable.",
    "key_risks": [],
    "main_variances": [],
    "schedule_health": "unknown",
    "recommended_actions": [],
}


# ── Cache helpers ─────────────────────────────────────────────────────────────

def cache_key(baseline_content: bytes, updated_content: bytes) -> str:
    """Return a SHA256 hex digest of both file contents concatenated."""
    h = hashlib.sha256()
    h.update(baseline_content)
    h.update(b"\x00")
    h.update(updated_content)
    return h.hexdigest()


def _cache_path(key: str) -> Path:
    _CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return _CACHE_DIR / f"{key[:16]}.json"


def load_cached(key: str) -> dict | None:
    p = _cache_path(key)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None


def save_cached(key: str, data: dict) -> None:
    try:
        _cache_path(key).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


# ── Prompt builder ────────────────────────────────────────────────────────────

def build_prompt(result: ComparisonResult) -> str:
    s = result.summary

    def _fmt(dt) -> str:
        if dt is None:
            return "N/A"
        try:
            return dt.strftime("%Y-%m-%d")
        except Exception:
            return str(dt)

    context = {
        "baseline_project": result.baseline_project_name,
        "updated_project": result.updated_project_name,
        "baseline_data_date": _fmt(result.baseline_data_date),
        "updated_data_date": _fmt(result.updated_data_date),
        "data_date_warning": result.data_date_warning,
        "summary": {
            "total_baseline_activities": s.get("total_baseline", 0),
            "total_updated_activities": s.get("total_updated", 0),
            "added": s.get("added", 0),
            "deleted": s.get("deleted", 0),
            "changed": s.get("changed", 0),
            "unchanged": s.get("unchanged", 0),
            "delayed_activities": s.get("delayed_activities", 0),
            "improved_activities": s.get("improved_activities", 0),
            "critical_activities": s.get("critical_activities_updated", 0),
            "max_delay_days": s.get("max_delay_days", 0),
            "avg_finish_variance_days": s.get("avg_finish_variance_days", 0),
            "relationships_added": s.get("relationships_added", 0),
            "relationships_deleted": s.get("relationships_deleted", 0),
            "relationships_changed": s.get("relationships_changed", 0),
        },
        "top_delayed_activities": s.get("top_delayed", [])[:10],
    }

    return (
        "You are a senior Primavera P6 schedule analyst. "
        "Analyze this schedule comparison and return ONLY valid JSON — no markdown, no code blocks.\n\n"
        f"SCHEDULE COMPARISON DATA:\n{json.dumps(context, indent=2)}\n\n"
        "Return exactly this JSON structure:\n"
        "{\n"
        '  "executive_summary": "<2-3 sentence overview of schedule health and key changes>",\n'
        '  "key_risks": ["<risk 1>", "<risk 2>", ...],\n'
        '  "main_variances": ["<variance description 1>", ...],\n'
        '  "schedule_health": "<one of: on_track | at_risk | critical>",\n'
        '  "recommended_actions": ["<action 1>", "<action 2>", ...]\n'
        "}"
    )


# ── OpenRouter call ───────────────────────────────────────────────────────────

def call_openrouter(prompt: str, api_key: str, model: str = _DEFAULT_MODEL) -> dict:
    """Call the OpenRouter chat completions endpoint and return parsed JSON response."""
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.2,
        "max_tokens": 800,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://github.com/ahmednasser37/testrepo",
        "X-Title": "P6 Schedule Comparison",
    }
    resp = requests.post(_OPENROUTER_URL, json=payload, headers=headers, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    content = data["choices"][0]["message"]["content"].strip()

    # Strip markdown code fences if present
    if content.startswith("```"):
        lines = content.splitlines()
        content = "\n".join(
            line for line in lines
            if not line.strip().startswith("```")
        ).strip()

    return json.loads(content)


# ── Public entry ──────────────────────────────────────────────────────────────

def get_ai_summary(
    result: ComparisonResult,
    ck: str,
    api_key: str | None = None,
    model: str = _DEFAULT_MODEL,
) -> dict:
    """
    Return an AI-written schedule summary dict.
    Uses cache if available. Falls back to stub if no API key is set.
    """
    cached = load_cached(ck)
    if cached is not None:
        return cached

    if not api_key:
        return dict(_STUB)

    try:
        prompt = build_prompt(result)
        summary = call_openrouter(prompt, api_key, model)
        # Validate expected keys exist
        for key in ("executive_summary", "key_risks", "schedule_health"):
            if key not in summary:
                summary[key] = _STUB[key]
        save_cached(ck, summary)
        return summary
    except Exception as exc:
        stub = dict(_STUB)
        stub["executive_summary"] = f"AI summary failed: {exc}"
        return stub


# ── Chat streaming ────────────────────────────────────────────────────────────

def _build_chat_context(full_data: dict) -> str:
    """Build a compact context string from the full dashboard data dict."""
    proj = full_data.get("project", {})
    summary = full_data.get("summary", {})
    kpis = full_data.get("kpis", {})
    milestones = full_data.get("milestones", [])
    procurement = full_data.get("procurement", {})
    ai_sum = full_data.get("ai_summary", {})

    ms_rows = [
        {"name": m.get("task_name"), "status": m.get("status"), "variance_days": m.get("finish_variance_days")}
        for m in milestones[:10]
    ]
    proc_items = [
        {"name": p.get("task_name"), "pct": p.get("pct_complete"), "status": p.get("status")}
        for p in procurement.get("items", [])[:10]
    ]

    ctx = {
        "project": {
            "baseline": proj.get("baseline_name"),
            "updated": proj.get("updated_name"),
            "data_date": proj.get("updated_data_date"),
            "plan_end": proj.get("plan_end"),
        },
        "schedule_health": ai_sum.get("schedule_health", "unknown"),
        "summary": {k: summary.get(k) for k in (
            "total_baseline", "total_updated", "added", "deleted", "changed",
            "delayed_activities", "max_delay_days", "avg_finish_variance_days",
        )},
        "kpis": {k: kpis.get(k) for k in (
            "float_consumption_days", "pct_complete_weighted", "spi_duration",
            "schedule_delay_days", "critical_total", "near_critical_count",
        )},
        "milestones": ms_rows,
        "procurement": {
            "total": procurement.get("summary", {}).get("total"),
            "late": procurement.get("summary", {}).get("late"),
            "items": proc_items,
        },
        "top_delayed": summary.get("top_delayed", [])[:10],
        "ai_assessment": ai_sum.get("executive_summary", ""),
    }
    return json.dumps(ctx, default=str)


def stream_chat_response(
    message: str,
    full_data: dict,
    api_key: str,
    model: str = _DEFAULT_MODEL,
) -> Iterator[dict]:
    """
    Stream a chat response from OpenRouter given a user message and full dashboard data.
    Yields dicts: {"content": str} for text chunks, {"error": str} on failure.
    """
    context = _build_chat_context(full_data)
    system_prompt = (
        "You are an expert Primavera P6 schedule analyst and project controls engineer. "
        "Answer questions about the schedule comparison data provided. "
        "Be concise, data-driven, and focus on actionable insights. "
        "Use the data context below for all answers.\n\n"
        f"SCHEDULE DATA:\n{context}"
    )

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": message},
        ],
        "temperature": 0.3,
        "max_tokens": 600,
        "stream": True,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://github.com/ahmednasser37/testrepo",
        "X-Title": "P6 Schedule Comparison",
    }

    try:
        with requests.post(
            _OPENROUTER_URL, json=payload, headers=headers, stream=True, timeout=60
        ) as resp:
            resp.raise_for_status()
            for raw_line in resp.iter_lines():
                if not raw_line:
                    continue
                line = raw_line.decode("utf-8") if isinstance(raw_line, bytes) else raw_line
                if not line.startswith("data: "):
                    continue
                data_str = line[6:].strip()
                if data_str == "[DONE]":
                    break
                try:
                    chunk = json.loads(data_str)
                    delta = chunk.get("choices", [{}])[0].get("delta", {})
                    text = delta.get("content", "")
                    if text:
                        yield {"content": text}
                except (json.JSONDecodeError, KeyError, IndexError):
                    continue
    except Exception as exc:
        yield {"error": str(exc)}
