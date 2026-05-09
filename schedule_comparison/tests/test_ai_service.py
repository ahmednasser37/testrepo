"""Tests for ai_service.py."""

import json
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from ai_service import (
    cache_key,
    load_cached,
    save_cached,
    build_prompt,
    call_openrouter,
    get_ai_summary,
    _STUB,
)
from comparison_engine import ComparisonResult


# ── Helpers ───────────────────────────────────────────────────────────────────

def _make_minimal_result() -> ComparisonResult:
    """Return a minimal ComparisonResult suitable for prompt building."""
    return ComparisonResult(
        baseline_project_name="Project A",
        updated_project_name="Project A",
        baseline_data_date=None,
        updated_data_date=None,
        plan_start=None,
        plan_end=None,
        data_date_warning=False,
        activity_variances=[],
        relationship_variances=[],
        summary={
            "total_baseline": 10,
            "total_updated": 10,
            "added": 0,
            "deleted": 0,
            "changed": 2,
            "unchanged": 8,
            "delayed_activities": 1,
            "improved_activities": 0,
            "critical_activities_updated": 1,
            "max_delay_days": 14.0,
            "avg_finish_variance_days": 7.0,
            "relationships_added": 0,
            "relationships_deleted": 0,
            "relationships_changed": 0,
            "top_delayed": [],
        },
    )


# ── cache_key ─────────────────────────────────────────────────────────────────

class TestCacheKey:
    def test_returns_string(self):
        k = cache_key(b"abc", b"def")
        assert isinstance(k, str)

    def test_deterministic(self):
        k1 = cache_key(b"baseline", b"updated")
        k2 = cache_key(b"baseline", b"updated")
        assert k1 == k2

    def test_different_inputs_differ(self):
        k1 = cache_key(b"a", b"b")
        k2 = cache_key(b"c", b"d")
        assert k1 != k2

    def test_order_matters(self):
        k1 = cache_key(b"x", b"y")
        k2 = cache_key(b"y", b"x")
        assert k1 != k2

    def test_length_is_64_hex_chars(self):
        k = cache_key(b"foo", b"bar")
        assert len(k) == 64
        assert all(c in "0123456789abcdef" for c in k)


# ── Cache load/save ───────────────────────────────────────────────────────────

class TestCacheLoadSave:
    def test_save_and_load_roundtrip(self, tmp_path, monkeypatch):
        monkeypatch.setenv("SCE_CACHE_DIR", str(tmp_path))
        import ai_service
        monkeypatch.setattr(ai_service, "_CACHE_DIR", tmp_path)

        key = "a" * 64
        data = {"executive_summary": "test"}
        save_cached(key, data)
        loaded = load_cached(key)
        assert loaded == data

    def test_load_missing_key_returns_none(self, tmp_path, monkeypatch):
        import ai_service
        monkeypatch.setattr(ai_service, "_CACHE_DIR", tmp_path)
        result = load_cached("no_such_key" + "0" * 50)
        assert result is None


# ── No API key stub ───────────────────────────────────────────────────────────

class TestNoApiKeyReturnsStub:
    def test_returns_dict_with_executive_summary(self):
        result = _make_minimal_result()
        summary = get_ai_summary(result, "somekey", None, "some-model")
        assert isinstance(summary, dict)
        assert "executive_summary" in summary

    def test_executive_summary_mentions_api_key(self):
        result = _make_minimal_result()
        summary = get_ai_summary(result, "somekey", None, "some-model")
        assert "OPENROUTER_API_KEY" in summary["executive_summary"]

    def test_schedule_health_unknown(self):
        result = _make_minimal_result()
        summary = get_ai_summary(result, "somekey", None, "some-model")
        assert summary["schedule_health"] == "unknown"

    def test_stub_has_all_required_keys(self):
        result = _make_minimal_result()
        summary = get_ai_summary(result, "somekey", None, "some-model")
        for k in ("executive_summary", "key_risks", "main_variances",
                  "schedule_health", "recommended_actions"):
            assert k in summary


# ── Cache hit ─────────────────────────────────────────────────────────────────

class TestCacheHit:
    def test_second_call_uses_cache(self, tmp_path, monkeypatch):
        """get_ai_summary should not call call_openrouter if cache has data."""
        import ai_service
        monkeypatch.setattr(ai_service, "_CACHE_DIR", tmp_path)

        ck = "b" * 64
        cached_data = {
            "executive_summary": "cached result",
            "key_risks": [],
            "main_variances": [],
            "schedule_health": "on_track",
            "recommended_actions": [],
        }
        save_cached(ck, cached_data)

        mock_call = MagicMock()
        monkeypatch.setattr(ai_service, "call_openrouter", mock_call)

        result = _make_minimal_result()
        summary = get_ai_summary(result, ck, "fake-api-key", "some-model")

        mock_call.assert_not_called()
        assert summary["executive_summary"] == "cached result"

    def test_cache_result_returned_verbatim(self, tmp_path, monkeypatch):
        import ai_service
        monkeypatch.setattr(ai_service, "_CACHE_DIR", tmp_path)

        ck = "c" * 64
        cached_data = {"executive_summary": "verbatim", "schedule_health": "critical",
                       "key_risks": ["risk 1"], "main_variances": [], "recommended_actions": []}
        save_cached(ck, cached_data)

        result = _make_minimal_result()
        summary = get_ai_summary(result, ck, None, "model")
        assert summary == cached_data


# ── build_prompt ──────────────────────────────────────────────────────────────

class TestBuildPrompt:
    def test_returns_string(self):
        prompt = build_prompt(_make_minimal_result())
        assert isinstance(prompt, str)
        assert len(prompt) > 100

    def test_contains_project_name(self):
        prompt = build_prompt(_make_minimal_result())
        assert "Project A" in prompt

    def test_instructs_json_output(self):
        prompt = build_prompt(_make_minimal_result())
        assert "JSON" in prompt


# ── call_openrouter (mocked) ──────────────────────────────────────────────────

class TestCallOpenrouterMocked:
    def _mock_response(self, content: dict) -> MagicMock:
        resp = MagicMock()
        resp.raise_for_status = MagicMock()
        resp.json.return_value = {
            "choices": [{"message": {"content": json.dumps(content)}}]
        }
        return resp

    def test_returns_parsed_dict(self):
        expected = {
            "executive_summary": "On track.",
            "key_risks": [],
            "main_variances": [],
            "schedule_health": "on_track",
            "recommended_actions": [],
        }
        with patch("ai_service.requests.post") as mock_post:
            mock_post.return_value = self._mock_response(expected)
            result = call_openrouter("test prompt", "fake-key", "some-model")
        assert result["schedule_health"] == "on_track"
        assert result["executive_summary"] == "On track."

    def test_strips_markdown_fences(self):
        content = {
            "executive_summary": "Great.",
            "key_risks": [],
            "main_variances": [],
            "schedule_health": "on_track",
            "recommended_actions": [],
        }
        fenced = f"```json\n{json.dumps(content)}\n```"
        resp = MagicMock()
        resp.raise_for_status = MagicMock()
        resp.json.return_value = {"choices": [{"message": {"content": fenced}}]}

        with patch("ai_service.requests.post", return_value=resp):
            result = call_openrouter("prompt", "key", "model")
        assert result["schedule_health"] == "on_track"

    def test_sends_api_key_in_header(self):
        content = {"executive_summary": "x", "key_risks": [], "main_variances": [],
                   "schedule_health": "on_track", "recommended_actions": []}
        with patch("ai_service.requests.post") as mock_post:
            mock_post.return_value = self._mock_response(content)
            call_openrouter("prompt", "MY-SECRET-KEY", "model")
        _, kwargs = mock_post.call_args
        headers = kwargs.get("headers", {})
        assert "MY-SECRET-KEY" in headers.get("Authorization", "")
