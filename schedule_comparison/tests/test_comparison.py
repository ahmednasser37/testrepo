"""Tests for comparison_engine.py."""

import pandas as pd
import pytest

from comparison_engine import (
    compare_schedules,
    ComparisonResult,
    match_activities,
)
from xer_parser import parse_xer_string


# ── Inline minimal XER for isolated tests ────────────────────────────────────

def _make_xer(activities: list[dict], data_date: str = "2025-06-01 00:00") -> str:
    """Build a minimal XER string from a list of activity dicts."""
    header = (
        f"ERMHDR\t19.12\t2026-01-01\n"
        f"%T\tPROJECT\n"
        f"%F\tproj_id\tproj_short_name\tproj_name\tplan_start_date\tplan_end_date\tdata_date\n"
        f"%R\tP001\tTEST\tTest Project\t2024-01-01 00:00\t2025-12-31 00:00\t{data_date}\n"
        f"%T\tWBS\n"
        f"%F\twbs_id\tproj_id\twbs_name\tparent_wbs_id\tproj_node_flag\n"
        f"%R\tW-001\tP001\tTest WBS\t\tY\n"
    )
    task_fields = (
        "task_id\tproj_id\twbs_id\ttask_code\ttask_name\tstatus_code"
        "\ttarget_start_date\ttarget_end_date"
        "\ttarget_drtn_hr_cnt\tremain_drtn_hr_cnt\tphys_complete_pct\ttotal_float_hr_cnt"
    )
    rows = [f"%T\tTASK", f"%F\t{task_fields}"]
    for act in activities:
        rows.append(
            "%R\t"
            + "\t".join([
                act.get("task_id", "T-001"),
                "P001",
                "W-001",
                act.get("task_code", "ACT-001"),
                act.get("task_name", "Activity"),
                act.get("status_code", "TK_NotStart"),
                act.get("target_start_date", "2025-01-01 00:00"),
                act.get("target_end_date", "2025-03-31 00:00"),
                act.get("target_drtn_hr_cnt", "480"),
                act.get("remain_drtn_hr_cnt", "480"),
                act.get("phys_complete_pct", "0.0"),
                act.get("total_float_hr_cnt", "8"),
            ])
        )
    return header + "\n".join(rows) + "\n%E\n"


_ACT_A = {
    "task_id": "T-001", "task_code": "ACT-001", "task_name": "Excavation",
    "status_code": "TK_NotStart",
    "target_start_date": "2025-01-01 00:00", "target_end_date": "2025-03-31 00:00",
    "target_drtn_hr_cnt": "480", "remain_drtn_hr_cnt": "480",
    "phys_complete_pct": "0.0", "total_float_hr_cnt": "8",
}
_ACT_B = {
    "task_id": "T-002", "task_code": "ACT-002", "task_name": "Pipe Laying",
    "status_code": "TK_NotStart",
    "target_start_date": "2025-04-01 00:00", "target_end_date": "2025-06-30 00:00",
    "target_drtn_hr_cnt": "480", "remain_drtn_hr_cnt": "480",
    "phys_complete_pct": "0.0", "total_float_hr_cnt": "4",
}


# ── Tests ─────────────────────────────────────────────────────────────────────

class TestCompareReturnsResult:
    def test_returns_comparison_result(self, baseline_tables, updated_tables):
        result = compare_schedules(baseline_tables, updated_tables)
        assert isinstance(result, ComparisonResult)

    def test_has_activity_variances(self, baseline_tables, updated_tables):
        result = compare_schedules(baseline_tables, updated_tables)
        assert isinstance(result.activity_variances, list)
        assert len(result.activity_variances) > 0

    def test_has_summary_dict(self, baseline_tables, updated_tables):
        result = compare_schedules(baseline_tables, updated_tables)
        assert isinstance(result.summary, dict)
        for key in ("total_baseline", "total_updated", "added", "deleted", "changed", "unchanged"):
            assert key in result.summary

    def test_project_names_populated(self, baseline_tables):
        result = compare_schedules(baseline_tables, baseline_tables)
        assert isinstance(result.baseline_project_name, str)
        assert isinstance(result.updated_project_name, str)


class TestIdenticalSchedules:
    def test_no_variances_against_self(self, baseline_tables):
        result = compare_schedules(baseline_tables, baseline_tables)
        changed = [v for v in result.activity_variances if v.change_type == "changed"]
        assert len(changed) == 0

    def test_all_unchanged_against_self(self, baseline_tables):
        result = compare_schedules(baseline_tables, baseline_tables)
        unchanged = [v for v in result.activity_variances if v.change_type == "unchanged"]
        assert len(unchanged) == result.summary["total_baseline"]

    def test_zero_variances_against_self(self, baseline_tables):
        result = compare_schedules(baseline_tables, baseline_tables)
        for v in result.activity_variances:
            assert v.finish_variance_days == 0.0
            assert v.start_variance_days == 0.0

    def test_no_data_date_warning_same_file(self, baseline_tables):
        result = compare_schedules(baseline_tables, baseline_tables)
        assert result.data_date_warning is False


class TestSummaryCounts:
    def test_added_count_matches_list(self, baseline_tables, updated_tables):
        result = compare_schedules(baseline_tables, updated_tables)
        actual_added = sum(1 for v in result.activity_variances if v.change_type == "added")
        assert result.summary["added"] == actual_added

    def test_deleted_count_matches_list(self, baseline_tables, updated_tables):
        result = compare_schedules(baseline_tables, updated_tables)
        actual_deleted = sum(1 for v in result.activity_variances if v.change_type == "deleted")
        assert result.summary["deleted"] == actual_deleted

    def test_changed_count_matches_list(self, baseline_tables, updated_tables):
        result = compare_schedules(baseline_tables, updated_tables)
        actual_changed = sum(1 for v in result.activity_variances if v.change_type == "changed")
        assert result.summary["changed"] == actual_changed

    def test_unchanged_count_matches_list(self, baseline_tables, updated_tables):
        result = compare_schedules(baseline_tables, updated_tables)
        actual_unchanged = sum(1 for v in result.activity_variances if v.change_type == "unchanged")
        assert result.summary["unchanged"] == actual_unchanged

    def test_total_baseline_is_baseline_activity_count(self, baseline_tables, updated_tables):
        result = compare_schedules(baseline_tables, updated_tables)
        assert result.summary["total_baseline"] == len(baseline_tables["TASK"])


class TestChangedDates:
    def test_shifted_finish_shows_positive_variance(self, baseline_tables, updated_tables):
        """updated_tables has 3 activities with finish date +30 days."""
        result = compare_schedules(baseline_tables, updated_tables)
        delayed = [v for v in result.activity_variances
                   if v.change_type == "changed" and v.finish_variance_days > 0]
        assert len(delayed) >= 1

    def test_variance_approximately_30_days(self, baseline_tables, updated_tables):
        result = compare_schedules(baseline_tables, updated_tables)
        large_delays = [v for v in result.activity_variances
                        if v.finish_variance_days >= 29.0]
        assert len(large_delays) >= 1

    def test_max_delay_days_positive(self, baseline_tables, updated_tables):
        result = compare_schedules(baseline_tables, updated_tables)
        assert result.summary["max_delay_days"] >= 29.0


class TestAddedDeletedActivities:
    def test_added_activity_detected(self):
        base_xer = _make_xer([_ACT_A])
        upd_xer = _make_xer([_ACT_A, _ACT_B])
        base_t = parse_xer_string(base_xer)
        upd_t = parse_xer_string(upd_xer)
        result = compare_schedules(base_t, upd_t)
        added = [v for v in result.activity_variances if v.change_type == "added"]
        assert any(v.task_code == "ACT-002" for v in added)

    def test_deleted_activity_detected(self):
        base_xer = _make_xer([_ACT_A, _ACT_B])
        upd_xer = _make_xer([_ACT_A])
        base_t = parse_xer_string(base_xer)
        upd_t = parse_xer_string(upd_xer)
        result = compare_schedules(base_t, upd_t)
        deleted = [v for v in result.activity_variances if v.change_type == "deleted"]
        assert any(v.task_code == "ACT-002" for v in deleted)

    def test_added_count_in_summary(self):
        base_xer = _make_xer([_ACT_A])
        upd_xer = _make_xer([_ACT_A, _ACT_B])
        result = compare_schedules(parse_xer_string(base_xer), parse_xer_string(upd_xer))
        assert result.summary["added"] == 1
        assert result.summary["deleted"] == 0


class TestDataDateWarning:
    def test_same_data_date_no_warning(self):
        xer = _make_xer([_ACT_A], data_date="2025-06-01 00:00")
        tables = parse_xer_string(xer)
        result = compare_schedules(tables, tables)
        assert result.data_date_warning is False

    def test_different_data_dates_triggers_warning(self):
        base_xer = _make_xer([_ACT_A], data_date="2025-03-01 00:00")
        upd_xer = _make_xer([_ACT_A], data_date="2025-06-01 00:00")
        result = compare_schedules(
            parse_xer_string(base_xer),
            parse_xer_string(upd_xer),
        )
        assert result.data_date_warning is True


class TestMatchActivities:
    def test_exact_code_match(self):
        base = pd.DataFrame({"task_code": ["A", "B", "C"]})
        updated = pd.DataFrame({"task_code": ["A", "B", "D"]})
        m = match_activities(base, updated)
        assert m.get("A") == "A"
        assert m.get("B") == "B"
        assert "C" not in m

    def test_empty_base_returns_empty(self):
        result = match_activities(pd.DataFrame(), pd.DataFrame({"task_code": ["A"]}))
        assert result == {}

    def test_empty_updated_returns_empty(self):
        result = match_activities(pd.DataFrame({"task_code": ["A"]}), pd.DataFrame())
        assert result == {}
