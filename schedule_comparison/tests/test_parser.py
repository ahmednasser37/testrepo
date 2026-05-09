"""Tests for xer_parser.py."""

import pandas as pd
import pytest

from xer_parser import parse_xer_string, extract_data_date, extract_project_meta, get_relationships


# ── Minimal inline XER fixtures ───────────────────────────────────────────────

_MINIMAL_XER = """\
ERMHDR\t19.12\t2026-01-01\tPrimavera\tadmin
%T\tPROJECT
%F\tproj_id\tproj_short_name\tproj_name\tplan_start_date\tplan_end_date\tdata_date
%R\tP001\tTEST\tTest Project\t2025-01-01 00:00\t2025-12-31 00:00\t2025-06-01 00:00
%T\tTASK
%F\ttask_id\tproj_id\twbs_id\ttask_code\ttask_name\tstatus_code\ttarget_start_date\ttarget_end_date\ttarget_drtn_hr_cnt\tremain_drtn_hr_cnt\tphys_complete_pct\ttotal_float_hr_cnt
%R\tT-001\tP001\tW-001\tACT-001\tFirst Activity\tTK_NotStart\t2025-01-01 00:00\t2025-03-31 00:00\t480\t480\t0.0\t8
%R\tT-002\tP001\tW-001\tACT-002\tSecond Activity\tTK_Active\t2025-04-01 00:00\t2025-06-30 00:00\t480\t240\t0.5\t0
%E
"""

_XER_WITH_TASKPRED = """\
ERMHDR\t19.12\t2026-01-01
%T\tPROJECT
%F\tproj_id\tproj_short_name\tproj_name\tplan_start_date\tplan_end_date\tdata_date
%R\tP001\tTEST\tTest Project\t2025-01-01 00:00\t2025-12-31 00:00\t2025-06-01 00:00
%T\tTASK
%F\ttask_id\ttask_code\ttask_name\tstatus_code\ttarget_start_date\ttarget_end_date\ttarget_drtn_hr_cnt\tremain_drtn_hr_cnt\tphys_complete_pct\ttotal_float_hr_cnt
%R\tT-001\tACT-001\tFirst Activity\tTK_NotStart\t2025-01-01 00:00\t2025-03-31 00:00\t480\t480\t0.0\t8
%R\tT-002\tACT-002\tSecond Activity\tTK_NotStart\t2025-04-01 00:00\t2025-06-30 00:00\t480\t480\t0.0\t4
%T\tTASKPRED
%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type\tlag_hr_cnt
%R\tTP-001\tT-002\tT-001\tFS\t0
%E
"""


# ── Tests ─────────────────────────────────────────────────────────────────────

class TestParseXerStringBasic:
    def test_project_table_present(self):
        tables = parse_xer_string(_MINIMAL_XER)
        assert "PROJECT" in tables
        assert not tables["PROJECT"].empty

    def test_task_table_present(self):
        tables = parse_xer_string(_MINIMAL_XER)
        assert "TASK" in tables
        assert not tables["TASK"].empty
        assert len(tables["TASK"]) == 2

    def test_task_codes_preserved(self):
        tables = parse_xer_string(_MINIMAL_XER)
        codes = set(tables["TASK"]["task_code"])
        assert "ACT-001" in codes
        assert "ACT-002" in codes

    def test_sample_xer_project_table(self, baseline_tables):
        assert "PROJECT" in baseline_tables
        assert len(baseline_tables["PROJECT"]) == 1

    def test_sample_xer_task_count(self, baseline_tables):
        assert "TASK" in baseline_tables
        assert len(baseline_tables["TASK"]) > 50


class TestDataDateExtraction:
    def test_returns_timestamp(self):
        tables = parse_xer_string(_MINIMAL_XER)
        dd = extract_data_date(tables)
        assert isinstance(dd, pd.Timestamp)

    def test_correct_date_value(self):
        tables = parse_xer_string(_MINIMAL_XER)
        dd = extract_data_date(tables)
        assert dd.year == 2025
        assert dd.month == 6
        assert dd.day == 1

    def test_empty_tables_returns_none(self):
        dd = extract_data_date({})
        assert dd is None

    def test_sample_xer_has_data_date(self, baseline_tables):
        dd = extract_data_date(baseline_tables)
        assert dd is not None
        assert isinstance(dd, pd.Timestamp)


class TestTaskpredParsing:
    def test_taskpred_table_present(self):
        tables = parse_xer_string(_XER_WITH_TASKPRED)
        assert "TASKPRED" in tables
        assert len(tables["TASKPRED"]) == 1

    def test_get_relationships_columns(self):
        tables = parse_xer_string(_XER_WITH_TASKPRED)
        rels = get_relationships(tables)
        for col in ("task_pred_id", "task_id", "pred_task_id", "pred_type", "lag_hr_cnt"):
            assert col in rels.columns

    def test_get_relationships_row_count(self):
        tables = parse_xer_string(_XER_WITH_TASKPRED)
        rels = get_relationships(tables)
        assert len(rels) == 1

    def test_get_relationships_lag_is_numeric(self):
        tables = parse_xer_string(_XER_WITH_TASKPRED)
        rels = get_relationships(tables)
        assert pd.api.types.is_numeric_dtype(rels["lag_hr_cnt"])

    def test_no_taskpred_returns_empty_df(self):
        tables = parse_xer_string(_MINIMAL_XER)
        rels = get_relationships(tables)
        assert rels.empty


class TestEmptyAndMalformedXer:
    def test_empty_string_returns_empty_dict(self):
        tables = parse_xer_string("")
        assert isinstance(tables, dict)

    def test_no_percent_e_still_parses(self):
        xer = """\
%T\tPROJECT
%F\tproj_id\tproj_name
%R\tP001\tTest
"""
        tables = parse_xer_string(xer)
        assert "PROJECT" in tables

    def test_table_with_no_rows_has_schema(self):
        xer = """\
%T\tPROJECT
%F\tproj_id\tproj_name
%E
"""
        tables = parse_xer_string(xer)
        assert "PROJECT" in tables
        assert list(tables["PROJECT"].columns) == ["proj_id", "proj_name"]
        assert tables["PROJECT"].empty


class TestProjectMeta:
    def test_returns_dict(self):
        tables = parse_xer_string(_MINIMAL_XER)
        meta = extract_project_meta(tables)
        assert isinstance(meta, dict)

    def test_required_keys_present(self):
        tables = parse_xer_string(_MINIMAL_XER)
        meta = extract_project_meta(tables)
        assert "proj_id" in meta
        assert "proj_name" in meta

    def test_proj_id_value(self):
        tables = parse_xer_string(_MINIMAL_XER)
        meta = extract_project_meta(tables)
        assert meta["proj_id"] == "P001"

    def test_empty_tables_returns_empty_dict(self):
        meta = extract_project_meta({})
        assert meta == {}
