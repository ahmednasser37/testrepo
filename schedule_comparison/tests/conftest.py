"""Shared fixtures for the schedule_comparison test suite."""

import sys
from pathlib import Path

# Ensure the schedule_comparison package root is importable
sys.path.insert(0, str(Path(__file__).parent.parent))

import pandas as pd
import pytest

from sample_generator import generate_sample_xer
from xer_parser import parse_xer_string


@pytest.fixture(scope="session")
def baseline_xer_str(tmp_path_factory):
    """Generate a realistic baseline XER and return its text content."""
    p = tmp_path_factory.mktemp("xer") / "baseline.xer"
    generate_sample_xer(p)
    return p.read_text(encoding="utf-8")


@pytest.fixture(scope="session")
def baseline_tables(baseline_xer_str):
    """Parsed tables from the baseline XER."""
    return parse_xer_string(baseline_xer_str)


@pytest.fixture(scope="session")
def updated_tables(baseline_xer_str):
    """
    Simulate a schedule update: the first three activities have their finish
    dates shifted forward by 30 days to model slippage.
    """
    tables = parse_xer_string(baseline_xer_str)
    task = tables.get("TASK")
    if task is not None and not task.empty and "target_end_date" in task.columns:
        task = task.copy()
        for idx in list(task.index[:3]):
            old = task.at[idx, "target_end_date"]
            if pd.notna(old):
                task.at[idx, "target_end_date"] = pd.Timestamp(old) + pd.Timedelta(days=30)
        tables = dict(tables)
        tables["TASK"] = task
    return tables
