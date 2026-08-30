"""
Shared fixtures for TicketAudit test suite.
"""
import sys
import os
import pytest
import pandas as pd

# Ensure project root is on the path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@pytest.fixture
def standard_df():
    """
    DataFrame with all 9 required columns using exact keyword names.

    This is the suite's "clean file" baseline, so it has to be clean on every
    dimension the analyzer checks - not just column names. Every ticket is
    closed *and* carries a resolution date: the earlier version mixed "Open"
    states with populated resolved_at values, which check_cross_field_logic
    correctly reports as a contradiction. Leaving the open rows without a
    resolution date is the other coherent option, but it puts resolved_at at 66%
    null and breaks the null-threshold tests that also rely on this fixture.

    Tests that need a mix of open and closed states should build their own frame
    (see _xfield_df in test_analyzer.py) rather than widen this one.
    """
    return pd.DataFrame({
        "number": ["INC0001", "INC0002", "INC0003"],
        "assignment_group": ["Network", "Desktop", "Network"],
        "cmdb_ci": ["Server-01", "Laptop-02", "Server-01"],
        "priority": ["P1", "P2", "P1"],
        "state": ["Closed", "Closed", "Closed"],
        "opened_at": pd.to_datetime(["2024-01-01", "2024-01-02", "2024-01-03"]),
        "resolved_at": pd.to_datetime(["2024-02-01", "2024-02-02", "2024-02-03"]),
        "short_description": ["Login issue", "Slow laptop", "Network down"],
        "description": ["User cannot log in", "Laptop runs slowly", "Network is down"],
    })


@pytest.fixture
def minimal_df():
    """DataFrame with only a few columns — some required columns are missing."""
    return pd.DataFrame({
        "ticket_id": ["INC001", "INC002"],
        "priority": ["P1", "P2"],
        "notes": ["Some note", "Another note"],
    })


@pytest.fixture
def df_with_nulls():
    """DataFrame where some columns have high null rates."""
    import numpy as np
    return pd.DataFrame({
        "number": ["INC001", "INC002", "INC003", "INC004", "INC005"],
        "priority": ["P1", None, None, None, "P2"],   # 60% nulls — above threshold
        # 20% nulls — at threshold. "Closed" rather than "Open" so the populated
        # resolved_at is not also a state/date contradiction; this fixture is
        # about null rates only.
        "state": ["Closed", "Closed", "Closed", None, "Closed"],
        "opened_at": pd.to_datetime(["2024-01-01"] * 5),
        "resolved_at": pd.to_datetime(["2024-02-01"] * 5),
        "short_description": ["desc"] * 5,
        "description": ["full"] * 5,
        "assignment_group": ["Team A"] * 5,
        "cmdb_ci": ["CI"] * 5,
    })


@pytest.fixture
def df_with_date_issues():
    """DataFrame with closed < created and future-dated rows."""
    return pd.DataFrame({
        "number": ["INC001", "INC002", "INC003"],
        "opened_at": pd.to_datetime(["2024-06-01", "2024-06-01", "2099-01-01"]),
        "resolved_at": pd.to_datetime(["2024-05-01", "2024-07-01", "2099-02-01"]),
        "assignment_group": ["Team"] * 3,
        "cmdb_ci": ["CI"] * 3,
        "priority": ["P1"] * 3,
        "state": ["Closed"] * 3,
        "short_description": ["d"] * 3,
        "description": ["d"] * 3,
    })


@pytest.fixture
def df_with_description_issues():
    """
    DataFrame exercising every description-quality dimension:
    empty, too short, too long, and the same issue repeated under different
    ticket IDs (which normalization should collapse together).
    """
    descriptions = [
        "Password reset required for account INC0001",  # ─┐ same issue,
        "Password reset required for account INC0002",  #  ├─ different IDs
        "Password reset required for account INC0003",  # ─┘
        "test",                                          # too short
        "n/a",                                           # too short
        "",                                              # empty
        "Printer on floor three is offline and refusing all queued jobs",
        "X" * 6000,                                      # too long
    ]
    n = len(descriptions)
    return pd.DataFrame({
        "number": [f"INC{i:04d}" for i in range(1, n + 1)],
        "assignment_group": ["Service Desk"] * n,
        "cmdb_ci": ["CI"] * n,
        "priority": ["P3"] * n,
        # "Closed" to match the populated resolved_at - this fixture is about
        # description quality and should exhibit no other finding.
        "state": ["Closed"] * n,
        "opened_at": pd.to_datetime(["2024-01-01"] * n),
        "resolved_at": pd.to_datetime(["2024-02-01"] * n),
        "short_description": descriptions,
        "description": ["A sufficiently long description for every row here"] * n,
    })


@pytest.fixture
def df_with_duplicates():
    """
    DataFrame with duplicate ticket IDs — and no other defect.

    The states are "Closed" to match the populated resolved_at: with "Open"
    there, check_cross_field_logic reports a state/date contradiction and the
    fixture no longer isolates duplicates, which is what its tests assert on.
    """
    return pd.DataFrame({
        "number": ["INC001", "INC001", "INC002", "INC003", "INC003"],
        "assignment_group": ["Team"] * 5,
        "cmdb_ci": ["CI"] * 5,
        "priority": ["P1"] * 5,
        "state": ["Closed"] * 5,
        "opened_at": pd.to_datetime(["2024-01-01"] * 5),
        "resolved_at": pd.to_datetime(["2024-02-01"] * 5),
        "short_description": ["d"] * 5,
        "description": ["d"] * 5,
    })
