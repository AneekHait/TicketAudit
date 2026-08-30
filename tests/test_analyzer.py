"""
Tests for core/analyzer.py — SanityAnalyzer.
"""
import pytest
import pandas as pd
import numpy as np
from core.analyzer import SanityAnalyzer, _score_description_rows, _SCORE_TIER_ORDER
from core.language import LanguageChecker


# ──────────────────────────────────────────────────────────────
# check_columns
# ──────────────────────────────────────────────────────────────

class TestCheckColumns:
    def test_exact_match_finds_all_required(self, standard_df):
        analyzer = SanityAnalyzer(standard_df)
        results = analyzer.check_columns()
        missing = [r for r, d in results.items() if not d['found']]
        assert missing == [], f"Unexpectedly missing: {missing}"

    def test_missing_columns_reported(self, minimal_df):
        analyzer = SanityAnalyzer(minimal_df)
        missing = analyzer.get_missing_columns()
        # minimal_df has no date, assignment group, CI, short_description, description, state columns
        assert "Created" in missing
        assert "Closed" in missing
        assert "Assignment Group" in missing

    def test_nonstandard_column_names_resolved(self):
        """Suffix/prefix matches should still be found."""
        df = pd.DataFrame({
            "inc_number": ["INC001"],           # prefix match for "number"
            "grp_assignment_group": ["Team"],   # suffix for "assignment_group"
            "ci_name": ["Server"],              # exact keyword
            "priority_level": ["P1"],           # suffix match
            "incident_state": ["Open"],         # exact keyword
            "sys_created_on": pd.to_datetime(["2024-01-01"]),
            "closed_at": pd.to_datetime(["2024-02-01"]),
            "short_description": ["Login issue"],
            "description": ["Full description"],
        })
        analyzer = SanityAnalyzer(df)
        results = analyzer.check_columns()
        assert results["Ticket Identification"]["found"]
        assert results["Priority"]["found"]
        assert results["State/Status"]["found"]

    def test_value_pattern_match_identifies_ticket_id_col(self):
        """A column named 'ref' with INC-style values should be found via value scanning."""
        df = pd.DataFrame({
            "ref": [f"INC{i:04d}" for i in range(1, 50)],
            "priority": ["P1"] * 49,
            "state": ["Open"] * 49,
            "opened_at": pd.to_datetime(["2024-01-01"] * 49),
            "resolved_at": pd.to_datetime(["2024-02-01"] * 49),
            "assignment_group": ["Team"] * 49,
            "cmdb_ci": ["CI"] * 49,
            "short_description": ["d"] * 49,
            "description": ["d"] * 49,
        })
        analyzer = SanityAnalyzer(df)
        results = analyzer.check_columns()
        assert results["Ticket Identification"]["found"]


# ──────────────────────────────────────────────────────────────
# check_nulls
# ──────────────────────────────────────────────────────────────

class TestCheckNulls:
    def test_no_nulls_none_above_threshold(self, standard_df):
        analyzer = SanityAnalyzer(standard_df)
        results = analyzer.check_nulls(threshold=20.0)
        assert not any(d["above_threshold"] for d in results.values())

    def test_high_null_column_flagged(self, df_with_nulls):
        analyzer = SanityAnalyzer(df_with_nulls)
        results = analyzer.check_nulls(threshold=20.0)
        # priority has 3/5 = 60% nulls → above threshold
        assert results["priority"]["above_threshold"]
        assert results["priority"]["percentage"] == pytest.approx(60.0)

    def test_exactly_at_threshold_not_flagged(self, df_with_nulls):
        analyzer = SanityAnalyzer(df_with_nulls)
        results = analyzer.check_nulls(threshold=20.0)
        # state has 1/5 = 20% nulls → NOT above (strict >)
        assert not results["state"]["above_threshold"]

    def test_just_above_threshold_flagged(self, df_with_nulls):
        analyzer = SanityAnalyzer(df_with_nulls)
        results = analyzer.check_nulls(threshold=19.0)
        # 20% > 19% → flagged
        assert results["state"]["above_threshold"]

    def test_null_counts_cached(self, standard_df):
        analyzer = SanityAnalyzer(standard_df)
        analyzer.check_nulls()
        # Second call should hit cache
        assert "null_counts" in analyzer._cache
        analyzer.check_nulls()  # no error on second call


# ──────────────────────────────────────────────────────────────
# check_date_logic
# ──────────────────────────────────────────────────────────────

class TestCheckDateLogic:
    def test_valid_dates_no_issues(self, standard_df):
        analyzer = SanityAnalyzer(standard_df)
        result = analyzer.check_date_logic()
        assert result["closed_before_created"] == 0
        assert result["future_created"] == 0
        assert result["future_closed"] == 0
        assert result["issues"] == []

    def test_closed_before_created_detected(self, df_with_date_issues):
        analyzer = SanityAnalyzer(df_with_date_issues)
        result = analyzer.check_date_logic()
        # row 0: resolved_at (2024-05-01) < opened_at (2024-06-01) — 1 bad row
        assert result["closed_before_created"] == 1

    def test_future_dates_detected(self, df_with_date_issues):
        analyzer = SanityAnalyzer(df_with_date_issues)
        result = analyzer.check_date_logic()
        # row 2: both dates are 2099 → future
        assert result["future_created"] >= 1
        assert result["future_closed"] >= 1

    def test_no_date_columns_returns_error(self):
        df = pd.DataFrame({"col_a": [1, 2, 3], "col_b": ["x", "y", "z"]})
        analyzer = SanityAnalyzer(df)
        result = analyzer.check_date_logic()
        assert result["has_date_columns"] is False
        assert len(result["issues"]) > 0


# ──────────────────────────────────────────────────────────────
# check_cross_field_logic
# ──────────────────────────────────────────────────────────────

def _xfield_df(states, priorities=None, resolved=None):
    """Build a frame with an exact, hand-countable state/date/priority layout."""
    n = len(states)
    return pd.DataFrame({
        "number": [f"INC{i:04d}" for i in range(n)],
        "state": states,
        "priority": priorities if priorities is not None else ["P1"] * n,
        "opened_at": pd.to_datetime(["2024-01-01"] * n),
        "resolved_at": pd.to_datetime(
            resolved if resolved is not None else ["2024-02-01"] * n),
        "assignment_group": ["Team"] * n,
        "cmdb_ci": ["CI"] * n,
        "short_description": ["d"] * n,
        "description": ["d"] * n,
    })


class TestCrossFieldLogic:
    def test_closed_without_resolution_date_flagged(self):
        df = _xfield_df(["Closed", "Closed", "Closed"],
                        resolved=["2024-02-01", None, None])
        result = SanityAnalyzer(df).check_cross_field_logic()
        assert result["closed_count"] == 3
        assert result["closed_no_date"] == 2

    def test_open_with_resolution_date_flagged(self):
        df = _xfield_df(["New", "In Progress", "New"],
                        resolved=[None, "2024-02-01", None])
        result = SanityAnalyzer(df).check_cross_field_logic()
        assert result["open_count"] == 3
        assert result["open_with_date"] == 1

    def test_open_missing_priority_flagged(self):
        """The regression that mattered: on pandas 3 astype(str) preserves NaN,
        so a sentinel-string test for blankness silently matched nothing."""
        df = _xfield_df(["New", "New", "Closed"],
                        priorities=["P1", None, "P2"], resolved=[None, None, "2024-02-01"])
        result = SanityAnalyzer(df).check_cross_field_logic()
        assert result["open_missing_priority"] == 1

    @pytest.mark.parametrize("blank", [None, "", "   ", float("nan")])
    def test_every_flavour_of_blank_priority_counts_as_missing(self, blank):
        df = _xfield_df(["New", "New"], priorities=["P1", blank], resolved=[None, None])
        result = SanityAnalyzer(df).check_cross_field_logic()
        assert result["open_missing_priority"] == 1, f"{blank!r} not treated as blank"

    def test_clean_data_reports_nothing(self):
        df = _xfield_df(["Closed", "Resolved"], resolved=["2024-02-01", "2024-02-01"])
        result = SanityAnalyzer(df).check_cross_field_logic()
        assert result["closed_no_date"] == 0
        assert result["open_with_date"] == 0
        assert result["open_missing_priority"] == 0
        assert result["unclassified_count"] == 0

    @pytest.mark.parametrize("value,expected", [
        ("Closed", "closed"), ("closed", "closed"), ("CLOSED", "closed"),
        ("Closed Complete", "closed"), ("Resolved", "closed"),
        ("Cancelled", "closed"), ("Completed", "closed"),
        ("New", "open"), ("In Progress", "open"), ("On Hold", "open"),
        ("Awaiting User Info", "open"), ("Reopened", "open"),
        ("Stage 7", "unclassified"), ("", "unclassified"),
    ])
    def test_state_vocabulary_classified(self, value, expected):
        df = _xfield_df([value] * 4, resolved=[None] * 4)
        result = SanityAnalyzer(df).check_cross_field_logic()
        key = {"closed": "closed_count", "open": "open_count",
               "unclassified": "unclassified_count"}[expected]
        assert result[key] == 4, f"{value!r} did not classify as {expected}"

    def test_resolved_pending_closure_counts_as_closed(self):
        """Contains both a terminal and an active keyword. Terminal wins, which
        is a documented judgement call, so pin it rather than leave it to drift."""
        df = _xfield_df(["Resolved Pending Closure"] * 3, resolved=[None] * 3)
        result = SanityAnalyzer(df).check_cross_field_logic()
        assert result["closed_count"] == 3
        assert result["closed_no_date"] == 3

    def test_unclassified_rows_excluded_not_assumed(self):
        df = _xfield_df(["Closed", "Stage 9"], resolved=[None, None])
        result = SanityAnalyzer(df).check_cross_field_logic()
        assert result["closed_count"] == 1
        assert result["open_count"] == 0
        assert result["unclassified_count"] == 1
        # The unknown row must not be counted as a closed-without-date issue
        assert result["closed_no_date"] == 1

    def test_no_state_column_reports_absence(self):
        df = pd.DataFrame({"number": ["INC1"], "opened_at": pd.to_datetime(["2024-01-01"])})
        result = SanityAnalyzer(df).check_cross_field_logic()
        assert result["has_state"] is False
        assert result["closed_no_date"] == 0

    def test_missing_priority_column_flagged_as_absent(self):
        df = _xfield_df(["New"], resolved=[None]).drop(columns=["priority"])
        result = SanityAnalyzer(df).check_cross_field_logic()
        assert result["has_priority"] is False
        assert result["open_missing_priority"] == 0


# ──────────────────────────────────────────────────────────────
# check_state_values
# ──────────────────────────────────────────────────────────────

class TestFoldedValueGrouping:
    """
    check_cross_field_logic classifies distinct values and uses isin, rather than
    transforming the whole column (0.43s -> 0.015s at 300k rows). That is only
    safe if it agrees exactly with the obvious row-wise implementation, so
    compare against one — the same guard TestCleanSeries applies to clean_series.
    """

    @staticmethod
    def _naive_counts(states, priorities):
        """Row-by-row reference implementation."""
        from core.analyzer import _classify_state
        open_n = closed_n = unknown_n = no_priority = 0
        for state, priority in zip(states, priorities):
            folded = "" if state is None or pd.isna(state) else str(state).strip().lower()
            group = _classify_state(folded)
            if group == "terminal":
                closed_n += 1
            elif group == "active":
                open_n += 1
                blank = (priority is None or pd.isna(priority)
                         or str(priority).strip() == "")
                if blank:
                    no_priority += 1
            else:
                unknown_n += 1
        return open_n, closed_n, unknown_n, no_priority

    @pytest.mark.parametrize("states,priorities", [
        (["Closed", "New", "Resolved"], ["P1", "P2", "P3"]),
        (["Closed", "New", None], ["P1", None, "P3"]),
        (["  Closed  ", "closed", "CLOSED"], ["P1", "P1", "P1"]),
        (["Stage 1", "Stage 2", "New"], [None, "P2", None]),
        ([None, None, None], [None, None, None]),
        (["", "   ", "New"], ["", "   ", None]),
        (["Closed", "Closed", "Closed"], ["P1", "  ", None]),
        (["Resolved Pending Closure", "On Hold", "Cancelled"], ["P1", None, "P3"]),
    ])
    def test_agrees_with_row_wise_implementation(self, states, priorities):
        df = _xfield_df(states, priorities=priorities, resolved=[None] * len(states))
        result = SanityAnalyzer(df).check_cross_field_logic()
        expected = self._naive_counts(states, priorities)
        assert (result["open_count"], result["closed_count"],
                result["unclassified_count"],
                result["open_missing_priority"]) == expected

    def test_counts_always_partition_the_frame(self):
        """Every row lands in exactly one of open/closed/unclassified."""
        states = ["Closed", "New", "Stage 9", None, "  ", "Resolved"]
        df = _xfield_df(states, resolved=[None] * len(states))
        r = SanityAnalyzer(df).check_cross_field_logic()
        assert r["open_count"] + r["closed_count"] + r["unclassified_count"] == len(states)


class TestStateValues:
    def test_clean_vocabulary_has_no_suspects(self):
        df = _xfield_df(["Closed"] * 30 + ["New"] * 30)
        result = SanityAnalyzer(df).check_state_values()
        assert result["total_distinct"] == 2
        assert result["suspect_values"] == []

    def test_typo_flagged_as_near_duplicate(self):
        """A 1-in-60 typo is 1.7% of the file, so frequency alone misses it."""
        df = _xfield_df(["Closed"] * 59 + ["Closd"])
        result = SanityAnalyzer(df).check_state_values()
        suspects = {e["value"]: e["reason"] for e in result["suspect_values"]}
        assert "Closd" in suspects
        assert "Closed" in suspects["Closd"]

    @pytest.mark.parametrize("variant", ["closed", "CLOSED", "Closed ", " Closed"])
    def test_case_and_whitespace_variants_flagged(self, variant):
        df = _xfield_df(["Closed"] * 40 + [variant] * 2)
        result = SanityAnalyzer(df).check_state_values()
        assert variant in {e["value"] for e in result["suspect_values"]}, \
            f"{variant!r} not reported as a variant of 'Closed'"

    def test_common_spelling_never_flagged_against_rare_one(self):
        df = _xfield_df(["Closed"] * 59 + ["Closd"])
        result = SanityAnalyzer(df).check_state_values()
        assert "Closed" in {e["value"] for e in result["standard_values"]}

    def test_distinct_codes_not_flagged_as_typos(self):
        """P1/P2-style values are short and similar but genuinely different."""
        df = _xfield_df(["P1"] * 25 + ["P2"] * 25 + ["P3"] * 10)
        result = SanityAnalyzer(df).check_state_values()
        assert result["suspect_values"] == []

    def test_rare_unrelated_value_flagged_by_frequency(self):
        df = _xfield_df(["Closed"] * 998 + ["Zzz Nonsense"] * 2)
        result = SanityAnalyzer(df).check_state_values()
        suspects = {e["value"]: e["reason"] for e in result["suspect_values"]}
        assert suspects.get("Zzz Nonsense") == "rare value"

    def test_rare_rule_stays_quiet_on_a_small_file(self):
        """1 row of 20 is 5% — normal at that scale, so frequency must not fire."""
        df = _xfield_df(["Closed"] * 10 + ["New"] * 9 + ["Zzz Nonsense"])
        result = SanityAnalyzer(df).check_state_values()
        assert result["suspect_values"] == []

    def test_blanks_excluded_from_vocabulary(self):
        df = _xfield_df(["Closed"] * 8 + [None, "   "])
        result = SanityAnalyzer(df).check_state_values()
        values = {e["value"] for e in
                  result["standard_values"] + result["suspect_values"]}
        assert not any(v.strip() == "" for v in values)
        assert result["total_distinct"] == 1

    def test_classification_carried_per_value(self):
        df = _xfield_df(["Closed"] * 20 + ["New"] * 20 + ["Stage 4"] * 20)
        result = SanityAnalyzer(df).check_state_values()
        reads = {e["value"]: e["classification"]
                 for e in result["standard_values"] + result["suspect_values"]}
        assert reads == {"Closed": "terminal", "New": "active", "Stage 4": "unknown"}

    def test_counts_and_percentages_correct(self):
        df = _xfield_df(["Closed"] * 30 + ["New"] * 10)
        result = SanityAnalyzer(df).check_state_values()
        by_value = {e["value"]: e for e in result["standard_values"]}
        assert by_value["Closed"]["count"] == 30
        assert by_value["Closed"]["pct"] == pytest.approx(75.0)
        assert by_value["New"]["pct"] == pytest.approx(25.0)

    def test_ordered_by_descending_count(self):
        df = _xfield_df(["New"] * 5 + ["Closed"] * 30 + ["Resolved"] * 12)
        result = SanityAnalyzer(df).check_state_values()
        counts = [e["count"] for e in result["standard_values"]]
        assert counts == sorted(counts, reverse=True)

    def test_high_cardinality_column_skips_fuzzy_matching(self):
        """A misdetected free-text column must not trigger an O(n^2) scan."""
        from core.analyzer import MAX_VOCABULARY_FOR_FUZZY
        n = MAX_VOCABULARY_FOR_FUZZY + 50
        df = _xfield_df([f"Value {i}" for i in range(n)])
        result = SanityAnalyzer(df).check_state_values()
        # Every value is unique, so none may be reported as a near-duplicate
        assert all(e["reason"] != "" and "looks like" not in e["reason"]
                   for e in result["suspect_values"])

    def test_no_state_column_reports_absence(self):
        df = pd.DataFrame({"number": ["INC1"], "opened_at": pd.to_datetime(["2024-01-01"])})
        result = SanityAnalyzer(df).check_state_values()
        assert result["has_state"] is False
        assert result["standard_values"] == []


# ──────────────────────────────────────────────────────────────
# check_duplicates
# ──────────────────────────────────────────────────────────────

class TestCheckDuplicates:
    def test_no_duplicates(self, standard_df):
        analyzer = SanityAnalyzer(standard_df)
        result = analyzer.check_duplicates()
        assert result["count"] == 0
        assert result["duplicate_ids"] == {}

    def test_duplicates_detected(self, df_with_duplicates):
        analyzer = SanityAnalyzer(df_with_duplicates)
        result = analyzer.check_duplicates()
        # INC001 appears twice → 1 extra; INC003 appears twice → 1 extra
        assert result["count"] == 2
        assert "INC001" in result["duplicate_ids"]
        assert "INC003" in result["duplicate_ids"]

    def test_all_unique(self):
        df = pd.DataFrame({
            "number": [f"INC{i:03d}" for i in range(1, 6)],
            "priority": ["P1"] * 5,
            "state": ["Open"] * 5,
            "opened_at": pd.to_datetime(["2024-01-01"] * 5),
            "resolved_at": pd.to_datetime(["2024-02-01"] * 5),
            "assignment_group": ["T"] * 5,
            "cmdb_ci": ["C"] * 5,
            "short_description": ["d"] * 5,
            "description": ["d"] * 5,
        })
        analyzer = SanityAnalyzer(df)
        result = analyzer.check_duplicates()
        assert result["count"] == 0

    def test_no_id_column_returns_gracefully(self):
        # Values are float — don't match any ticket-ID patterns; column names
        # don't match any COLUMN_KEYWORDS keyword, so id_col stays None.
        df = pd.DataFrame({
            "temperature": [98.6, 37.2],
            "pressure": [1.02, 0.98],
        })
        analyzer = SanityAnalyzer(df)
        result = analyzer.check_duplicates()
        assert result["id_col"] is None
        assert "error" in result


# ──────────────────────────────────────────────────────────────
# check_description_quality
# ──────────────────────────────────────────────────────────────

def _make_desc_df(descriptions):
    """Build a fully-mapped DataFrame around a list of short descriptions."""
    n = len(descriptions)
    return pd.DataFrame({
        "number": [f"INC{i:04d}" for i in range(1, n + 1)],
        "assignment_group": ["Team"] * n,
        "cmdb_ci": ["CI"] * n,
        "priority": ["P3"] * n,
        "state": ["Open"] * n,
        "opened_at": pd.to_datetime(["2024-01-01"] * n),
        "resolved_at": pd.to_datetime(["2024-02-01"] * n),
        "short_description": descriptions,
        "description": ["A sufficiently long description for every row here"] * n,
    })


class TestDescriptionColumnDiscovery:
    def test_finds_both_description_columns(self, standard_df):
        analyzer = SanityAnalyzer(standard_df)
        pairs = analyzer.description_columns
        categories = [category for category, _ in pairs]
        assert "Short Description" in categories
        assert "Description" in categories

    def test_no_description_column_returns_error(self):
        df = pd.DataFrame({"temperature": [98.6, 37.2], "pressure": [1.02, 0.98]})
        analyzer = SanityAnalyzer(df)
        result = analyzer.check_description_quality()
        assert result["columns"] == []
        assert "error" in result


class TestDescriptionLengthChecks:
    def test_clean_descriptions_have_no_length_issues(self, standard_df):
        # standard_df descriptions are short but meaningful; use a low threshold
        analyzer = SanityAnalyzer(standard_df)
        result = analyzer.check_description_quality(min_length=5, max_length=5000)
        short_desc = next(c for c in result["columns"]
                          if c["column"] == "short_description")
        assert short_desc["too_short"] == 0
        assert short_desc["too_long"] == 0
        assert short_desc["empty"] == 0

    def test_too_short_detected(self, df_with_description_issues):
        analyzer = SanityAnalyzer(df_with_description_issues)
        result = analyzer.check_description_quality(min_length=20, max_length=5000)
        short_desc = next(c for c in result["columns"]
                          if c["column"] == "short_description")
        # "test" and "n/a" are under 20 chars; the empty string is counted separately
        assert short_desc["too_short"] == 2
        assert "test" in short_desc["samples"]["too_short"]

    def test_too_long_detected(self, df_with_description_issues):
        analyzer = SanityAnalyzer(df_with_description_issues)
        result = analyzer.check_description_quality(min_length=20, max_length=5000)
        short_desc = next(c for c in result["columns"]
                          if c["column"] == "short_description")
        assert short_desc["too_long"] == 1

    def test_empty_counted_separately_from_too_short(self, df_with_description_issues):
        analyzer = SanityAnalyzer(df_with_description_issues)
        result = analyzer.check_description_quality()
        short_desc = next(c for c in result["columns"]
                          if c["column"] == "short_description")
        assert short_desc["empty"] == 1

    def test_length_stats_ignore_empty_rows(self):
        df = _make_desc_df(["", "abcdefghij", "abcdefghijabcdefghij"])
        analyzer = SanityAnalyzer(df)
        result = analyzer.check_description_quality(min_length=5)
        short_desc = next(c for c in result["columns"]
                          if c["column"] == "short_description")
        # min_length stat should be 10 (the shortest populated), not 0
        assert short_desc["min_length"] == 10
        assert short_desc["max_length"] == 20

    def test_thresholds_are_respected(self, df_with_description_issues):
        analyzer = SanityAnalyzer(df_with_description_issues)
        lenient = analyzer.check_description_quality(min_length=3, max_length=10000)
        strict = analyzer.check_description_quality(min_length=100, max_length=50)

        lenient_short = next(c for c in lenient["columns"]
                             if c["column"] == "short_description")
        strict_short = next(c for c in strict["columns"]
                            if c["column"] == "short_description")
        assert strict_short["too_short"] > lenient_short["too_short"]
        assert strict_short["too_long"] > lenient_short["too_long"]


class TestDescriptionRepetition:
    def test_same_issue_under_different_ticket_ids_groups_together(
        self, df_with_description_issues
    ):
        analyzer = SanityAnalyzer(df_with_description_issues)
        result = analyzer.check_description_quality()
        rep = next(r for r in result["repetition"]
                   if r["column"] == "short_description")

        assert rep["exact_group_count"] == 1
        assert rep["exact_repeated_rows"] == 3

        group = next(g for g in rep["top_groups"] if g["kind"] == "exact")
        assert group["count"] == 3
        # Sample IDs must be real ticket numbers so an SME can look them up
        assert group["sample_ids"] == ["INC0001", "INC0002", "INC0003"]

    def test_numbers_and_dates_are_normalized_away(self):
        df = _make_desc_df([
            "VPN connection dropping every 5 minutes from the home office",
            "VPN connection dropping every 30 minutes from the home office",
            "Completely unrelated storage array capacity warning raised",
        ])
        analyzer = SanityAnalyzer(df)
        result = analyzer.check_description_quality(min_length=5)
        rep = next(r for r in result["repetition"]
                   if r["column"] == "short_description")
        assert rep["exact_group_count"] == 1
        assert rep["exact_repeated_rows"] == 2

    def test_reworded_descriptions_cluster_as_similar(self):
        df = _make_desc_df([
            "Outlook crashes when opening calendar invitations repeatedly",
            "Calendar invitations repeatedly crashes Outlook",
            "Warehouse barcode scanner failing to transmit stock counts",
        ])
        analyzer = SanityAnalyzer(df)
        result = analyzer.check_description_quality(min_length=5)
        rep = next(r for r in result["repetition"]
                   if r["column"] == "short_description")

        assert rep["similar_group_count"] == 1
        similar = next(g for g in rep["top_groups"] if g["kind"] == "similar")
        assert similar["count"] == 2

    def test_distinct_descriptions_produce_no_repetition(self):
        df = _make_desc_df([
            "Warehouse barcode scanner failing to transmit stock counts",
            "Payroll export rejected by the downstream banking gateway",
            "Meeting room projector displaying a persistent green tint",
        ])
        analyzer = SanityAnalyzer(df)
        result = analyzer.check_description_quality(min_length=5)
        rep = next(r for r in result["repetition"]
                   if r["column"] == "short_description")
        assert rep["exact_group_count"] == 0
        assert rep["similar_group_count"] == 0
        assert rep["top_groups"] == []

    def test_falls_back_to_row_labels_without_id_column(self):
        df = pd.DataFrame({
            "short_description": ["Repeated issue text here"] * 3,
            "temperature": [1.5, 2.5, 3.5],
        })
        analyzer = SanityAnalyzer(df)
        result = analyzer.check_description_quality(min_length=5)
        rep = next(r for r in result["repetition"]
                   if r["column"] == "short_description")
        group = rep["top_groups"][0]
        assert all(s.startswith("row ") for s in group["sample_ids"])


class TestDescriptionQualityCaching:
    def test_result_is_cached_per_threshold(self, df_with_description_issues):
        analyzer = SanityAnalyzer(df_with_description_issues)
        first = analyzer.check_description_quality(min_length=20, max_length=5000)
        second = analyzer.check_description_quality(min_length=20, max_length=5000)
        assert first is second  # same object returned from cache

    def test_different_thresholds_are_cached_separately(self, df_with_description_issues):
        analyzer = SanityAnalyzer(df_with_description_issues)
        a = analyzer.check_description_quality(min_length=20)
        b = analyzer.check_description_quality(min_length=50)
        assert a is not b
        assert a["min_length"] == 20
        assert b["min_length"] == 50


class TestNormalizationIsSharedBetweenIdenticalColumns:
    """
    Normalizing description text is ~94% of the findings register, and Overview
    builds that register on every file open - so it dominates how long a file
    takes to become readable. ServiceNow's short_description and description
    very often hold the same text, and cleaning it twice for one answer took the
    register from 3.7s to 6.0s at 300k rows.

    Sharing a normalization between two columns is only safe if it is exactly
    the same answer, so that is what these assert. A cache that merged two
    *different* columns would silently combine their recurring-issue groups -
    a wrong answer, not a slow one.
    """

    @staticmethod
    def _frame(short_desc, desc):
        rows = len(short_desc)
        return pd.DataFrame({
            "number": [f"INC{i:05d}" for i in range(rows)],
            "opened_at": pd.date_range("2024-01-01", periods=rows, freq="h"),
            "state": ["Closed"] * rows,
            "short_description": short_desc,
            "description": desc,
        })

    @staticmethod
    def _uncached(analyzer, series):
        """The pre-cache implementation, so the fast path has a reference."""
        cleaned = LanguageChecker.clean_series(series).str.lower()
        cleaned = cleaned.str.replace(r'[^a-z0-9\s]', ' ', regex=True)
        return cleaned.str.replace(r'\s+', ' ', regex=True).str.strip()

    def test_identical_columns_are_normalized_once(self, mocker):
        """
        Asserted by counting the regex passes, not by inspecting the cache: the
        saving is the work not done, and a cache that stores the same answer
        twice would satisfy an identity check on nothing.
        """
        text = [f"Password reset failed for INC{i:05d} on 2024-01-01"
                for i in range(60)]
        spy = mocker.spy(LanguageChecker, "clean_series")
        analyzer = SanityAnalyzer(self._frame(list(text), list(text)))
        analyzer.check_description_quality(20, 5000)

        assert spy.call_count == 1, (
            f"identical columns cleaned {spy.call_count} times; the second "
            f"column must reuse the first")
        cache = analyzer._cache["normalized_text"]
        assert cache["short_description"] is cache["description"]

    def test_different_columns_are_each_normalized(self, mocker):
        """The other half: sharing must not be achieved by skipping work."""
        spy = mocker.spy(LanguageChecker, "clean_series")
        analyzer = SanityAnalyzer(self._frame(
            [f"Password reset failed for INC{i:05d}" for i in range(60)],
            [f"Mailbox quota exceeded on host {i}" for i in range(60)]))
        analyzer.check_description_quality(20, 5000)
        assert spy.call_count == 2, \
            "two columns of different text need two passes"

    def test_different_columns_do_not_share_an_answer(self):
        analyzer = SanityAnalyzer(self._frame(
            [f"Password reset failed for INC{i:05d}" for i in range(60)],
            [f"Mailbox quota exceeded on host {i}" for i in range(60)]))
        analyzer.check_description_quality(20, 5000)
        cache = analyzer._cache["normalized_text"]
        assert cache["short_description"] is not cache["description"]
        assert not cache["short_description"].equals(cache["description"]), \
            "sharing here would merge two columns' recurring-issue groups"

    @pytest.mark.parametrize("identical", [True, False])
    def test_the_cached_answer_equals_the_uncached_one(self, identical):
        short_desc = [f"VPN drops for user{i} at 10:{i:02d}" for i in range(60)]
        desc = list(short_desc) if identical else [
            f"Printer {i} offline in building {i % 4}" for i in range(60)]
        analyzer = SanityAnalyzer(self._frame(short_desc, desc))
        analyzer.check_description_quality(20, 5000)

        for column in ("short_description", "description"):
            raw = analyzer.df[column].astype(str)
            populated = raw[raw.str.strip() != ""]
            assert analyzer._cache["normalized_text"][column].equals(
                self._uncached(analyzer, populated)), \
                f"{column} normalized differently through the cache"

    @pytest.mark.parametrize("identical", [True, False])
    def test_repetition_findings_are_unchanged(self, identical):
        """
        The register's actual output, not just the intermediate - blanks
        included, because an empty string is dropped before normalizing and a
        shared mask has to survive that.
        """
        short_desc = [f"Recurring sync failure batch {i % 5}" for i in range(80)]
        for i in range(0, 80, 9):
            short_desc[i] = ""
        desc = list(short_desc) if identical else [
            f"Unique note number {i}" for i in range(80)]

        analyzer = SanityAnalyzer(self._frame(short_desc, desc))
        found = analyzer.check_description_quality(20, 5000)["repetition"]

        # A second analyzer sees each column first, so neither can be a reuse.
        expected = []
        for column in ("short_description", "description"):
            single = SanityAnalyzer(self._frame(
                short_desc if column == "short_description" else desc,
                desc if column == "description" else short_desc))
            for entry in single.check_description_quality(20, 5000)["repetition"]:
                if entry["column"] == column:
                    expected.append(entry)

        keys = ("column", "exact_group_count", "exact_repeated_rows",
                "similar_group_count", "similar_grouped_rows")
        assert [{k: e[k] for k in keys} for e in found] == \
               [{k: e[k] for k in keys} for e in expected]

    def test_an_entry_for_a_column_not_in_the_frame_is_not_reused(self):
        """
        Reuse is decided by comparing the raw columns the frame holds, so an
        entry naming a column that is gone cannot be compared and must be
        skipped. Unreachable through check_description_quality today; asserted
        so a future caller reusing an analyzer across frames fails loudly rather
        than silently inheriting another frame's answer.

        Every real entry is evicted first, so the poisoned one is the only
        candidate the loop can consider. Leaving the column's own entry in place
        would let it match itself and the guard would never be reached - which
        is what made an earlier version of this test pass against the bug.
        """
        text = [f"Same text everywhere {i % 3}" for i in range(40)]
        analyzer = SanityAnalyzer(self._frame(list(text), list(text)))
        analyzer.check_description_quality(20, 5000)

        cache = analyzer._cache["normalized_text"]
        evicted = cache.pop("description")
        cache.pop("short_description", None)
        assert not cache, "every real entry must be evicted for this to bite"
        cache["gone_column"] = pd.Series(["poisoned"] * len(evicted),
                                         index=evicted.index)

        raw = analyzer.df["description"].astype(str)
        populated = raw[raw.str.strip() != ""]
        result = analyzer._normalize_text_series(populated, "description")
        assert "poisoned" not in set(result), \
            "an entry for a column not in the frame must never be reused"
        assert result.equals(evicted)

    def test_a_differently_filtered_subset_is_not_given_another_columns_answer(self):
        """
        The length check guarding the reuse. Unreachable through today's single
        call site - both columns are filtered the same way - so it is exercised
        directly: a caller passing a narrower subset must get that subset
        normalized, not the full-length answer cached for an identical column.
        """
        text = [f"Recurring queue backlog {i % 4}" for i in range(40)]
        analyzer = SanityAnalyzer(self._frame(list(text), list(text)))

        raw = analyzer.df["short_description"].astype(str)
        populated = raw[raw.str.strip() != ""]
        analyzer._normalize_text_series(populated, "short_description")

        narrower = populated.head(10)
        result = analyzer._normalize_text_series(narrower, "description")
        assert len(result) == 10, (
            f"got {len(result)} rows for a 10-row subset - the identical "
            f"column's full-length answer was handed back")


# ──────────────────────────────────────────────────────────────
# Pivots
# ──────────────────────────────────────────────────────────────

class TestPivots:
    """
    The pivot view lists every column but displays one at a time. Building all
    of them up front cost ~4.5s on a 100k x 25 file, every visit, to show one —
    so listing must not count, and counting must be cached.
    """

    def test_pivot_columns_does_not_compute_counts(self, standard_df):
        analyzer = SanityAnalyzer(standard_df)
        analyzer.get_pivot_columns()
        assert not analyzer._cache.get('pivots'), \
            "listing columns must not populate the pivot cache"

    def test_operational_columns_come_first(self, standard_df):
        columns = SanityAnalyzer(standard_df).get_pivot_columns()
        assert "priority" in columns and "state" in columns
        # priority/state rank above an incidental column like cmdb_ci
        assert columns.index("priority") < columns.index("cmdb_ci")
        assert columns.index("state") < columns.index("cmdb_ci")

    def test_id_column_is_excluded(self, standard_df):
        analyzer = SanityAnalyzer(standard_df)
        assert analyzer.id_column not in analyzer.get_pivot_columns(), \
            "one row per ticket ID is not a useful pivot"

    def test_pivot_is_cached_per_column(self, standard_df):
        analyzer = SanityAnalyzer(standard_df)
        first = analyzer.get_pivot("priority")
        second = analyzer.get_pivot("priority")
        assert first is second

    def test_caching_does_not_leak_between_columns(self, standard_df):
        analyzer = SanityAnalyzer(standard_df)
        assert analyzer.get_pivot("priority")["column"] == "priority"
        assert analyzer.get_pivot("state")["column"] == "state"

    def test_lazy_and_eager_agree(self, standard_df):
        """get_all_pivots is still correct, and matches per-column results."""
        eager = SanityAnalyzer(standard_df).get_all_pivots()
        lazy_analyzer = SanityAnalyzer(standard_df)
        lazy = {c: lazy_analyzer.get_pivot(c)
                for c in lazy_analyzer.get_pivot_columns()}
        assert list(eager) == list(lazy)
        assert eager == lazy

    def test_counts_and_percentages_are_right(self, standard_df):
        pivot = SanityAnalyzer(standard_df).get_pivot("priority")
        by_value = {v["value"]: v for v in pivot["values"]}
        assert by_value["P1"]["count"] == 2
        assert by_value["P1"]["percentage"] == pytest.approx(66.67, abs=0.01)
        assert pivot["total_unique"] == 2

    def test_unknown_column_reports_error(self, standard_df):
        assert "error" in SanityAnalyzer(standard_df).get_pivot("nope")


# ──────────────────────────────────────────────────────────────
# Categorical dtype (memory optimisation on large files)
# ──────────────────────────────────────────────────────────────

def _categorize(df, *columns):
    out = df.copy()
    for col in columns:
        out[col] = out[col].astype("category")
    return out


class TestCategoricalColumns:
    """
    loader.optimize_dtypes stores low-cardinality string columns as `category`
    on large files, so every check has to work on both dtypes and agree.
    Two specific traps, both of which were live bugs:
      * fillna('') raises on a categorical unless '' is already a category
      * value_counts() reports unobserved categories with a count of 0
    """

    def test_description_quality_survives_categorical_with_nulls(self):
        """The fillna('') trap: fine without nulls, TypeError with them."""
        df = pd.DataFrame({
            "number": [f"INC{i:04d}" for i in range(6)],
            "short_description": [None, "Printer offline refusing jobs",
                                  "Printer offline refusing jobs", None,
                                  "VPN drops out", "VPN drops out"],
            "description": ["A sufficiently long description here"] * 6,
            "opened_at": pd.to_datetime(["2024-01-01"] * 6),
            "resolved_at": pd.to_datetime(["2024-02-01"] * 6),
        })
        plain = SanityAnalyzer(df).check_description_quality()
        cat = SanityAnalyzer(_categorize(df, "short_description")) \
            .check_description_quality()
        assert cat["columns"][0]["empty"] == plain["columns"][0]["empty"] == 2

    def test_pivot_omits_unobserved_categories(self):
        """A category present in the dtype but not the data is not a value."""
        series = pd.Series(["Open", "Open", "Closed"]).astype(
            pd.CategoricalDtype(["Open", "Closed", "Cancelled"]))
        df = pd.DataFrame({"number": ["INC1", "INC2", "INC3"], "state": series})
        pivot = SanityAnalyzer(df).get_pivot("state")
        assert [v["value"] for v in pivot["values"]] == ["Open", "Closed"]
        assert all(v["count"] > 0 for v in pivot["values"])

    def test_pivot_order_is_stable_across_dtypes(self):
        """Tied counts ordered by dtype internals gave two different orderings."""
        df = pd.DataFrame({
            "number": [f"INC{i}" for i in range(6)],
            "team": ["Wintel", "Network", "Service Desk"] * 2,
        })
        plain = SanityAnalyzer(df).get_pivot("team")
        cat = SanityAnalyzer(_categorize(df, "team")).get_pivot("team")
        assert [v["value"] for v in plain["values"]] == \
               [v["value"] for v in cat["values"]]

    def test_data_profile_reports_value_type_not_storage(self):
        """Showing "category" would leak an optimisation into the UI."""
        df = pd.DataFrame({"state": ["Open", "Closed", "Open"]})
        profile = SanityAnalyzer(_categorize(df, "state")).get_data_profile()
        assert profile["columns"][0]["dtype"] != "category"
        plain = SanityAnalyzer(df).get_data_profile()
        assert profile["columns"][0]["dtype"] == plain["columns"][0]["dtype"]

    def test_whitespace_detection_still_works(self, standard_df):
        df = standard_df.copy()
        df["state"] = ["  Open", "Closed", "Open"]
        cat = SanityAnalyzer(_categorize(df, "state")).get_data_profile()
        flagged = {c["name"]: c["has_whitespace"] for c in cat["columns"]}
        assert flagged["state"] is True

    @pytest.mark.parametrize("check", [
        "check_columns", "check_nulls", "check_date_logic", "check_duplicates",
        "get_monthly_inflow", "get_summary", "get_pivot_columns",
        "check_description_quality", "check_cross_field_logic",
        "check_state_values",
    ])
    def test_results_identical_on_both_dtypes(self, standard_df, check):
        cat_cols = ["priority", "state", "assignment_group", "short_description"]
        plain = getattr(SanityAnalyzer(standard_df), check)()
        cat = getattr(SanityAnalyzer(_categorize(standard_df, *cat_cols)), check)()
        assert plain == cat, f"{check} differs once columns are categorical"

    def test_duplicates_found_on_categorical_id(self, df_with_duplicates):
        plain = SanityAnalyzer(df_with_duplicates).check_duplicates()
        cat = SanityAnalyzer(_categorize(df_with_duplicates, "number")).check_duplicates()
        assert plain["count"] == cat["count"] > 0
        assert plain["duplicate_ids"] == cat["duplicate_ids"]


# ──────────────────────────────────────────────────────────────
# get_summary
# ──────────────────────────────────────────────────────────────

class TestGetSummary:
    def test_clean_data_no_issues(self, standard_df):
        analyzer = SanityAnalyzer(standard_df)
        summary = analyzer.get_summary()
        assert not summary["has_issues"]
        assert summary["missing_columns"] == []
        assert summary["duplicate_count"] == 0

    def test_issues_aggregated(self, df_with_nulls):
        analyzer = SanityAnalyzer(df_with_nulls)
        summary = analyzer.get_summary(null_threshold=20.0)
        assert summary["has_issues"]
        assert len(summary["high_null_columns"]) >= 1


# ──────────────────────────────────────────────────────────────
# Description quality score
# ──────────────────────────────────────────────────────────────

def _desc_df(descriptions, short_descs=None, n=None):
    """Minimal DataFrame with description columns for score tests."""
    if n is None:
        n = len(descriptions)
    return pd.DataFrame({
        "number": [f"INC{i:04d}" for i in range(n)],
        "state": ["Closed"] * n,
        "priority": ["P1"] * n,
        "opened_at": pd.to_datetime(["2024-01-01"] * n),
        "resolved_at": pd.to_datetime(["2024-02-01"] * n),
        "assignment_group": ["Team"] * n,
        "cmdb_ci": ["CI"] * n,
        "short_description": short_descs if short_descs is not None else descriptions,
        "description": descriptions,
    })


class TestScoreDescriptionRows:
    """Unit tests for the module-level _score_description_rows helper."""

    def _series(self, lengths_list):
        return pd.Series(lengths_list, dtype=int)

    def _uniform_mask(self, n, value):
        return pd.Series([value] * n)

    def test_empty_rows_score_zero_and_critical(self):
        lengths = self._series([0, 0])
        empty = self._uniform_mask(2, True)
        short = self._uniform_mask(2, False)
        long_ = self._uniform_mask(2, False)
        gs = self._series([0, 0])
        scores, tiers = _score_description_rows(lengths, empty, short, long_, gs, 20)
        assert list(scores) == [0, 0]
        assert list(tiers) == ['Critical', 'Critical']

    def test_short_rows_score_poor(self):
        lengths = self._series([5, 10])
        empty = self._uniform_mask(2, False)
        short = self._uniform_mask(2, True)
        long_ = self._uniform_mask(2, False)
        gs = self._series([1, 1])
        scores, tiers = _score_description_rows(lengths, empty, short, long_, gs, 20)
        assert list(tiers) == ['Poor', 'Poor']
        assert all(5 <= s <= 39 for s in scores)
        assert scores.iloc[1] > scores.iloc[0]  # longer short text scores higher

    def test_too_long_rows_score_fair(self):
        lengths = self._series([6000])
        empty = self._uniform_mask(1, False)
        short = self._uniform_mask(1, False)
        long_ = self._uniform_mask(1, True)
        gs = self._series([1])
        scores, tiers = _score_description_rows(lengths, empty, short, long_, gs, 20)
        assert tiers.iloc[0] == 'Fair'
        assert 40 <= scores.iloc[0] < 65

    def test_large_template_scores_fair(self):
        lengths = self._series([100])
        empty = self._uniform_mask(1, False)
        short = self._uniform_mask(1, False)
        long_ = self._uniform_mask(1, False)
        gs = self._series([15])  # 15 tickets share this text
        scores, tiers = _score_description_rows(lengths, empty, short, long_, gs, 20)
        assert tiers.iloc[0] == 'Fair'
        assert scores.iloc[0] == 45

    def test_moderate_repeat_scores_fair(self):
        lengths = self._series([100])
        empty = self._uniform_mask(1, False)
        short = self._uniform_mask(1, False)
        long_ = self._uniform_mask(1, False)
        gs = self._series([5])  # 5 tickets share this text
        scores, tiers = _score_description_rows(lengths, empty, short, long_, gs, 20)
        assert tiers.iloc[0] == 'Fair'
        assert scores.iloc[0] == 58

    def test_unique_moderate_length_scores_good(self):
        lengths = self._series([80])  # 80 chars, well above 20 min
        empty = self._uniform_mask(1, False)
        short = self._uniform_mask(1, False)
        long_ = self._uniform_mask(1, False)
        gs = self._series([1])
        scores, tiers = _score_description_rows(lengths, empty, short, long_, gs, 20)
        assert tiers.iloc[0] == 'Good'
        assert 65 <= scores.iloc[0] < 90

    def test_unique_long_description_scores_excellent(self):
        lengths = self._series([1100])  # over 1000 chars bonus cap
        empty = self._uniform_mask(1, False)
        short = self._uniform_mask(1, False)
        long_ = self._uniform_mask(1, False)
        gs = self._series([1])
        scores, tiers = _score_description_rows(lengths, empty, short, long_, gs, 20)
        assert tiers.iloc[0] == 'Excellent'
        assert scores.iloc[0] == 100

    def test_tier_order_constant_covers_all_tiers_returned(self):
        lengths = self._series([0, 5, 6000, 50, 1100])
        empty = pd.Series([True, False, False, False, False])
        short = pd.Series([False, True, False, False, False])
        long_ = pd.Series([False, False, True, False, False])
        gs = self._series([0, 1, 1, 1, 1])
        _, tiers = _score_description_rows(lengths, empty, short, long_, gs, 20)
        assert set(tiers).issubset(set(_SCORE_TIER_ORDER))


class TestDescriptionScoreIntegration:
    """Integration tests: overall_score and score_distribution come back from
    check_description_quality and have the expected structure."""

    def test_score_distribution_present_in_result(self, df_with_description_issues):
        result = SanityAnalyzer(df_with_description_issues).check_description_quality()
        for col_info in result['columns']:
            assert 'score_distribution' in col_info
            dist = col_info['score_distribution']
            assert set(dist.keys()) == set(_SCORE_TIER_ORDER)

    def test_score_distribution_sums_to_total_rows(self, df_with_description_issues):
        analyzer = SanityAnalyzer(df_with_description_issues)
        result = analyzer.check_description_quality()
        total = analyzer.total_rows
        for col_info in result['columns']:
            assert sum(col_info['score_distribution'].values()) == total

    def test_overall_score_present_and_in_range(self, df_with_description_issues):
        result = SanityAnalyzer(df_with_description_issues).check_description_quality()
        for col_info in result['columns']:
            assert 'overall_score' in col_info
            assert 0 <= col_info['overall_score'] <= 100

    def test_all_empty_column_scores_zero(self):
        descs = [""] * 5
        df = _desc_df(descs)
        result = SanityAnalyzer(df).check_description_quality()
        for col_info in result['columns']:
            assert col_info['overall_score'] == 0

    def test_all_excellent_column_scores_100(self):
        descs = [f"{'a' * 1100} ticket {i}" for i in range(5)]  # unique, well above min
        df = _desc_df(descs)
        result = SanityAnalyzer(df).check_description_quality()
        for col_info in result['columns']:
            assert col_info['overall_score'] == 100

    def test_template_text_lowers_score(self):
        template = "User unable to login to the system"
        good = "a" * 500
        descs = [template] * 15 + [good] * 5
        df = _desc_df(descs)
        result = SanityAnalyzer(df).check_description_quality()
        for col_info in result['columns']:
            dist = col_info['score_distribution']
            assert dist['Fair'] >= 15
            assert col_info['overall_score'] < 80

    def test_mixed_quality_score_is_between_extremes(self):
        descs = ["a" * 500] * 5 + [""] * 5  # 5 excellent, 5 critical
        df = _desc_df(descs)
        result = SanityAnalyzer(df).check_description_quality()
        for col_info in result['columns']:
            assert 0 < col_info['overall_score'] < 100

    def test_result_is_still_cached_after_scoring(self, df_with_description_issues):
        analyzer = SanityAnalyzer(df_with_description_issues)
        first = analyzer.check_description_quality()
        second = analyzer.check_description_quality()
        assert first is second


# ──────────────────────────────────────────────────────────────
# Per-row flags — one source of truth for counts and row lists
# ──────────────────────────────────────────────────────────────

class TestRowFlags:
    """
    get_row_flags() exists so a rule's headline count and its per-row answer
    come from the same array. If they were computed separately, the Logic Checks
    sheet and the Ticket Detail sheet could disagree about the same file — worse
    than either being wrong alone.
    """

    def test_every_flag_sums_to_the_count_its_check_reports(self):
        df = _xfield_df(
            states=["Closed"] * 4 + ["Open"] * 4,
            priorities=["P1", "P1", None, "  "] + ["P2"] * 4,
            resolved=["2024-02-01", None, "2024-02-01", "2024-02-01"]
                     + ["2024-02-01", None, None, None],
        )
        analyzer = SanityAnalyzer(df)
        flags = analyzer.get_row_flags()['flags']
        xfield = analyzer.check_cross_field_logic()

        assert int(flags[SanityAnalyzer.FLAG_CLOSED_NO_DATE].sum()) == \
            xfield['closed_no_date']
        assert int(flags[SanityAnalyzer.FLAG_OPEN_WITH_DATE].sum()) == \
            xfield['open_with_date']
        assert int(flags[SanityAnalyzer.FLAG_OPEN_MISSING_PRIORITY].sum()) == \
            xfield['open_missing_priority']

    def test_date_flags_sum_to_check_date_logic(self, df_with_date_issues):
        analyzer = SanityAnalyzer(df_with_date_issues)
        flags = analyzer.get_row_flags()['flags']
        dates = analyzer.check_date_logic()

        assert int(flags[SanityAnalyzer.FLAG_CLOSED_BEFORE_CREATED].sum()) == \
            dates['closed_before_created']
        assert int(flags[SanityAnalyzer.FLAG_FUTURE_CREATED].sum()) == \
            dates['future_created']
        assert int(flags[SanityAnalyzer.FLAG_FUTURE_CLOSED].sum()) == \
            dates['future_closed']

    def test_a_skipped_rule_is_absent_not_all_false(self):
        """
        Present-but-empty is indistinguishable from "checked, nothing found",
        which is the false-pass this project designs against.
        """
        df = _xfield_df(states=["Closed"] * 3).drop(columns=["state"])
        result = SanityAnalyzer(df).get_row_flags()

        for flag in (SanityAnalyzer.FLAG_CLOSED_NO_DATE,
                     SanityAnalyzer.FLAG_OPEN_WITH_DATE,
                     SanityAnalyzer.FLAG_OPEN_MISSING_PRIORITY):
            assert flag in result['skipped']
            assert flag not in result['flags']

    def test_missing_priority_column_skips_only_that_rule(self):
        df = _xfield_df(states=["Open"] * 3).drop(columns=["priority"])
        result = SanityAnalyzer(df).get_row_flags()
        assert SanityAnalyzer.FLAG_OPEN_MISSING_PRIORITY in result['skipped']
        # the state/date rules could still run
        assert SanityAnalyzer.FLAG_CLOSED_NO_DATE in result['flags']

    def test_missing_date_columns_skip_the_date_rules(self):
        df = _xfield_df(states=["Open"] * 3).drop(
            columns=["opened_at", "resolved_at"])
        result = SanityAnalyzer(df).get_row_flags()
        assert SanityAnalyzer.FLAG_CLOSED_BEFORE_CREATED in result['skipped']
        assert SanityAnalyzer.FLAG_FUTURE_CREATED in result['skipped']

    def test_flags_are_aligned_to_the_frame_index(self):
        df = _xfield_df(states=["Open", "Closed", "Open"])
        df.index = [10, 20, 30]
        flags = SanityAnalyzer(df).get_row_flags()['flags']
        for mask in flags.values():
            assert list(mask.index) == [10, 20, 30]

    def test_state_masks_do_not_overlap(self):
        df = _xfield_df(states=["Open", "Closed", "Stage 9", None])
        state = SanityAnalyzer(df).get_row_flags()['state']
        assert not (state['terminal'] & state['active']).any()

    def test_unclassified_rows_are_excluded_from_both_state_masks(self):
        df = _xfield_df(states=["Open", "Closed", "Stage 9"])
        analyzer = SanityAnalyzer(df)
        state = analyzer.get_row_flags()['state']
        assert int(state['active'].sum()) == 1
        assert int(state['terminal'].sum()) == 1
        assert analyzer.check_cross_field_logic()['unclassified_count'] == 1

    def test_result_is_cached(self):
        analyzer = SanityAnalyzer(_xfield_df(states=["Open"] * 3))
        assert analyzer.get_row_flags() is analyzer.get_row_flags()

    def test_no_state_column_still_returns_the_contract(self):
        df = _xfield_df(states=["Open"] * 3).drop(columns=["state"])
        result = SanityAnalyzer(df).get_row_flags()
        assert set(result) == {'flags', 'skipped', 'state'}
        assert result['state'] == {}


class TestColumnOverrides:
    """
    Automatic matching is a heuristic over column names, so it can pick the
    wrong field - and then every check measures the wrong thing, confidently.
    Before overrides existed there was no way to correct it.
    """

    @pytest.fixture
    def ambiguous_priority_df(self):
        """
        A ServiceNow-shaped export with no column literally named "priority".
        impact and urgency both score 130, so which one wins is arbitrary.
        """
        n = 4
        return pd.DataFrame({
            "number": [f"INC{i:04d}" for i in range(n)],
            "state": ["New"] * n,
            "urgency": ["P1", "P2", None, None],
            "impact": ["1", None, None, None],
            "assignment_group": ["Team"] * n,
            "cmdb_ci": ["CI"] * n,
            "opened_at": pd.to_datetime(["2024-01-01"] * n),
            "resolved_at": pd.to_datetime(["2024-02-01"] * n),
            "short_description": ["Login issue for the user account"] * n,
            "description": ["A sufficiently detailed description"] * n,
        })

    def test_scores_are_retained_for_every_candidate(self, ambiguous_priority_df):
        choices = SanityAnalyzer(ambiguous_priority_df).column_choices()
        priority = choices["Priority"]
        assert priority['score'] > 0
        assert priority['alternatives'], "the runner-up must be visible"

    def test_a_tie_is_visible_rather_than_hidden(self, ambiguous_priority_df):
        """impact and urgency score the same; the user has to be able to see it."""
        choices = SanityAnalyzer(ambiguous_priority_df).column_choices()
        priority = choices["Priority"]
        runner_up_score = priority['alternatives'][0][1]
        assert runner_up_score == priority['score']

    def test_an_override_redirects_the_checks(self, ambiguous_priority_df):
        """The whole point: a different column means different findings."""
        urgency = SanityAnalyzer(ambiguous_priority_df,
                                 column_overrides={"Priority": "urgency"})
        impact = SanityAnalyzer(ambiguous_priority_df,
                                column_overrides={"Priority": "impact"})
        assert urgency.priority_column == "urgency"
        assert impact.priority_column == "impact"
        assert (urgency.check_cross_field_logic()['open_missing_priority']
                != impact.check_cross_field_logic()['open_missing_priority'])

    def test_an_override_also_moves_the_derived_columns(self, ambiguous_priority_df):
        """id/created/closed are derived from the map, so they must follow it."""
        analyzer = SanityAnalyzer(ambiguous_priority_df,
                                  column_overrides={"Created": "resolved_at"})
        assert analyzer.created_column == "resolved_at"

    def test_check_columns_flags_an_overridden_match(self, ambiguous_priority_df):
        results = SanityAnalyzer(
            ambiguous_priority_df,
            column_overrides={"Priority": "urgency"}).check_columns()
        assert results["Priority"]['overridden'] is True
        assert results["State/Status"]['overridden'] is False

    def test_a_category_can_be_declared_absent(self, ambiguous_priority_df):
        """
        Not the same as leaving it to the matcher: it forces the checks that
        need it to report as not checked rather than reading a wrong column.
        """
        from core.analyzer import COLUMN_ABSENT
        analyzer = SanityAnalyzer(ambiguous_priority_df,
                                  column_overrides={"Priority": COLUMN_ABSENT})
        assert analyzer.priority_column is None
        assert "Priority" in analyzer.get_missing_columns()
        assert (SanityAnalyzer.FLAG_OPEN_MISSING_PRIORITY
                in analyzer.get_row_flags()['skipped'])

    def test_a_stale_override_is_reported_not_silently_dropped(
            self, ambiguous_priority_df):
        """A setting carried over from another export shape must not vanish."""
        analyzer = SanityAnalyzer(ambiguous_priority_df,
                                  column_overrides={"Priority": "not_a_column"})
        assert analyzer.priority_column is not None, "falls back to the match"
        assert any("ignored" in note for note in analyzer.override_notes)
        assert any("not_a_column" in note for note in analyzer.override_notes)

    def test_an_unknown_category_is_reported(self, ambiguous_priority_df):
        analyzer = SanityAnalyzer(ambiguous_priority_df,
                                  column_overrides={"Nonsense": "impact"})
        assert any("not one of the required categories" in note
                   for note in analyzer.override_notes)

    def test_an_applied_override_is_described(self, ambiguous_priority_df):
        analyzer = SanityAnalyzer(ambiguous_priority_df,
                                  column_overrides={"Priority": "urgency"})
        assert any("urgency" in note for note in analyzer.override_notes)

    def test_no_overrides_means_no_notes(self, ambiguous_priority_df):
        analyzer = SanityAnalyzer(ambiguous_priority_df)
        assert analyzer.override_notes == []
        assert analyzer.overridden == {}

    def test_an_override_can_pick_a_column_that_scored_nothing(
            self, ambiguous_priority_df):
        """The matcher scoring zero is exactly when an override is needed."""
        analyzer = SanityAnalyzer(ambiguous_priority_df,
                                  column_overrides={"Priority": "cmdb_ci"})
        assert analyzer.priority_column == "cmdb_ci"
        assert analyzer.check_columns()["Priority"]['score'] == 0

    def test_ties_still_break_by_file_order(self, ambiguous_priority_df):
        """
        Stable by column position, not alphabetical: exports put the more
        significant field earlier, and re-breaking ties by name would change
        which column existing files match.
        """
        choices = SanityAnalyzer(ambiguous_priority_df).column_choices()
        assert choices["Priority"]['chosen'] == "urgency"


class TestDateFieldOrder:
    """
    pandas will guess DD/MM vs MM/DD, and when nothing in the column
    disambiguates it assumes month-first. A European export of 05/03/2024 is
    then read as 5 May instead of 3 March, so a 3-day resolution is reported as
    92 while every check says "no issues". Guessing is unavoidable; guessing
    silently is the bug.
    """

    def _dates(self, opened, resolved=None):
        n = len(opened)
        return pd.DataFrame({
            "number": [f"INC{i:04d}" for i in range(n)],
            "state": ["Closed"] * n,
            "priority": ["P1"] * n,
            "assignment_group": ["Team"] * n,
            "cmdb_ci": ["CI"] * n,
            "opened_at": opened,
            "resolved_at": resolved if resolved is not None else opened,
            "short_description": ["Login issue for the user account"] * n,
            "description": ["A sufficiently detailed description"] * n,
        })

    def test_day_first_is_proven_by_a_day_over_twelve(self):
        df = self._dates(["13/01/2024", "03/04/2024"])
        info = SanityAnalyzer(df).check_date_formats()['columns'][0]
        assert info['order'] == "day_first"
        assert info['needs_confirmation'] is False

    def test_month_first_is_proven_by_a_second_field_over_twelve(self):
        df = self._dates(["01/13/2024", "04/03/2024"])
        info = SanityAnalyzer(df).check_date_formats()['columns'][0]
        assert info['order'] == "month_first"
        assert info['needs_confirmation'] is False

    def test_an_all_ambiguous_column_asks_for_confirmation(self):
        df = self._dates(["05/03/2024", "02/04/2024", "10/05/2024"])
        info = SanityAnalyzer(df).check_date_formats()['columns'][0]
        assert info['order'] == "ambiguous"
        assert info['needs_confirmation'] is True
        assert info['ambiguous'] == 3
        assert info['examples']

    def test_it_prices_the_other_reading(self):
        """A count of ambiguous values is abstract; rows changed is not."""
        df = self._dates(["05/03/2024", "02/04/2024", "10/05/2024"])
        info = SanityAnalyzer(df).check_date_formats()['columns'][0]
        assert info['rows_that_would_change'] == 3

    def test_iso_dates_need_no_decision(self):
        df = self._dates(["2024-03-05", "2024-04-02"])
        assert SanityAnalyzer(df).check_date_formats()['columns'] == []

    def test_a_real_datetime_column_needs_no_decision(self):
        df = self._dates(pd.to_datetime(["2024-03-05", "2024-04-02"]))
        assert SanityAnalyzer(df).check_date_formats()['columns'] == []

    def test_mixed_orders_are_flagged_as_a_defect_not_a_choice(self):
        """
        13/01 and 01/13 in one column: no single reading is right, so there is
        nothing to choose. pandas resolves both to 13 January, so the two
        readings agree while the data is still wrong - which is why this cannot
        depend on rows_that_would_change.
        """
        df = self._dates(["13/01/2024", "01/13/2024"])
        info = SanityAnalyzer(df).check_date_formats()['columns'][0]
        assert info['order'] == "conflicting"
        assert info['conflicting'] is True
        assert info['day_first_rows'] == 1
        assert info['month_first_rows'] == 1

    def test_the_setting_is_reported(self):
        df = self._dates(["05/03/2024"])
        assert SanityAnalyzer(df).check_date_formats()['setting'] == "inferred"
        assert SanityAnalyzer(df, dayfirst=True).check_date_formats()[
            'setting'] == "day_first"

    def test_choosing_day_first_changes_the_parse(self):
        df = self._dates(["05/03/2024"], ["08/03/2024"])
        inferred = SanityAnalyzer(df)
        chosen = SanityAnalyzer(df, dayfirst=True)
        assert inferred._get_created_dates()[0] == pd.Timestamp("2024-05-03")
        assert chosen._get_created_dates()[0] == pd.Timestamp("2024-03-05")

    def test_choosing_the_order_fixes_the_derived_numbers(self):
        """The whole point: resolution time was 30x wrong."""
        df = self._dates(["05/03/2024", "02/04/2024"],
                         ["08/03/2024", "09/04/2024"])
        wrong = SanityAnalyzer(df)
        right = SanityAnalyzer(df, dayfirst=True)

        def days(analyzer):
            return list((analyzer._get_closed_dates()
                         - analyzer._get_created_dates()).dt.days)

        assert days(wrong) == [92, 213]
        assert days(right) == [3, 7]

    def test_the_result_is_cached(self):
        analyzer = SanityAnalyzer(self._dates(["05/03/2024"]))
        assert analyzer.check_date_formats() is analyzer.check_date_formats()

    def test_no_date_column_is_handled(self, standard_df):
        df = standard_df.drop(columns=["opened_at", "resolved_at"])
        assert SanityAnalyzer(df).check_date_formats()['columns'] == []


class TestDetectDateOrderUnit:
    """The detector alone, on raw values."""

    def test_text_that_is_not_a_date(self):
        from core.analyzer import _detect_date_order
        assert _detect_date_order(["hello", "world"])['order'] == "not_numeric"

    def test_empty_input(self):
        from core.analyzer import _detect_date_order
        assert _detect_date_order([])['order'] == "not_numeric"

    def test_dashes_and_dots_are_recognised(self):
        from core.analyzer import _detect_date_order
        assert _detect_date_order(["13-01-2024"])['order'] == "day_first"
        assert _detect_date_order(["13.01.2024"])['order'] == "day_first"

    def test_two_digit_years(self):
        from core.analyzer import _detect_date_order
        assert _detect_date_order(["13/01/24"])['order'] == "day_first"

    def test_non_strings_are_skipped(self):
        from core.analyzer import _detect_date_order
        result = _detect_date_order([None, 5, pd.Timestamp("2024-01-01")])
        assert result['order'] == "not_numeric"

    def test_it_reads_distinct_values_not_rows(self):
        """The answer is a property of the format, so duplicates cost nothing."""
        from core.analyzer import _detect_date_order
        one = _detect_date_order(["05/03/2024"])
        many = _detect_date_order(["05/03/2024"] * 50)
        assert one['order'] == many['order']


class TestDuplicateRowMask:
    def test_mask_marks_every_member_of_a_duplicate_group(self):
        df = _xfield_df(states=["Closed"] * 4)
        df["number"] = ["INC1", "INC1", "INC2", "INC3"]
        mask = SanityAnalyzer(df).duplicate_row_mask()
        assert list(mask) == [True, True, False, False]

    def test_mask_and_count_measure_different_things(self):
        """
        The mask marks rows to inspect; check_duplicates counts surplus rows.
        A value appearing twice is 2 in the mask and 1 in the count.
        """
        df = _xfield_df(states=["Closed"] * 4)
        df["number"] = ["INC1", "INC1", "INC2", "INC3"]
        analyzer = SanityAnalyzer(df)
        assert int(analyzer.duplicate_row_mask().sum()) == 2
        assert analyzer.check_duplicates()['count'] == 1

    def test_get_duplicate_rows_uses_the_mask(self):
        df = _xfield_df(states=["Closed"] * 4)
        df["number"] = ["INC1", "INC1", "INC2", "INC3"]
        analyzer = SanityAnalyzer(df)
        rows = analyzer.get_duplicate_rows()
        assert len(rows) == int(analyzer.duplicate_row_mask().sum())

    def test_no_id_column_returns_none(self):
        df = _xfield_df(states=["Closed"] * 3).drop(columns=["number"])
        analyzer = SanityAnalyzer(df)
        if analyzer.id_column is None:      # nothing else scored as an ID
            assert analyzer.duplicate_row_mask() is None
            assert analyzer.get_duplicate_rows() is None
            assert analyzer.id_occurrence_counts() is None

    def test_occurrence_counts_are_per_row(self):
        df = _xfield_df(states=["Closed"] * 4)
        df["number"] = ["INC1", "INC1", "INC1", "INC2"]
        counts = SanityAnalyzer(df).id_occurrence_counts()
        assert list(counts) == [3, 3, 3, 1]

    def test_mask_is_cached(self):
        analyzer = SanityAnalyzer(_xfield_df(states=["Closed"] * 3))
        assert analyzer.duplicate_row_mask() is analyzer.duplicate_row_mask()


class TestDescriptionRowScores:
    """
    The per-row tiers must be the same arrays the aggregate is built from, and
    getting them must not pay for the repetition clustering.
    """

    def test_tiers_agree_with_the_score_distribution(self, df_with_description_issues):
        analyzer = SanityAnalyzer(df_with_description_issues)
        aggregate = analyzer.check_description_quality()
        per_row = analyzer.get_description_row_scores()

        for info in aggregate['columns']:
            tiers = per_row[info['column']]['tier']
            for tier, count in info['score_distribution'].items():
                assert int((tiers == tier).sum()) == count

    def test_does_not_run_the_repetition_clustering(self, df_with_description_issues, mocker):
        """That clustering is ~3.3s of the 3.3s cost at 300k rows."""
        analyzer = SanityAnalyzer(df_with_description_issues)
        spy = mocker.spy(analyzer, "_find_repeated_issues")
        analyzer.get_description_row_scores()
        assert spy.call_count == 0

    def test_scores_are_aligned_to_the_frame_index(self, df_with_description_issues):
        df = df_with_description_issues.copy()
        df.index = range(100, 100 + len(df))
        per_row = SanityAnalyzer(df).get_description_row_scores()
        for entry in per_row.values():
            assert list(entry['score'].index) == list(df.index)
            assert list(entry['tier'].index) == list(df.index)

    def test_every_description_column_is_covered(self, standard_df):
        analyzer = SanityAnalyzer(standard_df)
        expected = {col for _cat, col in analyzer.description_columns}
        assert set(analyzer.get_description_row_scores()) == expected

    def test_no_description_column_yields_nothing(self):
        df = _xfield_df(states=["Closed"] * 3).drop(
            columns=["short_description", "description"])
        assert SanityAnalyzer(df).get_description_row_scores() == {}

    def test_thresholds_are_honoured(self):
        df = _xfield_df(states=["Closed"] * 3)
        df["description"] = ["a" * 15] * 3
        analyzer = SanityAnalyzer(df)
        lenient = analyzer.get_description_row_scores(min_length=10)
        strict = analyzer.get_description_row_scores(min_length=50)
        assert set(lenient["description"]['tier']) != set(strict["description"]['tier'])

    def test_result_is_cached_per_threshold(self, df_with_description_issues):
        analyzer = SanityAnalyzer(df_with_description_issues)
        assert (analyzer.get_description_row_scores()
                is analyzer.get_description_row_scores())
        assert (analyzer.get_description_row_scores(min_length=99)
                is not analyzer.get_description_row_scores())
