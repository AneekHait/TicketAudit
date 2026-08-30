"""
Tests for core/language.py — LanguageChecker.
"""
import pytest
import pandas as pd
from unittest.mock import MagicMock, patch
from core.language import LanguageChecker


@pytest.fixture
def checker():
    return LanguageChecker()


# ──────────────────────────────────────────────────────────────
# Text cleaning
# ──────────────────────────────────────────────────────────────

def _clean_naively(series):
    """
    The pre-optimisation implementation: every pattern over every row.
    Kept as the reference definition of correct output — clean_series now
    de-duplicates first, and must not change what it produces.
    """
    cleaned = series.astype(str)
    for pattern, replacement in LanguageChecker._CLEAN_PATTERNS:
        cleaned = cleaned.str.replace(pattern, replacement, regex=True)
    return cleaned.str.strip()


class TestCleanSeries:
    """
    clean_series runs 14 regex passes, which at one re.sub per row per pattern
    dominated Description Quality (~3.2M calls, ~5s on a 100k-row file). It now
    cleans the distinct values and maps back — safe only because cleaning is a
    pure function of the string, so these tests pin the output.
    """

    @pytest.mark.parametrize("values", [
        pytest.param(["Password reset for INC0001234 user@corp.com"] * 20
                     + ["SAP dump on 2024-01-15 see https://kb/a/123"] * 20,
                     id="repetitive"),
        pytest.param([f"Issue {i} for CHG{i:06d}" for i in range(50)], id="all-unique"),
        pytest.param(["INC001 reset", None, "SAP error 42", None], id="with-nulls"),
        pytest.param(["", "   ", "INC1", ""], id="empty-strings"),
        pytest.param(["Password reset INC0001"], id="single-row"),
        pytest.param(["密码重置 INC001", "Passwort zurücksetzen 2024",
                      "密码重置 INC001"], id="non-latin"),
        pytest.param(["Printer PRT12AB at C:\\spool\\x.tmp"] * 3
                     + ["[NETWORK] VPN drops /var/log/vpn.log"] * 3, id="paths-and-tags"),
    ])
    def test_matches_the_naive_implementation(self, values):
        series = pd.Series(values)
        assert list(LanguageChecker.clean_series(series)) == list(_clean_naively(series))

    def test_empty_series(self):
        out = LanguageChecker.clean_series(pd.Series([], dtype=object))
        assert len(out) == 0

    def test_index_is_preserved(self):
        """Callers use the index to map findings back to rows."""
        series = pd.Series(["INC1 alpha", "INC2 beta", "INC1 alpha"], index=[10, 20, 30])
        out = LanguageChecker.clean_series(series)
        assert list(out.index) == [10, 20, 30]

    def test_duplicate_rows_clean_identically(self):
        """The recurring-issue check depends on equal input giving equal output."""
        series = pd.Series(["Reset password for INC0001", "Reset password for INC0002"])
        out = LanguageChecker.clean_series(series)
        assert out.iloc[0] == out.iloc[1], "ticket IDs should normalise away"

    def test_agrees_with_the_per_cell_helper(self):
        values = ["Password reset INC0001234", "SAP dump 2024-01-15", "Password reset INC0001234"]
        series = pd.Series(values)
        assert list(LanguageChecker.clean_series(series)) == \
            [LanguageChecker.clean_text_for_detection(v) for v in values]


# ──────────────────────────────────────────────────────────────
# Character-based detection (tier 1)
# ──────────────────────────────────────────────────────────────

class TestCharacterDetection:
    def test_cyrillic_text_detected_as_russian(self, checker):
        result = checker.detect_language("Привет мир это текст", mode='rules_only')
        assert result is not None
        lang, conf = result
        assert lang == "Russian"
        assert conf == 1.0

    def test_cjk_text_detected_as_chinese(self, checker):
        result = checker.detect_language("你好世界这是中文文本", mode='rules_only')
        assert result is not None
        lang, conf = result
        assert lang in ("Chinese", "Japanese")  # CJK overlaps
        assert conf == 1.0

    def test_arabic_text_detected(self, checker):
        result = checker.detect_language("مرحبا بالعالم هذا نص عربي", mode='rules_only')
        assert result is not None
        lang, _ = result
        assert lang == "Arabic"

    def test_plain_english_returns_none(self, checker):
        result = checker.detect_language("This is a normal English sentence about a computer issue.", mode='rules_only')
        assert result is None


# ──────────────────────────────────────────────────────────────
# Keyword-based detection (tier 2)
# ──────────────────────────────────────────────────────────────

class TestKeywordDetection:
    def test_french_keywords_detected(self, checker):
        text = "bonjour comment puis-je vous aider avec ce probleme"
        result = checker.detect_language(text, mode='rules_only')
        assert result is not None
        lang, _ = result
        assert lang == "French"

    def test_spanish_keywords_detected(self, checker):
        text = "hola buenos dias como puedo ayudarle con este problema del sistema"
        result = checker.detect_language(text, mode='rules_only')
        assert result is not None
        lang, _ = result
        assert lang == "Spanish"


# ──────────────────────────────────────────────────────────────
# Detection cache
# ──────────────────────────────────────────────────────────────

class TestDetectionCache:
    def test_repeated_text_uses_cache(self, checker):
        text = "Привет это тест"
        checker.detect_language(text)
        initial_cache_size = len(checker._detection_cache)
        checker.detect_language(text)  # should hit cache, not grow
        assert len(checker._detection_cache) == initial_cache_size

    def test_different_texts_populate_cache(self, checker):
        checker.detect_language("Привет один")
        checker.detect_language("Привет два")
        assert len(checker._detection_cache) >= 2


# ──────────────────────────────────────────────────────────────
# detect_language_for_row
# ──────────────────────────────────────────────────────────────

class TestDetectLanguageForRow:
    def test_returns_string(self, checker):
        result = checker.detect_language_for_row("Hello this is English text")
        assert isinstance(result, str)

    def test_non_english_row_not_english(self, checker):
        result = checker.detect_language_for_row("Привет мир это текст на русском языке")
        assert result != "English"

    def test_empty_string_returns_english(self, checker):
        result = checker.detect_language_for_row("")
        assert result == "English"


# ──────────────────────────────────────────────────────────────
# get_language_columns (vectorized, no iterrows)
# ──────────────────────────────────────────────────────────────

class TestGetLanguageColumns:
    def test_returns_dataframe_with_language_col(self, checker):
        df = pd.DataFrame({
            "short_description": ["Login issue", "Привет мир это текст на русском", "Network down"]
        })
        result = checker.get_language_columns(df, ["short_description"])
        assert "short_description_Language" in result.columns
        assert len(result) == 3

    def test_empty_values_produce_empty_string(self, checker):
        df = pd.DataFrame({"desc": [None, "", "   "]})
        result = checker.get_language_columns(df, ["desc"])
        for val in result["desc_Language"]:
            assert val == ""

    def test_missing_column_skipped_gracefully(self, checker):
        df = pd.DataFrame({"other": ["text"]})
        result = checker.get_language_columns(df, ["nonexistent_column"])
        # Should return a df with only _original_index
        assert "nonexistent_column_Language" not in result.columns


# ──────────────────────────────────────────────────────────────
# Detection mode is reported honestly
# ──────────────────────────────────────────────────────────────

@pytest.fixture
def multilingual_df():
    return pd.DataFrame({"desc": [
        "Le serveur ne repond plus depuis ce matin",       # French keywords
        "Der Drucker funktioniert nicht mehr richtig",      # German, needs Lingua
        "The printer is not working at all today",          # English
        "Привет мир это русский текст здесь",              # Cyrillic characters
        "tiny",                                            # under min_length
    ]})


def _no_lingua():
    """A checker with Lingua forced unavailable, without uninstalling it."""
    checker = LanguageChecker()
    checker._lingua_available = False
    return checker


class TestModeIsReportedAccurately:
    """
    'using_lingua' used to be is_lingua_available() - whether Lingua was
    *installed*, not whether it *ran*. The UI rendered that as
    "Detection: Lingua + rules" on every run, so selecting Lingua-only or
    Rules-only still claimed both had been used.
    """

    def test_rules_only_does_not_claim_lingua(self, checker, multilingual_df):
        result = checker.check_column(multilingual_df, "desc", mode="rules_only")
        assert result['lingua_available'] is True, "fixture needs Lingua installed"
        assert result['using_lingua'] is False, \
            "rules_only must not report Lingua as used just because it is installed"
        assert "Lingua" not in result['mode_label'] or "no Lingua" in result['mode_label']

    def test_lingua_only_does_not_claim_rules(self, checker, multilingual_df):
        result = checker.check_column(multilingual_df, "desc", mode="lingua_only")
        assert result['rules_used'] is False
        assert result['using_lingua'] is True
        assert result['mode_label'] == "Lingua only"

    def test_all_mode_reports_both(self, checker, multilingual_df):
        result = checker.check_column(multilingual_df, "desc", mode="all")
        assert result['rules_used'] is True
        assert result['using_lingua'] is True

    def test_mode_is_echoed_back(self, checker, multilingual_df):
        for mode in ("all", "lingua_only", "rules_only"):
            result = checker.check_column(multilingual_df, "desc", mode=mode)
            assert result['mode'] == mode

    def test_mode_actually_changes_the_result(self, checker, multilingual_df):
        """If the modes gave identical answers the mislabel would be cosmetic."""
        rules = checker.check_column(multilingual_df, "desc", mode="rules_only")
        every = checker.check_column(multilingual_df, "desc", mode="all")
        assert rules['non_english_count'] < every['non_english_count']


class TestCheckThatCannotRunIsNotAPass:
    """
    Lingua-only with Lingua missing disables every tier, so detection returns
    None for every row. That used to render as "all entries appear to be in
    English" on a column full of French and German.
    """

    def test_lingua_only_without_lingua_warns(self, multilingual_df):
        result = _no_lingua().check_column(multilingual_df, "desc", mode="lingua_only")
        assert 'warning' in result, \
            "a check that could not run must not report as a clean pass"
        assert result['non_english_count'] == 0

    def test_the_warning_names_the_cause_and_the_fix(self, multilingual_df):
        warning = _no_lingua().check_column(
            multilingual_df, "desc", mode="lingua_only")['warning']
        assert "not installed" in warning.lower()
        assert "rules_only" in warning or "all" in warning

    def test_rules_only_without_lingua_still_works(self, multilingual_df):
        """A missing model is irrelevant when the model was not requested."""
        result = _no_lingua().check_column(multilingual_df, "desc", mode="rules_only")
        assert 'warning' not in result
        assert result['non_english_count'] > 0

    def test_all_mode_without_lingua_still_works(self, multilingual_df):
        result = _no_lingua().check_column(multilingual_df, "desc", mode="all")
        assert 'warning' not in result
        assert result['non_english_count'] > 0
        assert result['using_lingua'] is False


class TestDetectionProvenance:
    """Each detection reports which tier found it."""

    _VALID = {LanguageChecker.METHOD_CHARACTERS,
              LanguageChecker.METHOD_KEYWORDS,
              LanguageChecker.METHOD_LINGUA}

    def test_detect_language_still_returns_a_pair(self, checker):
        """The public 2-tuple contract must not change."""
        result = checker.detect_language("Привет мир это текст", mode='rules_only')
        assert result is not None
        lang, conf = result
        assert lang == "Russian" and conf == 1.0

    def test_detailed_adds_the_method(self, checker):
        lang, conf, method = checker.detect_language_detailed(
            "Привет мир это текст", mode='rules_only')
        assert lang == "Russian"
        assert method == LanguageChecker.METHOD_CHARACTERS

    def test_keyword_hit_is_labelled_as_keywords(self, checker):
        detailed = checker.detect_language_detailed(
            "bonjour comment puis-je vous aider avec ce probleme", mode='rules_only')
        assert detailed is not None
        assert detailed[2] == LanguageChecker.METHOD_KEYWORDS

    def test_samples_carry_a_valid_method(self, checker, multilingual_df):
        result = checker.check_column(multilingual_df, "desc", mode="all")
        assert result['samples']
        for sample in result['samples']:
            assert len(sample) == 5, "samples are (idx, text, lang, conf, method)"
            assert sample[4] in self._VALID

    def test_method_counts_sum_to_detections(self, checker, multilingual_df):
        result = checker.check_column(multilingual_df, "desc", mode="all")
        assert sum(result['methods'].values()) == result['non_english_count']

    def test_disabled_tiers_report_no_detections(self, checker, multilingual_df):
        """Lingua-only must not attribute anything to the rule tiers."""
        result = checker.check_column(multilingual_df, "desc", mode="lingua_only")
        assert LanguageChecker.METHOD_CHARACTERS not in result['methods']
        assert LanguageChecker.METHOD_KEYWORDS not in result['methods']

    def test_lingua_reports_no_detections_in_rules_only(self, checker, multilingual_df):
        result = checker.check_column(multilingual_df, "desc", mode="rules_only")
        assert LanguageChecker.METHOD_LINGUA not in result['methods']


class TestPerRowResults:
    """
    check_column decides a language for every analysed row and then returned
    only ten samples. Keeping the per-row answer costs nothing and is the
    difference between a report that can show a language per ticket and one that
    would have to redo the slowest operation in the app.
    """

    def test_row_languages_keys_only_non_english_rows(self, checker, multilingual_df):
        result = checker.check_column(multilingual_df, "desc", mode="all")
        assert result['row_languages']
        assert len(result['row_languages']) == result['non_english_count']

    def test_row_languages_are_indexed_by_frame_label(self, checker):
        df = pd.DataFrame({"desc": [
            "Le serveur ne repond plus depuis ce matin merci",
            "The printer is offline and refusing all print jobs",
        ]}, index=[10, 20])
        result = checker.check_column(df, "desc", mode="all")
        assert set(result['row_languages']) <= {10, 20}

    def test_no_text_is_retained(self, checker, multilingual_df):
        """Only the verdict, so the dict stays proportional to the finding."""
        result = checker.check_column(multilingual_df, "desc", mode="all")
        for value in result['row_languages'].values():
            assert isinstance(value, str)
            assert " " not in value.strip(), "a language name, not a sentence"

    def test_analyzed_mask_matches_the_length_filter(self, checker, multilingual_df):
        result = checker.check_column(multilingual_df, "desc", mode="all",
                                      min_length=10)
        mask = result['analyzed_mask']
        assert len(mask) == len(multilingual_df)
        assert int(sum(mask)) == result['total_analyzed']

    def test_the_short_row_is_excluded_from_the_mask(self, checker, multilingual_df):
        result = checker.check_column(multilingual_df, "desc", mode="all",
                                      min_length=10)
        # the fixture's last row is "tiny"
        assert bool(result['analyzed_mask'][-1]) is False

    def test_every_detected_row_was_also_analysed(self, checker, multilingual_df):
        result = checker.check_column(multilingual_df, "desc", mode="all")
        mask = pd.Series(result['analyzed_mask'], index=multilingual_df.index)
        for index in result['row_languages']:
            assert bool(mask[index]) is True


class TestDetectLanguageSeries:
    """
    Detects over unique() and maps back. The detection cache is capped at 10,000
    entries, so a row-by-row pass over a column with more distinct texts than
    that stops getting cache hits entirely and pays the full model cost per row.
    """

    def test_results_match_a_row_by_row_pass(self, checker):
        texts = pd.Series([
            "Le serveur ne repond plus depuis ce matin merci",
            "The printer is offline and refusing all print jobs",
            "Der Drucker funktioniert nicht mehr richtig heute",
            "Le serveur ne repond plus depuis ce matin merci",
        ])
        deduped = checker.detect_language_series(texts, mode="all")
        reference = texts.map(
            lambda t: checker.detect_language_for_row(t, 0.75, "all"))
        assert list(deduped) == list(reference)

    def test_it_detects_once_per_distinct_text(self, checker):
        calls = []
        original = checker._detect_language_uncached

        def counting(*args, **kwargs):
            calls.append(args[0])
            return original(*args, **kwargs)

        checker._detect_language_uncached = counting
        texts = pd.Series(["Le serveur ne repond plus ce matin merci"] * 50
                          + ["The printer is offline refusing jobs"] * 50)
        checker.detect_language_series(texts, mode="all")
        assert len(calls) == 2, "one detection per distinct text, not per row"

    def test_blank_rows_are_empty_not_english(self, checker):
        """Nothing was judged, so nothing is claimed."""
        texts = pd.Series(["", "   ", None, "The printer is offline today"])
        result = checker.detect_language_series(texts, mode="all")
        assert list(result[:3]) == ["", "", ""]

    def test_the_index_is_preserved(self, checker):
        texts = pd.Series(["The printer is offline and refusing jobs"],
                          index=[42])
        assert list(checker.detect_language_series(texts, mode="all").index) == [42]

    def test_add_language_columns_uses_it(self, checker):
        df = pd.DataFrame({"desc": [
            "Le serveur ne repond plus depuis ce matin merci",
            "The printer is offline and refusing all print jobs",
        ]})
        out = checker.add_language_columns_to_df(df, ["desc"])
        assert "desc_Language" in out.columns
        assert out["desc_Language"].iloc[0] == "French"


class TestScopeReporting:
    def test_short_rows_are_counted_as_skipped(self, checker, multilingual_df):
        result = checker.check_column(multilingual_df, "desc", mode="all", min_length=10)
        assert result['skipped_short'] == 1, "the 'tiny' row is under 10 characters"
        assert result['total_analyzed'] + result['skipped_short'] == result['total_original']

    def test_min_length_is_echoed_back(self, checker, multilingual_df):
        result = checker.check_column(multilingual_df, "desc", min_length=25)
        assert result['min_length'] == 25

    def test_all_short_column_warns_with_the_threshold(self, checker):
        df = pd.DataFrame({"desc": ["a", "bb", "ccc"]})
        result = checker.check_column(df, "desc", min_length=50)
        assert 'warning' in result
        assert "50" in result['warning']

    def test_format_result_handles_the_method_field(self, checker, multilingual_df):
        """format_result unpacks samples; a 4-tuple unpack would raise here."""
        result = checker.check_column(multilingual_df, "desc", mode="all")
        lines = checker.format_result(result)
        assert any("Sample" in line for line in lines)
