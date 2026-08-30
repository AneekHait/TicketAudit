"""
Report generation module - creates detailed reports from analysis results.
"""
import datetime
import os
from typing import Dict, List, Any, Optional, Tuple
import numpy as np
import pandas as pd

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

from config import COLUMN_KEYWORDS
from core.analyzer import SanityAnalyzer
from core.language import LanguageChecker

# How each detection tier is described in a report. The reliability note matters:
# a script match is definitive, a keyword match is a heuristic that can misfire
# on short text, and only Lingua weighs the whole string - so a reader deciding
# whether to trust a finding needs to know which one produced it.
_METHOD_REPORT_LABELS = {
    LanguageChecker.METHOD_CHARACTERS: 'non-English script (definitive)',
    LanguageChecker.METHOD_KEYWORDS: 'keyword patterns (heuristic)',
    LanguageChecker.METHOD_LINGUA: 'Lingua language model',
}

# --- Findings register ordering -------------------------------------------
# Failures first, then coverage gaps, then judgement calls, then passes. "Not
# checked" deliberately outranks "Review" and "Pass": an unexamined dimension is
# easier to miss than a failing one and matters more than a clean result.
_RESULT_ORDER = {"Fail": 0, "Not checked": 1, "Review": 2, "Pass": 3}
_SEVERITY_ORDER = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3, "-": 4}
_SEVERITY_NONE = "-"

# --- Distribution / cube caps ---------------------------------------------
# A column with more distinct values than this is a free-text field, not a
# dimension: summarising it would swamp the sheet and pivoting it is meaningless.
MAX_DISTINCT_TO_SUMMARISE = 50
TOP_VALUES_PER_COLUMN = 20
MAX_CUBE_DIMENSIONS = 4
MAX_CUBE_DISTINCT = 50
MAX_CUBE_ROWS = 5000

# Past roughly seven classes adjacent categories stop being tellable apart, so a
# chart folds its tail into one "Other" slice rather than growing more colours.
MAX_CHART_CATEGORIES = 6

# Detail-sheet column titles for each per-row flag. Keyed by the analyzer's own
# flag constants so a rename cannot silently desynchronise the two.
_FLAG_LABELS = {
    SanityAnalyzer.FLAG_CLOSED_BEFORE_CREATED: 'Resolved before created',
    SanityAnalyzer.FLAG_FUTURE_CREATED: 'Created in future',
    SanityAnalyzer.FLAG_FUTURE_CLOSED: 'Resolved in future',
    SanityAnalyzer.FLAG_CLOSED_NO_DATE: 'Closed with no resolution date',
    SanityAnalyzer.FLAG_OPEN_WITH_DATE: 'Open with resolution date',
    SanityAnalyzer.FLAG_OPEN_MISSING_PRIORITY: 'Open with no priority',
}


class ReportGenerator:
    """
    Generates detailed sanity check reports in text and Excel formats.
    Optimized with:
    - Cached language checker instance
    - Batch Excel operations
    - Pre-computed analysis results
    """
    
    # Class-level shared language checker (avoid re-initialization)
    _shared_lang_checker = None
    
    @classmethod
    def get_lang_checker(cls):
        """Get shared language checker instance."""
        if cls._shared_lang_checker is None:
            cls._shared_lang_checker = LanguageChecker()
        return cls._shared_lang_checker
    
    def __init__(self, df: pd.DataFrame, analyzer: 'SanityAnalyzer',
                 filepath: str = None, null_threshold: float = 20.0,
                 lang_columns: List[str] = None,
                 desc_min_length: int = 20, desc_max_length: int = 5000,
                 lang_results: List[Dict[str, Any]] = None,
                 lang_mode: str = 'all', native_pivots: bool = True):
        self.df = df
        self.analyzer = analyzer
        self.filepath = filepath
        self.null_threshold = null_threshold
        self.lang_columns = lang_columns or []
        self.desc_min_length = desc_min_length
        self.desc_max_length = desc_max_length
        # Already-computed check_column() results, normally the ones the user
        # ran in the Language Check view. Preferred over recomputing: language
        # detection costs ~1.2ms per *distinct* text with Lingua enabled
        # (measured: 26s for 50k rows, ~21k distinct), which would stall an
        # export for minutes on a large file. Reusing them also guarantees the
        # report agrees with what was on screen.
        self.lang_results = lang_results or []
        self.lang_mode = lang_mode
        self.native_pivots = native_pivots
        self.lang_checker = self.get_lang_checker()  # Use shared instance

        # Filled in by export_to_excel. Anything the workbook could not contain
        # is recorded here rather than left for the reader to notice - a missing
        # sheet is indistinguishable from a clean one.
        self.export_disclosures: List[str] = []
        self.detail_path: Optional[str] = None
        self.detail_rows_written = 0
        self.detail_truncated = False
        # Columns the pivot cube's caps excluded, and why. Populated by
        # _build_pivot_cube_df and surfaced on the Overview sheet.
        self.pivot_cube_notes: List[str] = []

    def _language_results(self) -> List[Dict[str, Any]]:
        """
        Language findings for the report, in order of preference:
        results handed in by the caller, then an explicit lang_columns request
        (opt-in, because it recomputes), then nothing.
        """
        if self.lang_results:
            return self.lang_results
        if self.lang_columns:
            return [self.lang_checker.check_column(self.df, col, mode=self.lang_mode)
                    for col in self.lang_columns if col in self.df.columns]
        return []
    
    def generate_text_report(self) -> str:
        """Generate a detailed text-based sanity check report efficiently."""
        if self.df is None:
            return "No data loaded."
        
        # Use list for efficient string building
        parts = [
            "=" * 60,
            "           SANITY CHECK REPORT",
            "=" * 60,
            f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"File: {os.path.basename(self.filepath) if self.filepath else 'N/A'}",
            "",
            "📊 DATA OVERVIEW",
            "-" * 40,
            f"Total Rows: {self.analyzer.total_rows:,}",
            f"Total Columns: {self.analyzer.total_columns}",
            ""
        ]
        
        # Extend with sections
        parts.extend(self._section_column_check())
        parts.extend(self._section_null_analysis())
        parts.extend(self._section_date_formats())
        parts.extend(self._section_date_logic())
        parts.extend(self._section_cross_field_logic())
        parts.extend(self._section_duplicates())
        parts.extend(self._section_language())
        parts.extend(self._section_distributions())
        parts.extend(self._section_summary())
        
        return "\n".join(parts)
    
    def _section_column_check(self) -> List[str]:
        """Generate column check section."""
        lines = ["📋 COLUMN CHECK", "-" * 40]
        
        results = self.analyzer.check_columns()
        found_lines = []
        missing_lines = []
        
        for req_col, data in results.items():
            if data['found']:
                found_lines.append(f"✅ {req_col}: {', '.join(data['matched_columns'])}")
            else:
                missing_lines.append(f"❌ {req_col}: MISSING")
        
        lines.extend(found_lines)
        lines.extend(missing_lines)
        lines.append("")
        
        return lines
    
    def _section_null_analysis(self) -> List[str]:
        """Generate null analysis section."""
        lines = ["🔍 NULL ANALYSIS", "-" * 40]
        
        results = self.analyzer.check_nulls(self.null_threshold)
        high_nulls = []
        low_nulls = []
        
        for col, data in results.items():
            if data['above_threshold']:
                high_nulls.append(
                    f"❌ {col}: {data['count']:,} nulls ({data['percentage']:.1f}%) - ABOVE THRESHOLD"
                )
            elif data['count'] > 0:
                low_nulls.append(
                    f"⚠️ {col}: {data['count']:,} nulls ({data['percentage']:.1f}%)"
                )
        
        if high_nulls:
            lines.append(f"Columns above {self.null_threshold}% threshold:")
            for line in high_nulls:
                lines.append(f"  {line}")
        else:
            lines.append(f"✅ No columns exceed {self.null_threshold}% null threshold.")
        
        if low_nulls:
            lines.append(f"\nColumns with some nulls:")
            for line in low_nulls[:10]:
                lines.append(f"  {line}")
            if len(low_nulls) > 10:
                lines.append(f"  ... and {len(low_nulls) - 10} more columns with nulls")
        
        lines.append("")
        return lines
    
    def _section_date_logic(self) -> List[str]:
        """Generate date logic section."""
        lines = ["📅 DATE LOGIC CHECKS", "-" * 40]
        
        result = self.analyzer.check_date_logic()
        
        if not result['has_date_columns']:
            lines.append("⚠️ Could not find Created/Closed columns for date checks")
        else:
            if result['closed_before_created'] > 0:
                lines.append(
                    f"❌ {result['closed_before_created']:,} rows where "
                    f"{result['closed_col']} < {result['created_col']}"
                )
            else:
                lines.append(
                    f"✅ No rows where {result['closed_col']} < {result['created_col']}"
                )
            
            if result['future_created'] > 0:
                lines.append(f"❌ {result['future_created']:,} rows with future {result['created_col']}")
            if result['future_closed'] > 0:
                lines.append(f"❌ {result['future_closed']:,} rows with future {result['closed_col']}")
            if result['future_created'] == 0 and result['future_closed'] == 0:
                lines.append("✅ No future dates found")
        
        lines.append("")
        return lines
    
    def _section_cross_field_logic(self) -> List[str]:
        """State/resolution consistency and state vocabulary."""
        lines = ["🔗 CROSS-FIELD CHECKS", "-" * 40]

        def tickets(count: int) -> str:
            return f"{count:,} ticket" + ("" if count == 1 else "s")

        def rows_(count: int) -> str:
            return f"{count:,} row" + ("" if count == 1 else "s")

        result = self.analyzer.check_cross_field_logic()
        if not result['has_state']:
            lines.append("⚠️ Could not find a State/Status column for cross-field checks")
            lines.append("")
            return lines

        state_col = result['state_col']
        lines.append(
            f"State column: {state_col} "
            f"(open: {result['open_count']:,}, closed: {result['closed_count']:,})"
        )

        if result['has_closed_date']:
            closed_col = result['closed_col']
            if result['closed_no_date'] > 0:
                have = "has" if result['closed_no_date'] == 1 else "have"
                lines.append(
                    f"❌ {tickets(result['closed_no_date'])} marked closed/resolved "
                    f"{have} no {closed_col}"
                )
            else:
                lines.append(f"✅ All closed/resolved tickets have a {closed_col}")
            if result['open_with_date'] > 0:
                has = "has" if result['open_with_date'] == 1 else "have"
                lines.append(
                    f"❌ {tickets(result['open_with_date'])} still open already "
                    f"{has} a {closed_col}"
                )
        else:
            lines.append("⚠️ No Closed/Resolved column — resolution cross-check skipped")

        if result['has_priority']:
            if result['open_missing_priority'] > 0:
                have = "has" if result['open_missing_priority'] == 1 else "have"
                lines.append(
                    f"❌ {tickets(result['open_missing_priority'])} still open "
                    f"{have} no {result['priority_col']}"
                )
            else:
                lines.append("✅ All open tickets have a priority")
        else:
            lines.append("⚠️ No Priority column — open-ticket priority check skipped")

        if result['unclassified_count'] > 0:
            has = "has" if result['unclassified_count'] == 1 else "have"
            lines.append(
                f"⚠️ {rows_(result['unclassified_count'])} {has} a {state_col} value "
                f"that is neither open nor closed (excluded from the counts above)"
            )

        vocab = self.analyzer.check_state_values()
        if vocab['has_state']:
            suspects = vocab['suspect_values']
            if suspects:
                value_word = "value" if len(suspects) == 1 else "values"
                lines.append(
                    f"❌ {len(suspects)} non-standard {state_col} {value_word} "
                    f"covering {rows_(vocab['total_suspect_rows'])}:"
                )
                for entry in suspects[:10]:
                    lines.append(
                        f"    {entry['value']} — {rows_(entry['count'])} "
                        f"({entry['reason']})"
                    )
                if len(suspects) > 10:
                    lines.append(f"    ... and {len(suspects) - 10} more")
            else:
                lines.append(
                    f"✅ {state_col} vocabulary is consistent "
                    f"({vocab['total_distinct']} distinct values)"
                )

        lines.append("")
        return lines

    def _section_duplicates(self) -> List[str]:
        """Generate duplicates section."""
        lines = ["🔁 DUPLICATE CHECK", "-" * 40]
        
        result = self.analyzer.check_duplicates()
        
        if result.get('error'):
            lines.append(f"⚠️ {result['error']}")
        elif result['count'] > 0:
            lines.append(f"❌ {result['count']:,} duplicate rows found (based on {result['id_col']})")
            dup_ids = list(result['duplicate_ids'].keys())[:10]
            lines.append(f"   Duplicate IDs: {', '.join(dup_ids)}")
            if len(result['duplicate_ids']) > 10:
                lines.append(f"   ... and {len(result['duplicate_ids']) - 10} more duplicate IDs")
        else:
            lines.append(f"✅ No duplicates found (based on {result['id_col']})")
        
        lines.append("")
        return lines
    
    def _section_date_formats(self) -> List[str]:
        """
        How each date column's field order was decided.

        Stated even when the answer is confident, because "we read these as
        DD/MM" is the assumption every date-derived number in the report rests
        on, and it is invisible otherwise.
        """
        report = self.analyzer.check_date_formats()
        columns = report.get('columns', [])
        if not columns:
            return []

        lines = ["📅 DATE FORMAT", "-" * 40]
        for info in columns:
            column = info['column']
            if info.get('conflicting'):
                lines.append(
                    f"❌ {column}: mixed field orders in one column - "
                    f"{info['day_first_rows']} rows can only be DD/MM and "
                    f"{info['month_first_rows']} can only be MM/DD.")
                lines.append(
                    "   No single reading is correct, so some dates in this "
                    "column are wrong whatever is chosen. Fix it at source.")
            elif info.get('needs_confirmation'):
                lines.append(
                    f"⚠️ {column}: {info['ambiguous']} of {info['total']} values "
                    f"could be read as either DD/MM or MM/DD "
                    f"(e.g. {', '.join(info['examples'])}).")
                lines.append(
                    f"   Read here as "
                    f"{'DD/MM' if self.analyzer._effective_dayfirst(info) else 'MM/DD'}"
                    f"; the other reading would change "
                    f"{info['rows_that_would_change']:,} rows, and every "
                    f"resolution time and trend derived from them.")
            else:
                order = {'day_first': 'DD/MM', 'month_first': 'MM/DD'}.get(
                    info['order'], info['order'])
                lines.append(
                    f"✅ {column}: read as {order}, proven by values that can "
                    f"only be read one way.")
        lines.append("")
        return lines

    def _section_language(self) -> List[str]:
        """
        Language findings, including which detection methods actually ran.

        An absent language check is stated, not omitted: a report that simply
        has no language section reads as "nothing found" when the truth is
        "never looked".
        """
        lines = ["🌐 LANGUAGE CHECK", "-" * 40]
        results = self._language_results()

        if not results:
            lines.append("⚠️ Not checked — no language analysis has been run.")
            lines.append("   Open Language Check, pick a text column and click Run Check,")
            lines.append("   then generate this report again to include the results.")
            lines.append("")
            return lines

        for result in results:
            col = result.get('column', '?')

            if 'error' in result:
                lines.append(f"⚠️ {col}: {result['error']}")
                lines.append("")
                continue

            # A mode that could not run is not a clean column.
            if 'warning' in result:
                lines.append(f"⚠️ {col}: NOT CHECKED — {result['warning']}")
                lines.append("")
                continue

            analyzed = result.get('total_analyzed', 0)
            skipped = result.get('skipped_short', 0)
            non_english = result.get('non_english_count', 0)
            pct = result.get('percentage', 0)

            if non_english:
                lines.append(
                    f"❌ {col}: {non_english:,} of {analyzed:,} analyzed rows "
                    f"({pct:.1f}%) contain non-English text"
                )
            else:
                lines.append(f"✅ {col}: all {analyzed:,} analyzed rows appear to be English")

            lines.append(f"   Detection method: {result.get('mode_label', result.get('mode', 'n/a'))}")
            if skipped:
                lines.append(
                    f"   Rows skipped as too short (< {result.get('min_length', 0)} chars): {skipped:,}"
                )

            languages = result.get('languages', {})
            if languages:
                lines.append("   Languages found:")
                for lang, count in sorted(languages.items(), key=lambda kv: -kv[1]):
                    share = (count / analyzed * 100) if analyzed else 0.0
                    display = self.lang_checker.LANG_NAMES.get(lang, lang)
                    lines.append(f"     {display}: {count:,} ({share:.1f}% of analyzed)")

            methods = result.get('methods', {})
            if methods:
                lines.append("   How they were detected:")
                for method, count in sorted(methods.items(), key=lambda kv: -kv[1]):
                    lines.append(
                        f"     {_METHOD_REPORT_LABELS.get(method, method)}: {count:,}"
                    )

            lines.extend(self._language_examples(result))
            lines.append("")

        return lines

    def _language_examples(self, result: Dict[str, Any]) -> List[str]:
        """
        A few offending rows, with their ticket IDs.

        Without these the section states a count and leaves the reader no way to
        act on it - "7 rows are in French" is only useful if you can find them.
        """
        samples = result.get('samples') or []
        if not samples:
            return []

        id_col = getattr(self.analyzer, 'id_column', None)
        lines = ["   Examples:"]
        for sample in samples[:5]:
            idx, text, lang, _conf, method = sample
            display = self.lang_checker.LANG_NAMES.get(lang, lang)

            ticket = ""
            if id_col is not None and id_col in self.df.columns:
                try:
                    ticket = f"{self.df.loc[idx, id_col]} — "
                except Exception:
                    # An unusual or duplicated index must not break the report
                    ticket = ""

            snippet = " ".join(str(text).split())
            if len(snippet) > 90:
                snippet = snippet[:90] + "…"
            lines.append(f"     {ticket}{display}: {snippet}")

        remaining = result.get('non_english_count', 0) - len(samples[:5])
        if remaining > 0:
            lines.append(f"     … and {remaining:,} more (full list in the Excel export)")
        return lines

    def _language_findings(self) -> List[str]:
        """
        Language issues phrased for the SUMMARY block.

        The summary is built from analyzer.get_summary(), which has no language
        awareness - so a file whose only defect was 63% French text printed
        "No major issues found" directly under a section reporting it.
        """
        findings = []
        for result in self._language_results():
            col = result.get('column', '?')
            if 'error' in result:
                continue
            if 'warning' in result:
                findings.append(f'Language of "{col}" was NOT checked — {result["warning"]}')
                continue
            count = result.get('non_english_count', 0)
            if count:
                langs = ", ".join(
                    self.lang_checker.LANG_NAMES.get(l, l)
                    for l, _ in sorted(result.get('languages', {}).items(),
                                       key=lambda kv: -kv[1])[:3]
                )
                findings.append(
                    f'{count:,} rows in "{col}" ({result.get("percentage", 0):.1f}% of '
                    f'those analyzed) are not in English ({langs})'
                )
        return findings

    def _section_distributions(self) -> List[str]:
        """Generate key distributions section."""
        lines = ["📊 KEY DISTRIBUTIONS", "-" * 40]
        
        for kw in ["priority", "state", "status"]:
            for col in self.df.columns:
                if kw in col.lower():
                    pivot = self.analyzer.get_pivot(col)
                    lines.append(f"\n{col}:")
                    for item in pivot['values'][:8]:
                        lines.append(f"  {item['value']}: {item['count']:,} ({item['percentage']:.1f}%)")
                    if pivot['total_unique'] > 8:
                        lines.append(f"  ... and {pivot['total_unique'] - 8} more values")
                    break
        
        lines.append("")
        return lines
    
    def _section_summary(self) -> List[str]:
        """Generate summary section."""
        lines = [
            "=" * 60,
            "           SUMMARY",
            "=" * 60
        ]
        
        summary = self.analyzer.get_summary(self.null_threshold)
        # get_summary() covers the analyzer's own checks only; language findings
        # are computed outside it and have to be merged in, or the summary
        # contradicts the section above it.
        issues = list(summary['issues']) + self._language_findings()

        if issues:
            lines.append("⚠️ ISSUES FOUND:")
            for issue in issues:
                lines.append(f"   • {issue}")
        else:
            lines.append("✅ No major issues found!")
        
        lines.append("")
        lines.append("=" * 60)
        
        return lines

    def generate_summary_report(self) -> str:
        """
        A findings summary written to be pasted into a ticket or an email.

        Model-free and deterministic: it reads the analyzer cache plus the same
        checks generate_text_report uses, so it costs what those checks cost and
        nothing more. Plain text throughout - the destination is a mail client,
        which is why there is no HTML and no Markdown here.

        It differs from generate_text_report in what it is for, not in what it
        knows: the text report walks every section in order, this one sorts the
        same findings into critical / warning / passed and derives the actions
        they imply. Anything new belongs in both.

        Language findings come from self._language_results() rather than being
        recomputed, for the reason __init__ documents. Nothing at all is said
        when no check was run: that is not the false clean bill of health "no
        language issues" would be, and emitting "language not checked" to a
        stakeholder for a view they may never open is noise. (The text report
        makes the opposite choice, and deliberately - it is a document about the
        whole file, so a silently absent section there reads as "nothing
        found".)
        """
        if self.df is None:
            return "No data loaded."

        def plural(count: int, singular: str, suffix: str = "s") -> str:
            return f"{count:,} {singular}{'' if count == 1 else suffix}"

        analyzer = self.analyzer
        cache = getattr(analyzer, '_cache', {}) or {}
        null_threshold = self.null_threshold
        lang_results = self._language_results()

        total_rows = cache.get('total_rows', len(self.df))
        total_cols = len(self.df.columns)

        critical: List[str] = []
        warnings: List[str] = []
        ok_notes: List[str] = []
        action_items: List[str] = []

        # -- Required columns ----------------------------------------------
        column_map = cache.get('column_map', {})
        missing_cols = [req for req in COLUMN_KEYWORDS if req not in column_map]
        found_cols   = [req for req in COLUMN_KEYWORDS if req in column_map]

        if missing_cols:
            for col in missing_cols:
                critical.append(
                    f'Required column "{col}" is missing — could not be matched '
                    f'to any column in the file'
                )
                action_items.append(
                    f'Confirm the "{col}" column name (it may be labelled '
                    f'differently in your export)'
                )
        else:
            preview = ", ".join(list(found_cols)[:5])
            if len(found_cols) > 5:
                preview += "…"
            ok_notes.append(
                f'All {len(COLUMN_KEYWORDS)} required columns found ({preview})'
            )

        # -- Null analysis -------------------------------------------------
        null_counts = cache.get('null_counts')
        if null_counts is None:
            null_counts = self.df.isnull().sum()

        high_null, warn_null = [], []
        for col in self.df.columns:
            pct = (null_counts[col] / total_rows * 100) if total_rows > 0 else 0
            if pct > null_threshold:
                high_null.append((col, int(null_counts[col]), pct))
            elif pct > 0:
                warn_null.append((col, int(null_counts[col]), pct))

        for col, count, pct in high_null:
            critical.append(
                f'"{col}" has {pct:.0f}% null values ({count:,} of {total_rows:,} '
                f'rows) — exceeds the {null_threshold:.0f}% threshold'
            )
            action_items.append(
                f'Populate "{col}" for all records before resubmitting'
            )
        if not high_null:
            ok_notes.append(
                f'No columns exceed the {null_threshold:.0f}% null threshold'
            )
        for col, count, pct in warn_null[:3]:
            warnings.append(
                f'"{col}" has {pct:.1f}% null values ({plural(count, "row")})'
            )

        # -- Date logic ----------------------------------------------------
        date_result = None
        if analyzer is not None:
            try:
                date_result = analyzer.check_date_logic()
            except Exception:
                pass

        if date_result and date_result.get('has_date_columns'):
            seq    = date_result.get('closed_before_created', 0)
            fut_c  = date_result.get('future_created', 0)
            fut_cl = date_result.get('future_closed', 0)
            c_col  = date_result.get('created_col', 'opened')
            cl_col = date_result.get('closed_col', 'closed')

            if seq > 0:
                critical.append(
                    f'{plural(seq, "row")} where "{cl_col}" precedes "{c_col}" '
                    f'(impossible date sequence)'
                )
                action_items.append(
                    f'Correct the date sequence for {plural(seq, "record")} '
                    f'where close predates open'
                )
            else:
                ok_notes.append('Date sequence is valid (no close-before-open records)')

            if fut_c + fut_cl > 0:
                fut_total = fut_c + fut_cl
                warnings.append(
                    f'{plural(fut_total, "row")} '
                    f'{"contains" if fut_total == 1 else "contain"} future dates '
                    f'({fut_c:,} in "{c_col}", {fut_cl:,} in "{cl_col}")'
                )
                action_items.append(
                    f'Review and correct {plural(fut_c + fut_cl, "record")} '
                    f'with future dates'
                )

        # -- State / resolution consistency --------------------------------
        xfield = None
        if analyzer is not None:
            try:
                xfield = analyzer.check_cross_field_logic()
            except Exception:
                pass

        if xfield and xfield.get('has_state'):
            state_col = xfield.get('state_col', 'state')
            closed_col = xfield.get('closed_col', 'closed')
            prio_col = xfield.get('priority_col', 'priority')

            closed_no_date = xfield.get('closed_no_date', 0)
            open_with_date = xfield.get('open_with_date', 0)
            no_priority = xfield.get('open_missing_priority', 0)
            unclassified = xfield.get('unclassified_count', 0)

            if closed_no_date > 0:
                critical.append(
                    f'{plural(closed_no_date, "ticket")} marked closed/resolved in '
                    f'"{state_col}" {"has" if closed_no_date == 1 else "have"} no '
                    f'"{closed_col}" value — the resolution date was never recorded'
                )
                action_items.append(
                    f'Backfill "{closed_col}" for {plural(closed_no_date, "closed ticket")}'
                )
            if open_with_date > 0:
                critical.append(
                    f'{plural(open_with_date, "ticket")} still open in "{state_col}" '
                    f'already {"has" if open_with_date == 1 else "have"} a '
                    f'"{closed_col}" value — the state or the date is wrong'
                )
                action_items.append(
                    f'Reconcile "{state_col}" against "{closed_col}" for '
                    f'{plural(open_with_date, "ticket")}'
                )
            if no_priority > 0:
                warnings.append(
                    f'{plural(no_priority, "open ticket")} '
                    f'{"has" if no_priority == 1 else "have"} no "{prio_col}" value '
                    f'— cannot be triaged or SLA-tracked'
                )
                action_items.append(
                    f'Assign "{prio_col}" to {plural(no_priority, "open ticket")}'
                )
            if closed_no_date == 0 and open_with_date == 0 and no_priority == 0:
                ok_notes.append(
                    f'"{state_col}" is consistent with "{closed_col}" '
                    f'(no closed tickets missing a resolution date)'
                )
            if unclassified > 0:
                # Stated as a limit on the check, not as a defect in the data:
                # these rows were excluded from the counts above, so a reader
                # must not take those counts as covering the whole file.
                warnings.append(
                    f'{plural(unclassified, "row")} use a "{state_col}" value that '
                    f'could not be read as either open or closed and '
                    f'{"was" if unclassified == 1 else "were"} excluded from the '
                    f'state checks above'
                )

        # -- State vocabulary ----------------------------------------------
        vocab = None
        if analyzer is not None:
            try:
                vocab = analyzer.check_state_values()
            except Exception:
                pass

        if vocab and vocab.get('has_state'):
            suspects = vocab.get('suspect_values', [])
            state_col = vocab.get('state_col', 'state')
            if suspects:
                listed = "; ".join(
                    f'"{e["value"]}" ({e["count"]:,} — {e["reason"]})'
                    for e in suspects[:5]
                )
                extra = f' and {len(suspects) - 5} more' if len(suspects) > 5 else ''
                warnings.append(
                    f'{plural(len(suspects), "non-standard value")} in "{state_col}" '
                    f'covering {plural(vocab.get("total_suspect_rows", 0), "row")}: '
                    f'{listed}{extra}'
                )
                action_items.append(
                    f'Correct the non-standard "{state_col}" values so reporting '
                    f'groups them with the intended state'
                )
            else:
                ok_notes.append(
                    f'"{state_col}" uses a consistent vocabulary '
                    f'({plural(vocab.get("total_distinct", 0), "distinct value")}, '
                    f'no typos or case variants)'
                )

        # -- Language ------------------------------------------------------
        # Handed over, never recomputed: detection costs ~1.2ms per *distinct*
        # text with Lingua, which would turn an instant summary into a
        # minutes-long wait on a real file - and would report different numbers
        # than the screen if the modes differed.
        #
        # With no results this block emits nothing at all. See the docstring.
        for lang_result in lang_results:
            col = lang_result.get('column', 'text column')

            if 'error' in lang_result:
                continue

            if 'warning' in lang_result:
                warnings.append(
                    f'Language of "{col}" could not be checked — '
                    f'{lang_result["warning"]}'
                )
                action_items.append(
                    f'Re-run the language check on "{col}" with a detection mode '
                    f'that can run, so the column is actually verified'
                )
                continue

            count = lang_result.get('non_english_count', 0)
            analyzed = lang_result.get('total_analyzed', 0)
            method = lang_result.get('mode_label', 'the selected methods')

            if count:
                named = ", ".join(
                    f'{lang} ({n:,})'
                    for lang, n in sorted(lang_result.get('languages', {}).items(),
                                          key=lambda kv: -kv[1])[:4]
                )
                warnings.append(
                    f'{plural(count, "row")} in "{col}" '
                    f'({lang_result.get("percentage", 0):.1f}% of the {analyzed:,} '
                    f'analyzed) are not in English: {named}'
                )
                action_items.append(
                    f'Translate or re-key the non-English "{col}" entries — '
                    f'keyword search and reporting will miss them as they stand'
                )
            else:
                ok_notes.append(
                    f'"{col}" is entirely English across {analyzed:,} analyzed rows '
                    f'(checked using {method})'
                )

        # -- Duplicates ----------------------------------------------------
        dup_result = None
        if analyzer is not None:
            try:
                dup_result = analyzer.check_duplicates()
            except Exception:
                pass

        if dup_result and not dup_result.get('error'):
            dup_count = dup_result.get('count', 0)
            id_col    = dup_result.get('id_col', 'ID')
            if dup_count > 0:
                warnings.append(
                    f'{plural(dup_count, "duplicate record")} found '
                    f'(based on "{id_col}")'
                )
                action_items.append(
                    f'Remove or deduplicate the {plural(dup_count, "record")} '
                    f'with repeated "{id_col}" values'
                )
            else:
                ok_notes.append(f'No duplicate records detected (checked "{id_col}")')

        # -- Compose -------------------------------------------------------
        today = datetime.date.today().strftime('%B %d, %Y')

        if critical:
            n = len(critical)
            subject = (
                f'[ACTION REQUIRED] Data Quality Review — '
                f'{n} Critical Issue{"s" if n > 1 else ""} Found'
            )
        elif warnings:
            n = len(warnings)
            subject = (
                f'[PLEASE REVIEW] Data Quality Review — '
                f'{n} Warning{"s" if n > 1 else ""}'
            )
        else:
            subject = 'Data Quality Review — All Checks Passed'

        lines = [
            f'Subject: {subject}',
            '',
            'Hi Team,',
            '',
            (
                f'Thank you for providing the data file. We have completed an '
                f'automated quality review ({today}) against our ITSM data standards. '
                f'Below is a summary of our findings for {total_rows:,} rows across '
                f'{total_cols} columns.'
            ),
            '',
        ]

        if critical:
            lines.append(f'CRITICAL ISSUES ({len(critical)}):')
            for i, item in enumerate(critical, 1):
                lines.append(f'  {i}. {item}.')
            lines.append('')

        if warnings:
            lines.append(f'WARNINGS ({len(warnings)}):')
            for i, item in enumerate(warnings, 1):
                lines.append(f'  {i}. {item}.')
            lines.append('')

        if ok_notes:
            lines.append('PASSED:')
            for item in ok_notes:
                lines.append(f'  • {item}.')
            lines.append('')

        if action_items:
            lines.append('REQUESTED ACTIONS:')
            for i, item in enumerate(action_items, 1):
                lines.append(f'  {i}. {item}.')
            lines.append('')
            lines.append(
                'Please address the items above and send the corrected file at your '
                'earliest convenience. If you believe any finding is incorrect, do '
                'not hesitate to reach out.'
            )
        else:
            lines.append(
                'The data file has passed all automated quality checks. '
                'No further action is required at this time.'
            )

        lines += [
            '',
            'Best regards,',
            '[Your Name]',
            'Data Quality Team',
            '',
            '─' * 52,
            f'[Generated by TicketAudit Summary Report — {today}]',
        ]

        return '\n'.join(lines)

    def export_to_excel(self, export_path: str, *, progress=None) -> bool:
        """
        Export the consolidated report workbook. Returns True if successful.

        Assembly lives in core/excel_report.py; this method owns the data and
        delegates the layout. Notes about anything the workbook could not
        contain end up in self.export_disclosures for the caller to surface.
        """
        if not OPENPYXL_AVAILABLE:
            raise ImportError("openpyxl is required for Excel export")

        if self.df is None:
            return False

        # Imported here rather than at module scope: excel_report imports this
        # module's constants, so a top-level import would be circular.
        from core.excel_report import build_workbook

        builder = build_workbook(
            self, export_path,
            enable_pivots=self.native_pivots,
            progress=progress)

        self.export_disclosures = list(builder.omissions)
        self.detail_path = builder.detail_path
        self.detail_rows_written = builder.detail_rows_written
        self.detail_truncated = builder.detail_truncated
        return True
    
    def _build_logic_checks_df(self) -> pd.DataFrame:
        """One row per logic rule, with the result and how many rows failed it."""
        rows = []

        dates = self.analyzer.check_date_logic()
        if dates.get('has_date_columns'):
            c_col = dates.get('created_col', 'created')
            cl_col = dates.get('closed_col', 'closed')
            rows.append(("Date sequence", f"{cl_col} earlier than {c_col}",
                         dates.get('closed_before_created', 0)))
            rows.append(("Future dates", f"{c_col} in the future",
                         dates.get('future_created', 0)))
            rows.append(("Future dates", f"{cl_col} in the future",
                         dates.get('future_closed', 0)))
        else:
            rows.append(("Date sequence", "Not checked - no date columns found", None))
            rows.append(("Future dates", "Not checked - no date columns found", None))

        xf = self.analyzer.check_cross_field_logic()
        if xf.get('has_state'):
            state_col = xf.get('state_col', 'state')
            if xf.get('has_closed_date'):
                closed_col = xf.get('closed_col', 'closed')
                rows.append(("State / resolution",
                             f"Closed in {state_col} but no {closed_col}",
                             xf.get('closed_no_date', 0)))
                rows.append(("State / resolution",
                             f"Open in {state_col} but {closed_col} is set",
                             xf.get('open_with_date', 0)))
            else:
                rows.append(("State / resolution",
                             "Not checked - no closed/resolved date column", None))
            if xf.get('has_priority'):
                rows.append(("Open ticket priority",
                             f"Open in {state_col} with no "
                             f"{xf.get('priority_col', 'priority')}",
                             xf.get('open_missing_priority', 0)))
            else:
                rows.append(("Open ticket priority",
                             "Not checked - no priority column", None))
            rows.append(("State classification",
                         f"{state_col} value read as neither open nor closed",
                         xf.get('unclassified_count', 0)))
        else:
            rows.append(("State / resolution", "Not checked - no state column", None))
            rows.append(("Open ticket priority", "Not checked - no state column", None))

        vocab = self.analyzer.check_state_values()
        if vocab.get('has_state'):
            rows.append(("State vocabulary", "Non-standard or rare state values",
                         vocab.get('total_suspect_rows', 0)))

        def verdict(check, count):
            # A rule that could not run reads as "Not checked", never as 0/Pass:
            # a zero would credit the file for validation that never happened.
            if count is None:
                return "Not checked"
            if count == 0:
                return "Pass"
            # An unreadable state value is a limit of this tool's vocabulary, not
            # a defect in the data, so it must not be reported as a failure.
            return "Review" if check == "State classification" else "Fail"

        return pd.DataFrame(
            [{'Check': check, 'Rule': rule,
              'Rows Affected': "Not checked" if count is None else count,
              'Result': verdict(check, count)}
             for check, rule, count in rows]
        )

    def _build_state_values_df(self) -> pd.DataFrame:
        """The state vocabulary, with each value's frequency and assessment."""
        vocab = self.analyzer.check_state_values()
        if not vocab.get('has_state'):
            return pd.DataFrame()

        reads_as = {'terminal': 'Closed', 'active': 'Open', 'unknown': 'Unrecognised'}
        rows = []
        for entry in vocab.get('suspect_values', []) + vocab.get('standard_values', []):
            rows.append({
                'State Value': entry['value'],
                'Count': entry['count'],
                'Share %': entry['pct'],
                'Reads As': reads_as.get(entry.get('classification'), 'Unrecognised'),
                'Assessment': entry.get('reason') or 'Standard',
            })
        return pd.DataFrame(rows)

    def _build_description_quality_df(self, quality: Dict[str, Any]) -> pd.DataFrame:
        """Build per-column description length metrics for Excel."""
        data = []
        for info in quality.get('columns', []):
            data.append({
                'Field': info['category'],
                'Column': info['column'],
                'Total Rows': info['total'],
                'Empty': info['empty'],
                'Empty %': info['empty_pct'],
                f"Under {quality['min_length']} chars": info['too_short'],
                'Too Short %': info['too_short_pct'],
                f"Over {quality['max_length']} chars": info['too_long'],
                'Too Long %': info['too_long_pct'],
                'Min Length': info['min_length'],
                'Avg Length': info['avg_length'],
                'Median Length': info['median_length'],
                'Max Length': info['max_length'],
            })
        return pd.DataFrame(data)

    def _build_recurring_issues_df(self, quality: Dict[str, Any]) -> pd.DataFrame:
        """Build the recurring / reworded issue groups for Excel."""
        data = []
        for rep in quality.get('repetition', []):
            for group in rep['top_groups']:
                data.append({
                    'Column': rep['column'],
                    'Type': 'Repeated' if group['kind'] == 'exact' else 'Reworded',
                    'Tickets': group['count'],
                    '% of Rows': group['percentage'],
                    'Sample Ticket IDs': ', '.join(group['sample_ids']),
                    'Description': group['text'],
                })
        df = pd.DataFrame(data)
        if df.empty:
            return df
        return df.sort_values('Tickets', ascending=False)

    def _build_language_summary_df(self) -> pd.DataFrame:
        """
        One row per checked column.

        'Status' distinguishes a column that came back clean from one that was
        never checked - a 0 in the count column means opposite things in those
        two cases, and 'Detection Method' is what makes the count meaningful at
        all, since a rules-only run finds roughly half of what a full run does.
        """
        data = []

        for result in self._language_results():
            col = result.get('column', '?')

            if 'error' in result:
                data.append({
                    'Column': col, 'Status': 'Error',
                    'Detection Method': 'n/a',
                    'Rows Analyzed': 0, 'Rows Skipped (too short)': 0,
                    'Non-English Rows': 'Not checked', 'Non-English %': 'Not checked',
                    'Top Languages': result['error'],
                })
                continue

            if 'warning' in result:
                data.append({
                    'Column': col, 'Status': 'Not checked',
                    'Detection Method': result.get('mode_label', 'n/a'),
                    'Rows Analyzed': result.get('total_analyzed', 0),
                    'Rows Skipped (too short)': result.get('skipped_short', 0),
                    'Non-English Rows': 'Not checked', 'Non-English %': 'Not checked',
                    'Top Languages': result['warning'],
                })
                continue

            top_langs = sorted(result['languages'].items(), key=lambda x: -x[1])[:5]
            lang_str = ', '.join(
                f"{self.lang_checker.LANG_NAMES.get(l, l)}: {c:,}" for l, c in top_langs
            )
            data.append({
                'Column': col, 'Status': 'Checked',
                'Detection Method': result.get('mode_label', result.get('mode', 'n/a')),
                'Rows Analyzed': result.get('total_analyzed', 0),
                'Rows Skipped (too short)': result.get('skipped_short', 0),
                'Non-English Rows': result['non_english_count'],
                'Non-English %': result['percentage'],
                'Top Languages': lang_str if lang_str else 'All English',
            })

        return pd.DataFrame(data)

    def _build_language_detail_df(self) -> pd.DataFrame:
        """
        One row per (column, language), with its share of the analyzed rows.

        Counts alone cannot be compared between columns or between files, which
        is why the share is stored next to them rather than left for the reader
        to divide out.
        """
        data = []
        for result in self._language_results():
            if 'error' in result or 'warning' in result:
                continue
            analyzed = result.get('total_analyzed', 0)
            methods = result.get('methods', {})
            for lang, count in sorted(result.get('languages', {}).items(),
                                      key=lambda kv: -kv[1]):
                data.append({
                    'Column': result.get('column', '?'),
                    'Language': self.lang_checker.LANG_NAMES.get(lang, lang),
                    'Rows': count,
                    '% of Analyzed': round(count / analyzed * 100, 2) if analyzed else 0.0,
                    'Detection Method': result.get('mode_label', result.get('mode', 'n/a')),
                })
            # Provenance rows: which tier found the detections in this column
            for method, count in sorted(methods.items(), key=lambda kv: -kv[1]):
                data.append({
                    'Column': result.get('column', '?'),
                    'Language': f'[found by {method}]',
                    'Rows': count,
                    '% of Analyzed': round(count / analyzed * 100, 2) if analyzed else 0.0,
                    'Detection Method': result.get('mode_label', result.get('mode', 'n/a')),
                })
        return pd.DataFrame(data)
    
    def _build_data_with_languages(self) -> pd.DataFrame:
        """Build DataFrame with original data plus language detection columns."""
        if not self.lang_columns:
            return self.df.copy()
        
        # Add language columns to the DataFrame
        df_with_lang = self.lang_checker.add_language_columns_to_df(
            self.df, self.lang_columns
        )
        
        return df_with_lang
    
    def _build_summary_df(self) -> pd.DataFrame:
        """Build summary DataFrame for Excel."""
        summary = self.analyzer.get_summary(self.null_threshold)
        
        data = [
            ["Total Rows", self.analyzer.total_rows],
            ["Total Columns", self.analyzer.total_columns],
            ["Missing Required Columns", len(summary['missing_columns'])],
            ["High Null Columns", len(summary['high_null_columns'])],
            ["Duplicate Count", summary['duplicate_count']],
            ["Date Issues", len(summary['date_issues'])],
            ["Overall Status", "Issues Found" if summary['has_issues'] else "OK"]
        ]
        
        return pd.DataFrame(data, columns=["Metric", "Value"])
    
    def _build_column_check_df(self) -> pd.DataFrame:
        """Build column check DataFrame for Excel."""
        results = self.analyzer.check_columns()
        
        data = []
        for req_col, info in results.items():
            data.append({
                "Required Column": req_col,
                "Status": "Found" if info['found'] else "Missing",
                "Matched Columns": ", ".join(info['matched_columns']) if info['found'] else ""
            })
        
        return pd.DataFrame(data)
    
    def _build_null_analysis_df(self) -> pd.DataFrame:
        """Build null analysis DataFrame for Excel - reuses cached analysis."""
        results = self.analyzer.check_nulls(self.null_threshold)
        
        # Build DataFrame directly from dict (faster than row-by-row)
        data = {
            "Column": list(results.keys()),
            "Null Count": [info['count'] for info in results.values()],
            "Null %": [info['percentage'] for info in results.values()],
            "Above Threshold": ["Yes" if info['above_threshold'] else "No" for info in results.values()]
        }
        
        df = pd.DataFrame(data)
        return df.sort_values("Null %", ascending=False)
    
    # ==================== New report surfaces ====================

    def _build_findings_df(self) -> pd.DataFrame:
        """
        Every check reduced to one comparable row: the report's single
        pane of glass.

        `Result` reuses the vocabulary from _build_logic_checks_df, so
        "Not checked" stays distinguishable from "Pass" here too - the whole
        point of the register is that a reader can see, in one place, both what
        failed and what was never examined.

        Sorted Fail, then Not checked, then Review, then Pass, and by severity
        within each: gaps in coverage deserve to sit near the failures rather
        than be buried under a wall of passes.
        """
        rows: List[Dict[str, Any]] = []

        def add(area, check, severity, count, detail, result=None):
            if result is None:
                result = ("Not checked" if count is None
                          else "Pass" if count == 0 else "Fail")
            rows.append({
                'Area': area, 'Check': check,
                'Severity': _SEVERITY_NONE if result == "Pass" else severity,
                'Rows Affected': "Not checked" if count is None else count,
                'Detail': detail, 'Result': result,
            })

        # --- Required columns ---
        column_results = self.analyzer.check_columns()
        missing = [req for req, data in column_results.items() if not data['found']]
        for req in missing:
            add("Structure", f"Required column: {req}", "Critical", 1,
                "No column in the file could be matched to this requirement")
        if not missing:
            add("Structure", "Required columns", "Critical", 0,
                f"All {len(column_results)} required columns matched")

        # --- Nulls ---
        nulls = self.analyzer.check_nulls(self.null_threshold)
        breaching = {col: info for col, info in nulls.items()
                     if info['above_threshold']}
        for col, info in sorted(breaching.items(),
                                key=lambda kv: -kv[1]['percentage']):
            add("Completeness", f"Nulls in {col}", "High", info['count'],
                f"{info['percentage']}% null, over the {self.null_threshold}% threshold")
        if not breaching:
            add("Completeness", "Null thresholds", "High", 0,
                f"No column exceeds {self.null_threshold}% null")

        # --- Date and cross-field logic: reuse the logic-rule rows verbatim,
        # so the register and the Logic Checks block can never disagree.
        severity_by_check = {
            "Date sequence": "Critical",
            "Future dates": "Medium",
            "State / resolution": "High",
            "Open ticket priority": "Medium",
            "State classification": "Low",
            "State vocabulary": "Medium",
        }
        logic = self._build_logic_checks_df()
        for _, row in logic.iterrows():
            count = row['Rows Affected']
            add("Logic", row['Check'], severity_by_check.get(row['Check'], "Medium"),
                None if count == "Not checked" else int(count),
                row['Rule'], result=row['Result'])

        # --- Date field order ---
        # Ranked Critical: a misread date does not fail a check, it silently
        # changes every number derived from it. A 3-day resolution read the
        # other way round becomes 92.
        for info in self.analyzer.check_date_formats().get('columns', []):
            column = info['column']
            if info.get('conflicting'):
                add("Structure", f"Date format of {column}", "Critical",
                    info['total'],
                    f"Mixed field orders in one column "
                    f"({info['day_first_rows']} rows can only be DD/MM, "
                    f"{info['month_first_rows']} can only be MM/DD) - no single "
                    f"reading is correct")
            elif info.get('needs_confirmation'):
                add("Structure", f"Date format of {column}", "Critical",
                    info['ambiguous'],
                    f"{info['ambiguous']} of {info['total']} values are readable "
                    f"either way (e.g. {', '.join(info['examples'])}); the other "
                    f"reading would change "
                    f"{info['rows_that_would_change']:,} rows",
                    result="Review")

        # --- Duplicates ---
        dupes = self.analyzer.check_duplicates()
        if dupes.get('error'):
            add("Uniqueness", "Duplicate tickets", "High", None, dupes['error'])
        else:
            add("Uniqueness", "Duplicate tickets", "High", dupes['count'],
                f"Surplus rows sharing a {dupes['id_col']} value")

        # --- Description quality ---
        quality = self.analyzer.check_description_quality(
            self.desc_min_length, self.desc_max_length)
        if quality.get('error'):
            add("Text quality", "Description quality", "Medium", None,
                quality['error'])
        else:
            for info in quality.get('columns', []):
                col = info['column']
                add("Text quality", f"Empty {col}", "High", info['empty'],
                    f"{info['empty_pct']}% of rows have no {col}")
                add("Text quality", f"{col} under {self.desc_min_length} chars",
                    "Medium", info['too_short'],
                    f"{info['too_short_pct']}% too brief to analyse")
                add("Text quality", f"{col} over {self.desc_max_length} chars",
                    "Low", info['too_long'],
                    f"{info['too_long_pct']}% too long to review")
                # Rows Affected is the number of rows below Good, not a
                # restatement of the score - a "1" here would read as one bad row.
                distribution = info.get('score_distribution', {})
                below_good = sum(distribution.get(tier, 0)
                                 for tier in ('Critical', 'Poor', 'Fair'))
                score = info.get('overall_score', 0)
                add("Text quality", f"{col} quality score", "Medium", below_good,
                    f"Field scores {score}/100; {below_good:,} rows below Good",
                    result="Pass" if score >= 65 else "Review")
            for rep in quality.get('repetition', []):
                if rep.get('exact_group_count'):
                    add("Text quality", f"Recurring issues in {rep['column']}",
                        "Low", rep['exact_repeated_rows'],
                        f"{rep['exact_group_count']} issues repeat across "
                        f"{rep['exact_repeated_pct']}% of rows", result="Review")

        # --- Language ---
        language_results = self._language_results()
        if not language_results:
            add("Language", "Non-English content", "Medium", None,
                "No language analysis has been run")
        for result in language_results:
            col = result.get('column', '?')
            if 'error' in result:
                add("Language", f"Non-English in {col}", "Medium", None,
                    result['error'])
            elif 'warning' in result:
                add("Language", f"Non-English in {col}", "Medium", None,
                    result['warning'])
            else:
                langs = ", ".join(sorted(result.get('languages', {}))) or "none"
                add("Language", f"Non-English in {col}", "Medium",
                    result.get('non_english_count', 0),
                    f"{result.get('percentage', 0)}% of analysed rows "
                    f"({langs}), via {result.get('mode_label', 'n/a')}")

        frame = pd.DataFrame(rows)
        if frame.empty:
            return frame
        frame['_r'] = frame['Result'].map(_RESULT_ORDER).fillna(len(_RESULT_ORDER))
        frame['_s'] = frame['Severity'].map(_SEVERITY_ORDER).fillna(len(_SEVERITY_ORDER))
        return (frame.sort_values(['_r', '_s', 'Area'])
                     .drop(columns=['_r', '_s'])
                     .reset_index(drop=True))

    def _build_column_health_df(self) -> pd.DataFrame:
        """
        One row per column: what it is, how complete it is, and whether it
        matched a required category.

        check_nulls and get_data_profile were two views of the same key, so they
        are merged here rather than shown as separate tables the reader has to
        line up by eye.
        """
        nulls = self.analyzer.check_nulls(self.null_threshold)
        profile = self.analyzer.get_data_profile()

        # column -> the requirement it satisfied, for the columns that matched
        matched: Dict[str, str] = {}
        for requirement, data in self.analyzer.check_columns().items():
            for col in data.get('matched_columns', []):
                matched.setdefault(col, requirement)

        rows = []
        for info in profile.get('columns', []):
            col = info['name']
            null_info = nulls.get(col, {})
            rows.append({
                'Column': col,
                'Matched Requirement': matched.get(col, ''),
                'Type': info.get('dtype', ''),
                'Sample Value': info.get('sample'),
                'Nulls': null_info.get('count', 0),
                'Null %': null_info.get('percentage', 0.0),
                'Above Threshold': 'Yes' if null_info.get('above_threshold') else 'No',
                'Padded Whitespace': 'Yes' if info.get('has_whitespace') else 'No',
            })
        return pd.DataFrame(rows)

    def _build_monthly_inflow_df(self) -> pd.DataFrame:
        """
        Tickets created per month - the trend the GUI charts and no report has
        ever contained. Also the source range for the inflow chart.
        """
        inflow = self.analyzer.get_monthly_inflow()
        data = inflow.get('data') or {}
        if not data:
            return pd.DataFrame()

        total = sum(data.values())
        return pd.DataFrame([
            {'Month': month, 'Tickets': count,
             'Share %': round(count / total * 100, 2) if total else 0.0}
            for month, count in data.items()
        ])

    def _build_distributions_df(self) -> pd.DataFrame:
        """
        Value counts for every pivotable column, not the four hardcoded
        keywords the old _build_pivot_df matched.

        Capped in two ways, because an unbounded version would bury the sheet:
        columns with more distinct values than MAX_DISTINCT_TO_SUMMARISE are not
        expanded, and each column shows at most TOP_VALUES_PER_COLUMN values.
        Both caps are *reported* rather than applied silently - the remainder
        collapses into one row that keeps its real count, so the totals still
        add up and nobody reads a truncated list as the whole story.
        """
        total = self.analyzer.total_rows

        def share(count: int) -> float:
            # From the count, never by summing the per-value percentages: those
            # are individually rounded, so 60 of them added up to 100.2%.
            return round(count / total * 100, 2) if total else 0.0

        rows = []
        for col in self.analyzer.get_pivot_columns():
            pivot = self.analyzer.get_pivot(col)
            if pivot.get('error'):
                continue
            values = pivot.get('values', [])
            distinct = pivot.get('total_unique', len(values))

            if distinct > MAX_DISTINCT_TO_SUMMARISE:
                covered = sum(item['count'] for item in values)
                rows.append({
                    'Column': col,
                    'Value': f"(not summarised - {distinct:,} distinct values)",
                    'Count': covered,
                    'Share %': share(covered),
                })
                continue

            for item in values[:TOP_VALUES_PER_COLUMN]:
                rows.append({'Column': col, 'Value': item['value'],
                             'Count': item['count'], 'Share %': item['percentage']})

            remainder = values[TOP_VALUES_PER_COLUMN:]
            if remainder:
                remaining_count = sum(item['count'] for item in remainder)
                rows.append({
                    'Column': col,
                    'Value': f"(and {len(remainder):,} more values)",
                    'Count': remaining_count,
                    'Share %': share(remaining_count),
                })

        return pd.DataFrame(rows)

    def _build_pivot_cube_df(self) -> pd.DataFrame:
        """
        Pre-aggregated dimension combinations with a ticket count: the source
        the native PivotTables read from.

        Aggregating first is what keeps pivots affordable - the cube is tens to
        hundreds of rows where the frame is hundreds of thousands, and writing
        its cache records in full is then trivial. The measure is additive on
        purpose; a distinct count or an average cannot be derived from a cube
        like this and must not be added as one.

        Every column the caps exclude is recorded in self.pivot_cube_notes.
        Without that, a reader who expected to pivot by assignment group and
        cannot see it has no way to tell whether the tool ignored the column,
        failed on it, or decided against it - and "the pivots are coarser than
        your data" is exactly the kind of limit this report states rather than
        leaves to be discovered.
        """
        self.pivot_cube_notes = []
        dimensions: List[str] = []
        skipped_dates: List[str] = []
        too_many_values: List[Tuple[str, int]] = []
        past_cap: List[str] = []

        for col in self.analyzer.get_pivot_columns():
            # The date test comes before the dimension cap on purpose: a date is
            # never a useful dimension whether or not slots remain, so reporting
            # it as "we ran out of room" would misattribute the reason. A raw
            # timestamp slices into one row per instant, and the month-level view
            # people actually want is already the Monthly inflow table. A
            # constant date column would otherwise sneak in purely because it has
            # one distinct value.
            if pd.api.types.is_datetime64_any_dtype(self.df[col]):
                skipped_dates.append(col)
                continue

            # The distinct count is also tested before the cap, and for the same
            # reason: a column with thousands of values can never be a dimension,
            # so reporting it as "no slots left" would imply it might be if one
            # were freed. Costs a nunique on every candidate rather than only the
            # ones considered - tens of milliseconds per column, worth it to
            # attribute the reason correctly.
            distinct = int(self.df[col].nunique(dropna=True))
            if distinct > MAX_CUBE_DISTINCT:
                too_many_values.append((col, distinct))
                continue

            if len(dimensions) >= MAX_CUBE_DIMENSIONS:
                past_cap.append(col)
                continue
            dimensions.append(col)

        if skipped_dates:
            self.pivot_cube_notes.append(
                f"Pivots: date columns are not pivot dimensions "
                f"({', '.join(skipped_dates)}) - a raw timestamp splits into one "
                f"row per instant. Use the Monthly inflow table for trend.")
        for col, distinct in too_many_values:
            self.pivot_cube_notes.append(
                f"Pivots: \"{col}\" has {distinct:,} distinct values, past the "
                f"{MAX_CUBE_DISTINCT} allowed for a pivot dimension. Pivot it "
                f"yourself from the Ticket Detail data if you need it.")
        if past_cap:
            self.pivot_cube_notes.append(
                f"Pivots: {len(past_cap)} further column(s) were not used as "
                f"dimensions ({', '.join(past_cap[:4])}"
                f"{'…' if len(past_cap) > 4 else ''}) - the cube uses at most "
                f"{MAX_CUBE_DIMENSIONS}.")

        if not dimensions:
            self.pivot_cube_notes.append(
                "Pivots: no column had few enough distinct values to pivot on, "
                "so the Pivots tab is empty.")
            return pd.DataFrame()

        # observed=True keeps unobserved category combinations out of the cube;
        # without it a categorical dtype produces the full cartesian product.
        grouped = (self.df.groupby(dimensions, dropna=False, observed=True)
                          .size()
                          .reset_index(name='Ticket count'))

        # Drop the least selective dimension until the cube fits. Better a
        # coarser cube than one that dwarfs the findings it supports.
        dropped: List[str] = []
        while len(grouped) > MAX_CUBE_ROWS and len(dimensions) > 1:
            dropped.append(dimensions.pop())
            grouped = (self.df.groupby(dimensions, dropna=False, observed=True)
                              .size()
                              .reset_index(name='Ticket count'))

        if dropped:
            self.pivot_cube_notes.append(
                f"Pivots: dropped {', '.join(dropped)} to keep the cube under "
                f"{MAX_CUBE_ROWS:,} combinations, so the pivots are coarser "
                f"than the data allows. Kept: {', '.join(dimensions)}.")

        # A pivot cache field name must match its header cell exactly, and a
        # NaN dimension value would become a blank cacheField entry.
        for col in dimensions:
            grouped[col] = grouped[col].astype(str).fillna('(blank)')
        return grouped.sort_values('Ticket count', ascending=False).reset_index(drop=True)

    def _build_quality_tiers_df(self) -> pd.DataFrame:
        """
        Description quality tier counts, one row per column: the source for the
        stacked quality chart. Tier order is _SCORE_TIER_ORDER, which is what
        makes the stack read worst-to-best rather than alphabetically.
        """
        quality = self.analyzer.check_description_quality(
            self.desc_min_length, self.desc_max_length)
        if quality.get('error'):
            return pd.DataFrame()

        from core.analyzer import _SCORE_TIER_ORDER
        rows = []
        for info in quality.get('columns', []):
            distribution = info.get('score_distribution', {})
            row = {'Field': info['column']}
            row.update({tier: distribution.get(tier, 0)
                        for tier in _SCORE_TIER_ORDER})
            rows.append(row)
        return pd.DataFrame(rows)

    def _build_language_mix_df(self) -> pd.DataFrame:
        """
        Non-English row counts by language across every checked column.

        Folded to the top MAX_CHART_CATEGORIES with the tail as one "Other" row:
        past ~7 classes adjacent categories stop being distinguishable, and the
        chart is a summary rather than the record - Language Detail keeps the
        full breakdown.
        """
        totals: Dict[str, int] = {}
        for result in self._language_results():
            if 'error' in result or 'warning' in result:
                continue
            for language, count in result.get('languages', {}).items():
                display = self.lang_checker.LANG_NAMES.get(language, language)
                totals[display] = totals.get(display, 0) + count

        if not totals:
            return pd.DataFrame()

        ranked = sorted(totals.items(), key=lambda kv: (-kv[1], kv[0]))
        rows = [{'Language': name, 'Rows': count}
                for name, count in ranked[:MAX_CHART_CATEGORIES]]
        tail = ranked[MAX_CHART_CATEGORIES:]
        if tail:
            rows.append({'Language': f'Other ({len(tail)} languages)',
                         'Rows': sum(count for _name, count in tail)})
        return pd.DataFrame(rows)

    def _build_dimension_mix_df(self, column: Optional[str]) -> pd.DataFrame:
        """
        Value counts for one dimension, folded to MAX_CHART_CATEGORIES.

        Used for the priority and state charts. Kept separate from
        _build_distributions_df because a chart needs a small contiguous range
        that will not move, while that table is a stacked human-readable record.
        """
        if not column:
            return pd.DataFrame()
        pivot = self.analyzer.get_pivot(column)
        if pivot.get('error'):
            return pd.DataFrame()

        values = pivot.get('values', [])
        rows = [{'Value': item['value'], 'Tickets': item['count']}
                for item in values[:MAX_CHART_CATEGORIES]]
        tail = values[MAX_CHART_CATEGORIES:]
        if tail:
            rows.append({'Value': f'Other ({len(tail)} values)',
                         'Tickets': sum(item['count'] for item in tail)})
        return pd.DataFrame(rows)

    def _build_ticket_detail_df(self) -> pd.DataFrame:
        """
        One row per ticket: every original column plus the derived QA flags.

        This is the pivot source, so the flags are written as words rather than
        booleans - a pivot on "Yes"/"No"/"Not checked" is readable, and a rule
        that could not run says so instead of leaving a blank that reads as
        "fine". np.bool_ would also render as 1/0 in Excel rather than TRUE.

        Deliberately does not astype(str) the frame: keeping category dtypes
        roughly halves its memory, and the columns are sanitised for Excel
        separately at write time.
        """
        analyzer = self.analyzer
        detail = self.df.copy()

        created = analyzer._get_created_dates() if analyzer.created_column else None
        closed = analyzer._get_closed_dates() if analyzer.closed_column else None
        if created is not None and closed is not None:
            # Left negative on purpose: a negative resolution time *is* the
            # closed-before-created defect, and clipping it would hide the
            # very rows the flag column marks.
            detail['Resolution Days'] = (
                (closed - created).dt.total_seconds() / 86400).round(2)

        row_flags = analyzer.get_row_flags()
        state = row_flags.get('state') or {}
        if state:
            reads_as = pd.Series('Unrecognised', index=detail.index, dtype=object)
            reads_as[state['terminal']] = 'Closed'
            reads_as[state['active']] = 'Open'
            detail['State Reads As'] = reads_as

        duplicate_mask = analyzer.duplicate_row_mask()
        if duplicate_mask is not None:
            detail['Duplicate ID'] = np.where(duplicate_mask, 'Yes', 'No')
            detail['ID Occurrences'] = analyzer.id_occurrence_counts()

        flag_columns: List[str] = []
        for flag, label in _FLAG_LABELS.items():
            if flag in row_flags['flags']:
                detail[label] = np.where(row_flags['flags'][flag], 'Yes', 'No')
            else:
                detail[label] = 'Not checked'
            flag_columns.append(label)

        # QA Flag Count / QA Findings from the distinct mask combinations rather
        # than a per-row join: at 300k rows the combinations number a handful.
        present = [f for f in _FLAG_LABELS if f in row_flags['flags']]
        if present:
            stacked = pd.concat([row_flags['flags'][f] for f in present], axis=1)
            stacked.columns = present
            detail['QA Flag Count'] = stacked.sum(axis=1).astype(int)

            # Packed in numpy, not pandas: a Series has no __lshift__.
            codes = np.zeros(len(detail), dtype='int64')
            for bit, flag in enumerate(present):
                codes |= row_flags['flags'][flag].to_numpy().astype('int64') << bit
            packed = pd.Series(codes, index=detail.index)
            labels = {
                code: ", ".join(_FLAG_LABELS[f] for bit, f in enumerate(present)
                                if code >> bit & 1) or 'None'
                for code in packed.unique()
            }
            detail['QA Findings'] = packed.map(labels)
        else:
            detail['QA Flag Count'] = 0
            detail['QA Findings'] = 'Not checked'

        for col, scores in analyzer.get_description_row_scores(
                self.desc_min_length, self.desc_max_length).items():
            detail[f'{col} Score'] = scores['score']
            detail[f'{col} Tier'] = scores['tier']

        for col, language in self._row_language_columns().items():
            detail[f'{col} Language'] = language

        return detail

    def _row_language_columns(self) -> Dict[str, pd.Series]:
        """
        Per-row detected language for each column a check has covered.

        Built from the keys check_column already returns, so it costs nothing:
        detection is the slowest operation in the app and re-running it per row
        would add minutes. That also means only columns the user actually
        checked get a column here - which is the honest outcome, and why the
        values distinguish four states rather than defaulting to English:

            Not checked          no language check has run for this column
            Too short to check   below the check's min_length, never judged
            English              looked at, no non-English signal found
            <language>           the detected language

        Never blank, and never "English" for a row nobody looked at.
        """
        columns: Dict[str, pd.Series] = {}
        for result in self._language_results():
            column = result.get('column')
            if not column or column not in self.df.columns:
                continue
            if 'error' in result or 'warning' in result:
                # The check could not run, so every row is unknown - not English.
                columns[column] = pd.Series(
                    'Not checked', index=self.df.index, dtype=object)
                continue

            # Defensive: tests and older callers hand-build these dicts.
            row_languages = result.get('row_languages') or {}
            analyzed = result.get('analyzed_mask')
            if analyzed is None:
                continue

            language = pd.Series(
                np.where(np.asarray(analyzed), 'English', 'Too short to check'),
                index=self.df.index, dtype=object)
            if row_languages:
                detected = pd.Series(row_languages, dtype=object)
                detected = detected.reindex(
                    detected.index.intersection(self.df.index))
                language.loc[detected.index] = detected.map(
                    lambda value: self.lang_checker.LANG_NAMES.get(value, value))
            columns[column] = language
        return columns

