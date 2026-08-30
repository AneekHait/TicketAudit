"""
Core data analysis logic for Sanity Check App.
All business logic for data validation, null analysis, duplicate detection, etc.

Performance optimizations:
- Aggressive caching of expensive operations
- NumPy vectorization for calculations
- Lazy property evaluation
- Minimized DataFrame copies
"""
import pandas as pd
import numpy as np
import datetime
import difflib
import re
import warnings
from typing import Any, Dict, List, NamedTuple, Optional, Sequence, Tuple
from functools import lru_cache
from config import (
    COLUMN_KEYWORDS, ID_KEYWORDS, ID_PATTERNS, VALUE_PATTERNS,
    EXCLUDE_PATTERNS, DESCRIPTION_STOPWORDS
)
from core.language import LanguageChecker

# Keywords used to classify state/status field values as terminal or active.
# Substring matching is intentional: "Closed - Resolved", "Cancelled by user",
# "Work In Progress" all classify correctly without needing an exhaustive list.
_TERMINAL_STATE_KEYWORDS = frozenset({
    'closed', 'resolved', 'completed', 'cancelled', 'canceled',
    'done', 'solved', 'close', 'complete', 'cancel', 'fulfilled',
})
_ACTIVE_STATE_KEYWORDS = frozenset({
    'open', 'new', 'active', 'pending', 'assigned', 'awaiting',
    'hold', 'wip', 'acknowledged', 'investigating', 'reopened', 'progress',
})

# Cancelled is a subset of terminal, kept separate because dropping those rows
# is a thing people want to do - a cancelled ticket was never worked, so it
# flatters resolution times and pads volume counts.
#
# Deliberately narrow. One substring covers "Cancelled", "Canceled",
# "Cancelling", "Closed Cancelled" and "Cancelled by user" without guessing:
# every one of them contains "cancel". Words that arguably mean the same thing -
# withdrawn, abandoned, rejected, void - are *not* here, because whether they
# belong is a judgement about someone else's vocabulary that this tool cannot
# make. What it can do instead is show which values it matched, so a vocabulary
# it does not cover is visible rather than silently kept.
_CANCELLED_STATE_KEYWORDS = frozenset({'cancel'})

# A state value below this share of records AND at most this many occurrences is
# reported as suspect. Both conditions are required: 4 rows out of 20 is 20% and
# perfectly normal, while 4 rows out of 200k is a typo or a stray integration
# value. Either test alone gets one of those two cases wrong.
SUSPECT_STATE_PCT = 1.0
SUSPECT_STATE_COUNT = 5

# Similarity at or above which a rarer value is reported as a likely misspelling
# of a more common one. 0.85 separates the cases that matter from the ones that
# must not fire: 'closd'/'closed' scores 0.91 and 'in progres'/'in progress'
# 0.95, while genuinely distinct codes stay well below - 'p1'/'p2' is 0.50 and
# 'open'/'reopened' 0.67. Lowering it starts flagging real vocabulary.
NEAR_DUPLICATE_RATIO = 0.85

# Above this many distinct values the field is not a controlled vocabulary, so
# the pairwise comparison is skipped rather than run over free text.
MAX_VOCABULARY_FOR_FUZZY = 200


def _find_near_duplicates(entries: List[Dict[str, Any]]) -> Dict[str, str]:
    """
    Map each state value that looks like a misspelling of a *more common* value
    to that value: {'Closd': 'Closed'}. `entries` must be sorted by descending
    count.

    This exists because rarity alone is the weaker signal and is scale
    dependent: 'Closd' is 1.7% of a 60-row export, so a frequency threshold
    either misses it there or drowns a 200k-row file in false positives. Being
    one edit away from a value used far more often is scale-free, and it is the
    actual tell for the three ways a controlled vocabulary degrades - typos
    ('Closd'), case variants ('closed'), and whitespace variants ('Closed ').

    Only values strictly rarer than their twin are flagged, so the common
    spelling is never reported as a misspelling of the rare one, and two values
    of equal frequency never flag each other.
    """
    if len(entries) < 2 or len(entries) > MAX_VOCABULARY_FOR_FUZZY:
        # Guard against a misdetected high-cardinality column: this comparison
        # is O(n^2) in the distinct count, which is free for the dozen values a
        # real state field holds and ruinous for a free-text one.
        return {}

    matches: Dict[str, str] = {}
    for i, entry in enumerate(entries):
        value = entry['value']
        folded = value.strip().lower()
        if not folded:
            continue
        for other in entries[:i]:                       # strictly more common
            if other['count'] <= entry['count']:
                continue
            other_folded = other['value'].strip().lower()
            if not other_folded or other_folded == folded:
                # Identical once folded: a pure case/whitespace variant
                matches[value] = other['value']
                break
            if difflib.SequenceMatcher(None, folded, other_folded).ratio() >= NEAR_DUPLICATE_RATIO:
                matches[value] = other['value']
                break
    return matches


def _fold(value: Any) -> str:
    """Lowercased, whitespace-stripped form of a single cell; '' when missing."""
    if value is None or value is pd.NaT:
        return ''
    try:
        if pd.isna(value):
            return ''
    except (TypeError, ValueError):
        pass                      # array-like or otherwise not a scalar NA
    return str(value).strip().lower()


def _masks_by_folded_value(series: pd.Series, folded_to_group) -> Dict[str, pd.Series]:
    """
    Group a column's rows by a classification of its *distinct* values.

    `folded_to_group` is called once per distinct value (on its folded form) and
    returns a group name. The result maps each group name to a boolean row mask.

    Doing it this way rather than transforming the column is what keeps this
    affordable. The obvious implementation - `astype(str).str.strip().str.lower()`
    then compare - builds three full-length intermediate string arrays and cost
    0.43s on a 300k-row file, enough to make the Logic Checks view visibly
    stutter. A state or priority field holds a handful of distinct values, so
    folding those and using `isin` runs the string work a few times instead of a
    few hundred thousand: the same trick, and the same reasoning, as
    LanguageChecker.clean_series.

    Using `isin` on the original values also sidesteps two dtype traps that
    caught earlier versions of this code: `fillna('')` raises on a categorical
    unless '' is already a category, and under pandas 3's `str` dtype
    `astype(str)` *preserves* NaN instead of rendering it as the string 'nan',
    so testing for sentinel strings silently matches nothing.
    """
    groups: Dict[str, list] = {}
    for value in series.dropna().unique():
        groups.setdefault(folded_to_group(_fold(value)), []).append(value)

    masks = {name: series.isin(values) for name, values in groups.items()}
    # Nulls belong to no value group, so they are folded in explicitly rather
    # than being silently dropped from every mask.
    null_mask = series.isna()
    if null_mask.any():
        blank_group = folded_to_group('')
        masks[blank_group] = masks.get(blank_group, null_mask & False) | null_mask
    return masks


def _classify_state(value: str) -> str:
    """
    Classify a normalised (stripped, lowercased) state value as
    'terminal', 'active' or 'unknown'.

    Terminal keywords are tested first, which decides the genuinely ambiguous
    values: "Resolved Pending Closure" and "Closed Pending Verification" both
    contain an active keyword too, and both are counted as terminal. That is the
    right call for ServiceNow's vocabulary, where such values follow the actual
    resolution - but it is a judgement, not a fact, and it will misread a value
    like "Pending Close" that means the opposite. Anything matching neither list
    lands in 'unknown' and is counted separately rather than being forced into a
    bucket; check_state_values() lists the real values so an unexpected
    vocabulary is visible rather than silently miscounted.
    """
    if value in ('nan', 'none', '', '<na>'):
        return 'unknown'
    for keyword in _TERMINAL_STATE_KEYWORDS:
        if keyword in value:
            return 'terminal'
    for keyword in _ACTIVE_STATE_KEYWORDS:
        if keyword in value:
            return 'active'
    return 'unknown'


# A slash- or dash-separated date whose first two components are numeric, e.g.
# 05/03/2024 or 5-3-24. ISO (2024-03-05) is caught too but resolved as
# unambiguous below, because a 4-digit leading year fixes the field order.
_NUMERIC_DATE = re.compile(r'^\s*(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{2,4})')

# Field order verdicts from _detect_date_order.
DATE_ORDER_DAY_FIRST = 'day_first'
DATE_ORDER_MONTH_FIRST = 'month_first'
DATE_ORDER_ISO = 'iso'
DATE_ORDER_AMBIGUOUS = 'ambiguous'
DATE_ORDER_CONFLICTING = 'conflicting'
DATE_ORDER_NOT_NUMERIC = 'not_numeric'


def _detect_date_order(values: Sequence[Any]) -> Dict[str, Any]:
    """
    Work out whether a text date column is day-first, month-first, or genuinely
    undecidable.

    This exists because pandas will happily guess. Given a column where every
    value has both components <= 12 there is no signal at all, so it silently
    assumes month-first: a European export of 05/03/2024 is read as 5 May
    instead of 3 March, and the resolution time it feeds into comes out ~30x
    wrong while every check reports "no issues". Guessing is unavoidable;
    guessing *silently* is not.

    Returns {
      'order':      one of the DATE_ORDER_* verdicts,
      'ambiguous':  how many values could be read either way,
      'decisive':   how many values fix the order on their own,
      'total':      how many values looked like numeric dates,
      'examples':   a few ambiguous values, for the message,
    }

    Reads distinct values, not rows - the answer is a property of the format.
    """
    day_first_evidence = 0        # a first component > 12 can only be a day
    month_first_evidence = 0      # a second component > 12 can only be a day
    ambiguous = 0
    iso = 0
    total = 0
    examples: List[str] = []

    for value in values:
        if not isinstance(value, str):
            continue
        match = _NUMERIC_DATE.match(value)
        if not match:
            continue
        total += 1
        first, second = int(match.group(1)), int(match.group(2))

        if len(match.group(1)) == 4:
            iso += 1              # a leading 4-digit year settles the order
            continue
        if first > 12 and second <= 12:
            day_first_evidence += 1
        elif second > 12 and first <= 12:
            month_first_evidence += 1
        elif first <= 12 and second <= 12:
            ambiguous += 1
            if len(examples) < 3:
                examples.append(value.strip())

    if total == 0:
        order = DATE_ORDER_NOT_NUMERIC
    elif iso == total:
        order = DATE_ORDER_ISO
    elif day_first_evidence and month_first_evidence:
        # Both orders are proven present, so no single reading is right. That is
        # a defect in the data, not a choice to be made.
        order = DATE_ORDER_CONFLICTING
    elif day_first_evidence:
        order = DATE_ORDER_DAY_FIRST
    elif month_first_evidence:
        order = DATE_ORDER_MONTH_FIRST
    else:
        order = DATE_ORDER_AMBIGUOUS

    return {
        'order': order,
        'ambiguous': ambiguous,
        'decisive': day_first_evidence + month_first_evidence,
        'day_first_evidence': day_first_evidence,
        'month_first_evidence': month_first_evidence,
        'total': total,
        'examples': examples,
    }


# Below this a candidate is not considered a match at all. Scores are 100 exact
# name / 80 endswith / 70 startswith / 50 word-boundary, plus 30 for matching the
# category's value pattern, plus 60 for ID-shaped values.
MATCH_SCORE_THRESHOLD = 30

# Sentinel for "this category is genuinely absent" in a column override, as
# distinct from "no override set". None cannot carry that meaning because a dict
# lookup returning None is indistinguishable from a missing key.
COLUMN_ABSENT = '__absent__'


_SCORE_TIER_ORDER = ('Critical', 'Poor', 'Fair', 'Good', 'Excellent')

# Tier -> contribution to a field's single 0-100 score.
_TIER_WEIGHTS = {'Critical': 0, 'Poor': 25, 'Fair': 50, 'Good': 80, 'Excellent': 100}


class DescriptionRowFrame(NamedTuple):
    """
    The per-row inputs every description measure is derived from.

    Extracted so the aggregate metrics and the per-row scores are computed from
    one set of arrays. It carries no scores itself - see
    SanityAnalyzer.get_description_row_scores.
    """
    text: 'pd.Series'          # stripped strings, never NaN
    lengths: 'pd.Series'
    empty_mask: 'pd.Series'
    populated_mask: 'pd.Series'
    short_mask: 'pd.Series'    # populated but under min_length
    long_mask: 'pd.Series'
    group_sizes: 'pd.Series'   # how many rows share this exact text; 0 when empty


def _score_description_rows(
    lengths: 'pd.Series',
    empty_mask: 'pd.Series',
    short_mask: 'pd.Series',
    long_mask: 'pd.Series',
    group_sizes: 'pd.Series',
    min_length: int,
) -> 'Tuple[pd.Series, pd.Series]':
    """
    Assign a 0-100 quality score and tier label to every row of a description
    column.  Works on the full index (including empty rows).

    Tiers and thresholds:
      Critical  0        – empty / null
      Poor      5-39     – under min_length; scaled so longer-but-still-short
                           scores higher than a single character
      Fair      40-58    – too long (>max_length) OR in-range but the same text
                           appears on ≥3 tickets (template / boilerplate)
      Good      65-89    – in range, unique or near-unique, moderate length
      Excellent 90-100   – in range, unique, substantive length (>= min + 1000
                           chars caps the bonus)

    group_sizes is a Series aligned to lengths.index; for each row it contains
    the number of tickets that share the identical raw description text.
    Rows not in populated_mask should carry 0.
    """
    idx = lengths.index
    ln = lengths.to_numpy(dtype=float)
    em = empty_mask.to_numpy()
    sm = short_mask.to_numpy()
    lm = long_mask.to_numpy()
    gs = group_sizes.to_numpy(dtype=float)

    scores_arr = np.zeros(len(ln), dtype=int)
    tiers_arr = np.full(len(ln), 'Critical', dtype=object)

    # Poor: populated but too short
    scores_arr[sm] = np.clip(ln[sm] / max(min_length, 1) * 39, 5, 39).astype(int)
    tiers_arr[sm] = 'Poor'

    # Fair: too long
    scores_arr[lm] = 40
    tiers_arr[lm] = 'Fair'

    # In range — further split by template size
    ir = ~em & ~sm & ~lm
    large_tpl = ir & (gs >= 10)
    mod_rep   = ir & (gs >= 3) & (gs < 10)
    unique_   = ir & (gs < 3)

    scores_arr[large_tpl] = 45
    tiers_arr[large_tpl] = 'Fair'

    scores_arr[mod_rep] = 58
    tiers_arr[mod_rep] = 'Fair'

    bonus = np.clip((ln[unique_] - min_length) / 1000.0 * 35, 0, 35).astype(int)
    base = np.clip(65 + bonus, 65, 100)
    scores_arr[unique_] = base
    tiers_arr[unique_] = np.where(base >= 90, 'Excellent', 'Good')

    return pd.Series(scores_arr, index=idx), pd.Series(tiers_arr, index=idx)


class SanityAnalyzer:
    """
    Core analyzer class that performs all data quality checks.
    Completely independent of GUI - returns structured data.
    """
    
    def __init__(self, df: pd.DataFrame, dayfirst: Optional[bool] = None,
                 column_overrides: Optional[Dict[str, str]] = None):
        """
        dayfirst pins how text dates are read: None leaves pandas to infer (the
        historical behaviour), True forces DD/MM, False forces MM/DD. See
        check_date_formats() for why the caller may need to decide - pandas
        infers per column, so two date columns in one file can silently disagree
        about the same format.

        column_overrides maps a required category to the column it should read,
        overriding the automatic name matching: {"Priority": "urgency"}. Use
        COLUMN_ABSENT as the value to declare a category genuinely missing. See
        _apply_column_overrides.
        """
        self.df = df
        self.dayfirst = dayfirst
        self.column_overrides = dict(column_overrides or {})
        self.override_notes: List[str] = []
        self.overridden: Dict[str, Optional[str]] = {}
        self._cache: Dict[str, Any] = {}
        self._build_cache()
    
    def _normalize_column_name(self, col: str) -> str:
        """Normalize column name for comparison."""
        return col.lower().strip().replace(" ", "_").replace("-", "_")
    
    def _is_excluded(self, required: str, col: str) -> bool:
        """Check if column should be excluded from this match."""
        col_normalized = self._normalize_column_name(col)
        exclude_patterns = EXCLUDE_PATTERNS.get(required, [])
        
        for pattern in exclude_patterns:
            if re.match(pattern, col_normalized, re.IGNORECASE):
                return True
        return False
    
    def _check_column_values_for_ids(self, col: str) -> bool:
        """Check if column contains ticket ID patterns like INC12345, CR12345, etc."""
        try:
            sample = self.df[col].dropna().head(100).astype(str)
            if len(sample) == 0:
                return False
            
            match_count = 0
            for value in sample:
                value = value.strip().upper()
                if not value or len(value) > 50:
                    continue
                for pattern in ID_PATTERNS:
                    if re.match(pattern, value, re.IGNORECASE):
                        match_count += 1
                        break
            
            return match_count / len(sample) > 0.5
        except Exception:
            return False

    def _check_column_values_for_type(self, col: str, required: str) -> bool:
        """Check if column values match expected patterns for a column type."""
        patterns = VALUE_PATTERNS.get(required, [])
        if not patterns:
            return False
        
        try:
            sample = self.df[col].dropna().head(100).astype(str)
            if len(sample) == 0:
                return False
            
            match_count = 0
            for value in sample:
                value = value.strip()
                for pattern in patterns:
                    if re.match(pattern, value, re.IGNORECASE):
                        match_count += 1
                        break
            
            return match_count / len(sample) > 0.5
        except Exception:
            return False

    def _match_column_exact(self, required: str, col: str) -> bool:
        """Check for exact or near-exact column name match."""
        if self._is_excluded(required, col):
            return False
            
        col_normalized = self._normalize_column_name(col)
        keywords = COLUMN_KEYWORDS.get(required, [])
        
        for keyword in keywords:
            keyword_normalized = keyword.lower().strip().replace(" ", "_").replace("-", "_")
            if col_normalized == keyword_normalized:
                return True
            if col_normalized.endswith("_" + keyword_normalized):
                return True
            if col_normalized.startswith(keyword_normalized + "_"):
                return True
        
        return False
    
    def _match_column_contains(self, required: str, col: str) -> bool:
        """Check if column name contains keyword as a word boundary."""
        if self._is_excluded(required, col):
            return False
            
        col_normalized = self._normalize_column_name(col)
        keywords = COLUMN_KEYWORDS.get(required, [])
        
        for keyword in keywords:
            keyword_normalized = keyword.lower().strip().replace(" ", "_").replace("-", "_")
            if len(keyword_normalized) >= 3:
                pattern = r'(^|_)' + re.escape(keyword_normalized) + r'($|_)'
                if re.search(pattern, col_normalized):
                    return True
        
        return False
    
    def _calculate_match_score(self, required: str, col: str) -> int:
        """Calculate a match score for ranking candidates."""
        score = 0
        col_normalized = self._normalize_column_name(col)
        keywords = COLUMN_KEYWORDS.get(required, [])
        
        for keyword in keywords:
            keyword_normalized = keyword.lower().strip().replace(" ", "_").replace("-", "_")
            
            if col_normalized == keyword_normalized:
                score = max(score, 100)
            elif col_normalized.endswith("_" + keyword_normalized):
                score = max(score, 80)
            elif col_normalized.startswith(keyword_normalized + "_"):
                score = max(score, 70)
            elif len(keyword_normalized) >= 3:
                pattern = r'(^|_)' + re.escape(keyword_normalized) + r'($|_)'
                if re.search(pattern, col_normalized):
                    score = max(score, 50)
        
        # Bonus for value pattern match
        if self._check_column_values_for_type(col, required):
            score += 30
        
        return score
    
    def _build_cache(self) -> None:
        """Pre-compute and cache expensive calculations using vectorized ops."""
        if self.df is None:
            return
        
        # Cache total rows (accessed frequently)
        self._cache['total_rows'] = len(self.df)
        
        # Pre-compute lowercase column names for faster lookups
        self._cache['col_lower_map'] = {col: col.lower() for col in self.df.columns}
        
        # Build column mappings with improved matching, then let the user's
        # overrides win. Applied here rather than at read time so everything
        # derived below - the ID column, the date columns, and through them every
        # check - sees the corrected mapping.
        self._cache['column_map'] = self._apply_column_overrides(
            self._build_column_map())

        # Cache ID column detection
        self._cache['id_col'] = self._cache['column_map'].get('Ticket Identification', [None])[0] if self._cache['column_map'].get('Ticket Identification') else self._find_col(ID_KEYWORDS)
        
        # Cache date columns
        created_matches = self._cache['column_map'].get('Created', [])
        closed_matches = self._cache['column_map'].get('Closed', [])
        self._cache['created_col'] = created_matches[0] if created_matches else self._find_date_col(["created", "opened"])
        self._cache['closed_col'] = closed_matches[0] if closed_matches else self._find_date_col(["closed", "resolved"])
    
    def _build_column_map(self) -> Dict[str, List[str]]:
        """Build column mapping with improved matching logic."""
        column_map = {}
        scores: Dict[str, List[Tuple[str, int]]] = {}
        used_columns = set()
        
        # Priority order for matching
        priority_order = [
            "Ticket Identification", "Assignment Group", "Configuration Item",
            "Priority", "State/Status", "Created", "Closed",
            "Short Description", "Description"
        ]
        
        for required in priority_order:
            if required not in COLUMN_KEYWORDS:
                continue
                
            candidates = []
            
            for col in self.df.columns:
                if col in used_columns:
                    continue
                
                score = 0
                
                if self._match_column_exact(required, col):
                    score = self._calculate_match_score(required, col)
                elif self._match_column_contains(required, col):
                    score = self._calculate_match_score(required, col)
                
                if score > 0:
                    candidates.append((col, score))
            
            # Special: Check column VALUES for Ticket Identification
            if required == "Ticket Identification":
                for col in self.df.columns:
                    if col in used_columns:
                        continue
                    if self._check_column_values_for_ids(col):
                        existing = [c for c, s in candidates if c == col]
                        if not existing:
                            candidates.append((col, 60))
                        else:
                            # Boost score if already a candidate
                            candidates = [(c, s + 30 if c == col else s) for c, s in candidates]
            
            # Sort by score only. Python's sort is stable and reverse=True keeps
            # the order of equal elements, so a tie falls to whichever column
            # comes first in the file - which is deliberate, not incidental:
            # exports tend to put the more significant field earlier, and
            # re-breaking ties alphabetically would silently change which column
            # existing files match. Ties are common and real (impact and urgency
            # both score 130 for Priority), which is what column_choices exposes
            # so the user can settle it.
            candidates.sort(key=lambda pair: pair[1], reverse=True)

            # Retained so the UI can show *why* a column was chosen and what the
            # runners-up were. Without the scores a wrong match is invisible: the
            # user sees a confident answer and no way to tell it is a 50-point
            # word-boundary hit over a 100-point exact one that got taken first.
            scores[required] = list(candidates)

            # Get matched columns above threshold
            matched = [col for col, score in candidates if score >= MATCH_SCORE_THRESHOLD]

            if matched:
                used_columns.add(matched[0])
                column_map[required] = matched

        self._cache['column_scores'] = scores
        return column_map
    
    def _apply_column_overrides(
            self, column_map: Dict[str, List[str]]) -> Dict[str, List[str]]:
        """
        Let the caller's overrides replace what the scoring chose.

        Automatic matching is a heuristic over column *names*, so it can pick
        `sys_updated_on` as Created or a `tower` column as Assignment Group. Every
        check downstream then measures the wrong field, confidently. Before this
        there was no way to correct it.

        An override naming a column that is not in this file is ignored rather
        than guessed at, and recorded in self.override_notes so a stale setting
        from another export shape cannot silently do nothing.
        """
        self.override_notes: List[str] = []
        self.overridden: Dict[str, Optional[str]] = {}
        if not self.column_overrides:
            return column_map

        for required, chosen in self.column_overrides.items():
            if required not in COLUMN_KEYWORDS:
                self.override_notes.append(
                    f'Column override ignored: "{required}" is not one of the '
                    f'required categories.')
                continue

            if chosen == COLUMN_ABSENT:
                column_map.pop(required, None)
                self.overridden[required] = None
                self.override_notes.append(
                    f'Column override: "{required}" marked as not present in '
                    f'this file, so its checks are reported as not checked.')
                continue

            if chosen not in self.df.columns:
                self.override_notes.append(
                    f'Column override ignored: "{required}" was set to '
                    f'"{chosen}", which is not a column in this file.')
                continue

            # Put the chosen column at the head; keep the rest as alternatives.
            others = [c for c in column_map.get(required, []) if c != chosen]
            column_map[required] = [chosen] + others
            self.overridden[required] = chosen
            self.override_notes.append(
                f'Column override: "{required}" read from "{chosen}" instead of '
                f'the automatic match.')

        return column_map

    def _find_col(self, keywords: List[str]) -> Optional[str]:
        """Helper to find a column by keywords using cached lowercase names."""
        col_lower_map = self._cache.get('col_lower_map', {})
        for c in self.df.columns:
            c_lower = col_lower_map.get(c, c.lower())
            if any(k in c_lower for k in keywords):
                return c
        return None
    
    def _find_date_col(self, keywords: List[str]) -> Optional[str]:
        """
        Find a date column by keywords, excluding non-date columns like 'Created By'.
        Validates that the column actually contains datetime-parseable values.
        """
        exclude_patterns = ['by', 'user', 'name', 'person', 'agent', 'owner']
        
        for col in self.df.columns:
            col_lower = col.lower()
            # Check if column matches keywords
            if not any(k in col_lower for k in keywords):
                continue
            # Exclude columns like "Created By", "Resolved By User", etc.
            if any(excl in col_lower for excl in exclude_patterns):
                continue
            # Validate that column contains date/time values
            try:
                sample = self.df[col].dropna().head(10)
                if len(sample) > 0:
                    parsed = pd.to_datetime(sample, errors='coerce')
                    # If at least 50% parse as dates, it's a date column
                    if parsed.notna().sum() >= len(sample) * 0.5:
                        return col
            except Exception:
                continue
        return None

    def _parse_dates(self, column: str) -> pd.Series:
        """
        Parse one date column under this analyzer's day-first setting.

        `dayfirst=None` leaves pandas to infer, which is the historical
        behaviour; True or False pins it. Pinning matters because pandas infers
        per column, so two date columns in the same file can silently disagree
        about the same format.
        """
        kwargs = {'errors': 'coerce'}
        if self.dayfirst is not None:
            kwargs['dayfirst'] = self.dayfirst
        # pandas warns when it infers day-first while dayfirst=False; the warning
        # goes to stderr where no user sees it, so check_date_formats() reports
        # the same fact where they will.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            return pd.to_datetime(self.df[column], **kwargs)

    def _get_created_dates(self) -> Optional[pd.Series]:
        """Lazy load created dates."""
        if 'created_dates' not in self._cache:
            col = self.created_column
            self._cache['created_dates'] = self._parse_dates(col) if col else None
        return self._cache['created_dates']

    def _get_closed_dates(self) -> Optional[pd.Series]:
        """Lazy load closed dates."""
        if 'closed_dates' not in self._cache:
            col = self.closed_column
            self._cache['closed_dates'] = self._parse_dates(col) if col else None
        return self._cache['closed_dates']

    def check_date_formats(self) -> Dict[str, Any]:
        """
        Report how each date column's field order was decided, and what the
        other reading would cost.

        The point is `rows_that_would_change`: "1,204 values are ambiguous" is
        abstract, while "reading them the other way moves 87 rows to a different
        date" tells the user immediately whether the guess matters. A guess
        nobody can see is the problem; a guess with its cost priced is a
        decision.

        Returns {
          'setting':  'inferred' | 'day_first' | 'month_first',
          'columns':  [{column, order, ambiguous, total, examples,
                        rows_that_would_change, needs_confirmation,
                        conflicting, day_first_rows, month_first_rows}],
        }

        Only columns with a real choice to report appear in `columns`: one the
        reader already typed as datetime, or whose text is ISO or non-numeric,
        has no field order to infer and is left out rather than listed as fine.
        """
        if 'date_formats' in self._cache:
            return self._cache['date_formats']

        setting = ('inferred' if self.dayfirst is None
                   else 'day_first' if self.dayfirst else 'month_first')
        result: Dict[str, Any] = {'setting': setting, 'columns': []}

        candidates = [c for c in (self.created_column, self.closed_column) if c]
        for column in candidates:
            series = self.df[column]
            if pd.api.types.is_datetime64_any_dtype(series):
                # Already typed by the reader, so no text order to infer.
                continue

            detection = _detect_date_order(series.dropna().astype(str).unique())
            if detection['order'] in (DATE_ORDER_NOT_NUMERIC, DATE_ORDER_ISO):
                continue

            here = self._parse_dates(column)
            other_dayfirst = not self._effective_dayfirst(detection)
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                other = pd.to_datetime(series, errors='coerce',
                                       dayfirst=other_dayfirst)
            differs = int((here != other).sum() - (here.isna() & other.isna()).sum())

            result['columns'].append({
                'column': column,
                'order': detection['order'],
                'ambiguous': detection['ambiguous'],
                'total': detection['total'],
                'examples': detection['examples'],
                'rows_that_would_change': max(differs, 0),
                # A column whose order nothing proves, and where the other
                # reading changes rows, is a decision for the user.
                'needs_confirmation': (
                    detection['order'] == DATE_ORDER_AMBIGUOUS and differs > 0),
                # Mixed orders in one column are a different thing: no single
                # reading is right, so there is nothing to choose. It is flagged
                # on its own, and independently of rows_that_would_change -
                # pandas resolves 13/01 and 01/13 to the same date, so the two
                # readings agree while the data is still wrong.
                'conflicting': detection['order'] == DATE_ORDER_CONFLICTING,
                'day_first_rows': detection['day_first_evidence'],
                'month_first_rows': detection['month_first_evidence'],
            })

        self._cache['date_formats'] = result
        return result

    def _effective_dayfirst(self, detection: Dict[str, Any]) -> bool:
        """Which order is actually in force for a column with this evidence."""
        if self.dayfirst is not None:
            return self.dayfirst
        # Mirror pandas: decisive evidence wins, otherwise month-first.
        return detection['order'] == DATE_ORDER_DAY_FIRST
    
    @property
    def total_rows(self) -> int:
        return self._cache.get('total_rows', len(self.df))
    
    @property
    def total_columns(self) -> int:
        return len(self.df.columns)
    
    @property
    def id_column(self) -> Optional[str]:
        return self._cache.get('id_col')
    
    @property
    def created_column(self) -> Optional[str]:
        return self._cache.get('created_col')
    
    @property
    def closed_column(self) -> Optional[str]:
        return self._cache.get('closed_col')

    @property
    def state_column(self) -> Optional[str]:
        matches = self._cache.get('column_map', {}).get('State/Status', [])
        return matches[0] if matches else None

    @property
    def priority_column(self) -> Optional[str]:
        matches = self._cache.get('column_map', {}).get('Priority', [])
        return matches[0] if matches else None

    # ==================== Column Check ====================
    
    def check_columns(self) -> Dict[str, Dict[str, Any]]:
        """
        Check for required columns with improved matching.

        Returns: {requirement: {
            'found': bool,
            'matched_columns': [str],       # ranked, chosen first
            'score': int,                   # of the chosen column, 0 if none
            'candidates': [(column, score)],# every scoring column, for the UI
            'overridden': bool,             # the user chose this, not the scoring
        }}

        The score and candidates are what make a wrong match correctable: a
        reader can see that Created was matched at 50 by a word boundary while
        another column scored 100, which is the information needed to override it.
        """
        results = {}
        column_map = self._cache.get('column_map', {})
        all_scores = self._cache.get('column_scores', {})

        for req_col in COLUMN_KEYWORDS.keys():
            matched = column_map.get(req_col, [])
            candidates = all_scores.get(req_col, [])
            by_column = dict(candidates)
            results[req_col] = {
                'found': len(matched) > 0,
                'matched_columns': matched,
                # An overridden column may have scored nothing at all, which is
                # the point of overriding it.
                'score': by_column.get(matched[0], 0) if matched else 0,
                'candidates': candidates,
                'overridden': req_col in self.overridden,
            }
        return results

    def column_choices(self) -> Dict[str, Dict[str, Any]]:
        """
        Everything a mapping-override UI needs, per required category.

        {requirement: {'chosen': str|None, 'score': int, 'overridden': bool,
                       'alternatives': [(column, score)]}}

        `alternatives` excludes the chosen column and keeps the scores, so the
        dialog can show "priority scored 100 but urgency was taken" rather than
        an unordered list of every column in the file.
        """
        choices = {}
        for required, data in self.check_columns().items():
            matched = data['matched_columns']
            chosen = matched[0] if matched else None
            choices[required] = {
                'chosen': chosen,
                'score': data['score'],
                'overridden': data['overridden'],
                'alternatives': [(col, score) for col, score in data['candidates']
                                 if col != chosen],
            }
        return choices
    
    def get_missing_columns(self) -> List[str]:
        """Get list of missing required columns."""
        results = self.check_columns()
        return [req for req, data in results.items() if not data['found']]
    
    # ==================== Null Analysis ====================
    
    def check_nulls(self, threshold: float = 20.0) -> Dict[str, Dict[str, Any]]:
        """
        Analyze null values in all columns using vectorized numpy operations.
        Returns: {column: {'count': int, 'percentage': float, 'above_threshold': bool}}
        """
        null_counts = self._cache.get('null_counts')
        if null_counts is None:
            null_counts = self.df.isnull().sum()
            self._cache['null_counts'] = null_counts
        
        total = self.total_rows
        
        # Vectorized percentage calculation
        if total > 0:
            percentages = (null_counts / total) * 100
        else:
            percentages = pd.Series(0, index=null_counts.index)
        
        # Build results dict efficiently
        results = {}
        for col in self.df.columns:
            count = int(null_counts[col])
            pct = round(float(percentages[col]), 2)
            results[col] = {
                'count': count,
                'percentage': pct,
                'above_threshold': pct > threshold
            }
        return results
    
    def get_high_null_columns(self, threshold: float = 20.0) -> List[str]:
        """Get columns that exceed the null threshold."""
        results = self.check_nulls(threshold)
        return [col for col, data in results.items() if data['above_threshold']]
    
    # ==================== Date Logic Checks ====================
    
    # Stable keys for the per-row defect masks. Used by check_date_logic,
    # check_cross_field_logic and the report's detail sheet, so they are part of
    # this class's contract rather than an implementation detail.
    FLAG_CLOSED_BEFORE_CREATED = 'closed_before_created'
    FLAG_FUTURE_CREATED = 'future_created'
    FLAG_FUTURE_CLOSED = 'future_closed'
    FLAG_CLOSED_NO_DATE = 'closed_no_date'
    FLAG_OPEN_WITH_DATE = 'open_with_date'
    FLAG_OPEN_MISSING_PRIORITY = 'open_missing_priority'

    def get_row_flags(self) -> Dict[str, Any]:
        """
        Per-row masks for every logic rule, plus the state classification.

        {
          'flags':   {flag_name: Series[bool]},   # defect masks, aligned to df.index
          'skipped': [flag_name, ...],            # rules that could not run
          'state':   {'terminal': Series[bool], 'active': Series[bool]},
        }

        This exists so a rule's row-level answer and its headline count come
        from the same array. check_date_logic and check_cross_field_logic sum
        these masks, and the report's detail sheet writes them per row - if the
        two were computed separately, the Logic Checks sheet and the Ticket
        Detail sheet could disagree about the same file, which is worse than
        either being wrong on its own.

        A flag that could not be evaluated is absent from 'flags' and named in
        'skipped'; it is never present-but-all-False, because that is
        indistinguishable from "checked, nothing found".

        Cached, which also pins "future" to one instant for the whole report
        rather than re-reading the clock per call.
        """
        if 'row_flags' in self._cache:
            return self._cache['row_flags']

        index = self.df.index
        empty = pd.Series(False, index=index)
        flags: Dict[str, pd.Series] = {}
        skipped: List[str] = []
        state: Dict[str, pd.Series] = {}

        # --- Date sequence and future dates ---
        created = self._get_created_dates() if self.created_column else None
        closed = self._get_closed_dates() if self.closed_column else None
        now = pd.Timestamp.now()

        if created is not None and closed is not None:
            # A comparison against NaT is False in pandas, so unparsed rows fall
            # out of the mask without a separate notna() guard.
            flags[self.FLAG_CLOSED_BEFORE_CREATED] = closed < created
        else:
            skipped.append(self.FLAG_CLOSED_BEFORE_CREATED)

        for key, series in ((self.FLAG_FUTURE_CREATED, created),
                            (self.FLAG_FUTURE_CLOSED, closed)):
            if series is None:
                skipped.append(key)
            else:
                flags[key] = series > now

        # --- State classification and the cross-field rules ---
        state_col = self.state_column
        if state_col is None:
            state_masks = {}
        else:
            state_masks = _masks_by_folded_value(self.df[state_col], _classify_state)
            state['terminal'] = state_masks.get('terminal', empty)
            state['active'] = state_masks.get('active', empty)

        if state_col is None:
            skipped.extend([self.FLAG_CLOSED_NO_DATE, self.FLAG_OPEN_WITH_DATE,
                            self.FLAG_OPEN_MISSING_PRIORITY])
        else:
            terminal = state['terminal']
            active = state['active']

            if closed is None:
                skipped.extend([self.FLAG_CLOSED_NO_DATE, self.FLAG_OPEN_WITH_DATE])
            else:
                missing_date = closed.isna()
                flags[self.FLAG_CLOSED_NO_DATE] = terminal & missing_date
                flags[self.FLAG_OPEN_WITH_DATE] = active & ~missing_date

            priority_col = self.priority_column
            if priority_col is None:
                skipped.append(self.FLAG_OPEN_MISSING_PRIORITY)
            else:
                blank = _masks_by_folded_value(
                    self.df[priority_col],
                    lambda folded: 'blank' if not folded else 'set',
                ).get('blank', empty)
                flags[self.FLAG_OPEN_MISSING_PRIORITY] = active & blank

        result = {'flags': flags, 'skipped': skipped, 'state': state}
        self._cache['row_flags'] = result
        return result

    def _flag_count(self, name: str) -> int:
        """Headline count for one rule, or 0 when the rule could not run."""
        mask = self.get_row_flags()['flags'].get(name)
        return 0 if mask is None else int(mask.sum())

    def check_date_logic(self) -> Dict[str, Any]:
        """
        Check for date logic issues.
        Returns: {'issues': [...], 'closed_before_created': int, 'future_created': int, 'future_closed': int}
        """
        result = {
            'has_date_columns': False,
            'created_col': self.created_column,
            'closed_col': self.closed_column,
            'issues': [],
            'closed_before_created': 0,
            'future_created': 0,
            'future_closed': 0
        }
        
        created_col = self.created_column
        closed_col = self.closed_column
        
        if not created_col or not closed_col:
            result['issues'].append("Could not find both Created and Closed/Resolved columns")
            return result
        
        result['has_date_columns'] = True

        # Counts come from get_row_flags' masks, which the detail sheet also
        # writes per row, so the two can never disagree.
        row_flags = self.get_row_flags()
        if self.FLAG_CLOSED_BEFORE_CREATED in row_flags['skipped']:
            result['issues'].append("Could not parse date columns")
            return result

        invalid_count = self._flag_count(self.FLAG_CLOSED_BEFORE_CREATED)
        result['closed_before_created'] = invalid_count
        if invalid_count > 0:
            result['issues'].append(f"{invalid_count} rows where {closed_col} < {created_col}")

        future_created = self._flag_count(self.FLAG_FUTURE_CREATED)
        future_closed = self._flag_count(self.FLAG_FUTURE_CLOSED)
        result['future_created'] = future_created
        result['future_closed'] = future_closed

        if future_created > 0:
            result['issues'].append(f"{future_created} rows with future {created_col}")
        if future_closed > 0:
            result['issues'].append(f"{future_closed} rows with future {closed_col}")

        return result
    
    # ==================== Cross-field Logic Checks ====================

    def check_cross_field_logic(self) -> Dict[str, Any]:
        """
        Ticket-specific cross-field consistency checks.

        Three rules that date-sequence checks alone cannot catch:
        - Terminal state (Closed/Resolved) with no resolution date recorded
        - Active state (Open/Pending) with a resolution date already set
        - Open ticket with no priority assigned

        Classification uses substring matching so values like "Closed - Resolved"
        and "Work In Progress" are handled without an exhaustive list.
        """
        state_col = self.state_column
        priority_col = self.priority_column
        closed_col = self.closed_column

        result = {
            'state_col': state_col,
            'priority_col': priority_col,
            'closed_col': closed_col,
            'has_state': state_col is not None,
            'has_priority': priority_col is not None,
            'has_closed_date': closed_col is not None,
            'open_count': 0,
            'closed_count': 0,
            'unclassified_count': 0,
            'closed_no_date': 0,
            'open_with_date': 0,
            'open_missing_priority': 0,
        }

        if not state_col:
            return result

        # Shared with check_date_logic and the report's per-row detail sheet.
        row_flags = self.get_row_flags()
        state = row_flags['state']

        result['open_count'] = int(state['active'].sum())
        result['closed_count'] = int(state['terminal'].sum())
        result['unclassified_count'] = int(len(self.df) - result['open_count']
                                           - result['closed_count'])

        result['closed_no_date'] = self._flag_count(self.FLAG_CLOSED_NO_DATE)
        result['open_with_date'] = self._flag_count(self.FLAG_OPEN_WITH_DATE)
        result['open_missing_priority'] = self._flag_count(
            self.FLAG_OPEN_MISSING_PRIORITY)

        return result

    def cancelled_states(self) -> Dict[str, Any]:
        """
        The state values in this file that mean cancelled, and the rows holding
        them.

        Returns {
          'state_col': str | None,
          'values':    [str],       # as they actually appear in the data
          'rows':      int,
          'mask':      pd.Series | None,   # True where the state is cancelled
        }

        `mask` is None when there is no state column - which is different from a
        mask of all False, and the caller has to be able to tell "nothing to
        exclude" from "no way to tell", the same rule the skipped logic checks
        follow.

        Folded once per distinct value and matched with isin, not by
        transforming the column: see _masks_by_folded_value for why that
        difference is 0.015s against 0.43s at 300k rows.
        """
        column = self.state_column
        result: Dict[str, Any] = {'state_col': column, 'values': [], 'rows': 0,
                                  'mask': None}
        if not column:
            return result

        cached = self._cache.get('cancelled_states')
        if cached is not None:
            return cached

        raw = self.df[column]
        matched = [value for value in raw.dropna().unique()
                   if any(keyword in _fold(value)
                          for keyword in _CANCELLED_STATE_KEYWORDS)]
        # Sorted for a stable label: the checkbox names these values, and an
        # order that shifts between runs on the same file reads as a change.
        result['values'] = sorted(str(value) for value in matched)
        mask = raw.isin(matched) if matched else pd.Series(False,
                                                          index=self.df.index)
        result['mask'] = mask
        result['rows'] = int(mask.sum())
        self._cache['cancelled_states'] = result
        return result

    def check_state_values(self) -> Dict[str, Any]:
        """
        Enumerate state/status field values and flag low-frequency entries.

        Low-frequency values in a controlled-vocabulary field (state, priority)
        are almost always typos or non-standard entries from integrations.
        A value that accounts for <1% of records AND appears ≤5 times is suspect.
        """
        state_col = self.state_column

        result = {
            'state_col': state_col,
            'has_state': state_col is not None,
            'total_distinct': 0,
            'standard_values': [],  # [{value, count, pct}]
            'suspect_values': [],   # [{value, count, pct}]
            'total_suspect_rows': 0,
        }

        if not state_col:
            return result

        total = len(self.df)
        raw = self.df[state_col]
        # Count the values as they actually appear - the vocabulary is what is
        # being reported, so "Closed" and "closed" are two findings, not one.
        # A missing state is a null-analysis finding, not a vocabulary one, so
        # blanks are dropped first. Blank-but-not-null values (a cell of spaces)
        # are identified from the distinct values rather than by transforming the
        # column; see _masks_by_folded_value.
        blank_values = [v for v in raw.dropna().unique() if not _fold(v)]
        present = raw.dropna()
        if blank_values:
            present = present[~present.isin(blank_values)]
        # value_counts on a categorical reports unobserved categories with a
        # zero count, so the cast to str has to happen first.
        vc = present.astype(str).value_counts()

        result['total_distinct'] = int(len(vc))

        # Ordered explicitly rather than trusting value_counts: it breaks ties by
        # dtype internals, so the same file could otherwise produce two different
        # orderings depending on whether the column arrived as str or category.
        # 'classification' is carried per value so the UI can show which values
        # check_cross_field_logic could not read as open or closed. Without it
        # that check can only report an unclassified *count*, leaving the user
        # no way to find out which values caused it.
        entries = [
            {'value': str(value), 'count': int(count),
             'pct': round(int(count) / total * 100, 2) if total else 0.0,
             'classification': _classify_state(_fold(value))}
            for value, count in vc.items()
        ]
        entries.sort(key=lambda e: (-e['count'], e['value']))

        near_dupes = _find_near_duplicates(entries)

        standard, suspect = [], []
        for entry in entries:
            twin = near_dupes.get(entry['value'])
            if twin is not None:
                suspect.append({**entry, 'reason': f"looks like '{twin}'"})
            elif entry['pct'] < SUSPECT_STATE_PCT and entry['count'] <= SUSPECT_STATE_COUNT:
                suspect.append({**entry, 'reason': 'rare value'})
            else:
                standard.append({**entry, 'reason': ''})

        result['standard_values'] = standard
        result['suspect_values'] = suspect
        result['total_suspect_rows'] = sum(e['count'] for e in suspect)

        return result

    # ==================== Duplicate Check ====================

    def check_duplicates(self) -> Dict[str, Any]:
        """
        Check for duplicate records based on ID column.
        Optimized with value_counts for faster counting.
        Returns: {'id_col': str, 'count': int, 'duplicate_ids': {id: count}}
        """
        id_col = self.id_column
        
        if not id_col:
            return {
                'id_col': None,
                'count': 0,
                'duplicate_ids': {},
                'error': 'Could not identify ID column'
            }
        
        # Use value_counts for efficient duplicate detection
        value_counts = self.df[id_col].value_counts()
        duplicates_mask = value_counts > 1
        
        if duplicates_mask.any():
            duplicate_values = value_counts[duplicates_mask]
            dup_count = int((duplicate_values - 1).sum())  # Total extra occurrences
            duplicate_ids = {str(k): int(v) for k, v in duplicate_values.items()}
        else:
            dup_count = 0
            duplicate_ids = {}
        
        return {
            'id_col': id_col,
            'count': dup_count,
            'duplicate_ids': duplicate_ids
        }
    
    def duplicate_row_mask(self) -> Optional[pd.Series]:
        """
        Boolean mask of every row whose ID value occurs more than once.

        Note this counts differently from check_duplicates()['count'], which is
        the number of *extra* occurrences: a value appearing twice puts 2 rows
        in this mask but adds 1 to that count. Both are correct for their
        purpose - the mask marks rows to inspect, the count says how many rows
        are surplus - so they must not be compared to each other.

        Cached: it is the source for both get_duplicate_rows() and the per-row
        duplicate flag on the report's detail sheet, which have to agree.
        """
        if 'duplicate_row_mask' in self._cache:
            return self._cache['duplicate_row_mask']

        id_col = self.id_column
        mask = (None if not id_col
                else self.df.duplicated(subset=[id_col], keep=False))
        self._cache['duplicate_row_mask'] = mask
        return mask

    def id_occurrence_counts(self) -> Optional[pd.Series]:
        """
        How many times each row's ID value appears, aligned to df.index.

        Maps the value counts back onto the rows rather than counting per row,
        so the cost is one value_counts plus a map regardless of frame size.
        """
        id_col = self.id_column
        if not id_col:
            return None
        counts = self.df[id_col].value_counts()
        return self.df[id_col].map(counts).fillna(0).astype(int)

    def get_duplicate_rows(self) -> Optional[pd.DataFrame]:
        """Get DataFrame of all duplicate rows."""
        mask = self.duplicate_row_mask()
        if mask is None:
            return None
        return self.df[mask]
    
    # ==================== Pivot Tables ====================
    
    def get_pivot(self, column: str) -> Dict[str, Any]:
        """
        Get value counts for a column.
        Optimized with direct numpy operations.
        Returns: {'values': [{value, count, percentage}], 'total_unique': int}

        Cached per column: the pivot view re-requests the selected column on
        every visit, and value_counts over a large free-text column is the
        expensive part.
        """
        if column not in self.df.columns:
            return {'error': f'Column {column} not found', 'values': []}

        pivot_cache = self._cache.setdefault('pivots', {})
        if column in pivot_cache:
            return pivot_cache[column]

        # Use cached total_rows
        total = self.total_rows

        # value_counts is already optimized in pandas.
        # A categorical column reports every *category*, including ones with a
        # zero count that are not present in the data, so drop empties rather
        # than listing values that do not occur.
        vc = self.df[column].value_counts()
        if len(vc):
            vc = vc[vc > 0]
        
        # Build values list efficiently
        counts = vc.values
        indices = vc.index
        
        if total > 0:
            percentages = (counts / total) * 100
        else:
            percentages = np.zeros(len(counts))
        
        values = [
            {
                'value': str(indices[i]),
                'count': int(counts[i]),
                'percentage': round(float(percentages[i]), 2)
            }
            for i in range(len(counts))
        ]

        # value_counts leaves tied counts in an order that depends on the
        # column's dtype (order of appearance for strings, category order for a
        # categorical), which made the same file render two different pivot
        # orderings. Break ties on the value so the output is stable.
        values.sort(key=lambda v: (-v['count'], v['value']))
        
        result = {
            'column': column,
            'values': values,
            'total_unique': len(vc)
        }
        pivot_cache[column] = result
        return result

    def get_pivot_columns(self) -> List[str]:
        """
        Columns offered by the pivot view, in display order (the operationally
        interesting ones first, then the rest alphabetically).

        Names only — deliberately does no counting, so the view can list every
        column instantly and compute just the one the user selects. Building all
        of them up front cost ~4.5s on a 100k x 25 file to display one.
        """
        id_col = self.id_column
        mandatory_keywords = ["priority", "state", "status", "assignment", "tower"]
        col_lower_map = self._cache.get('col_lower_map', {})

        mandatory_cols = [
            c for c in self.df.columns
            if any(k in col_lower_map.get(c, c.lower()) for k in mandatory_keywords)
        ]
        other_cols = sorted([
            c for c in self.df.columns
            if c not in mandatory_cols and c != id_col
        ])
        return mandatory_cols + other_cols

    def get_all_pivots(self) -> Dict[str, Dict[str, Any]]:
        """
        Get pivots for all columns.

        Eager and expensive on a wide file — prefer get_pivot_columns() plus a
        get_pivot() per selection, which is what the pivot view does.
        """
        return {col: self.get_pivot(col) for col in self.get_pivot_columns()}
    
    # ==================== Data Profile ====================
    
    def get_data_profile(self) -> Dict[str, Any]:
        """
        Get overall data profile.
        Optimized with vectorized operations and cached values.
        Returns comprehensive stats about the data.
        """
        profile = {
            'total_rows': self.total_rows,
            'total_columns': self.total_columns,
            'memory_mb': round(self.df.memory_usage(deep=False).sum() / 1024**2, 2),
            'columns': [],
            'date_range': None
        }
        
        # Add date range info if created column exists
        created_dates = self._get_created_dates()
        if created_dates is not None:
            valid_dates = created_dates.dropna()
            if len(valid_dates) > 0:
                min_date = valid_dates.min()
                max_date = valid_dates.max()
                profile['date_range'] = {
                    'column': self.created_column,
                    'first_date': min_date.strftime('%Y-%m-%d'),
                    'last_date': max_date.strftime('%Y-%m-%d'),
                    'first_datetime': min_date.strftime('%Y-%m-%d %H:%M'),
                    'last_datetime': max_date.strftime('%Y-%m-%d %H:%M'),
                }
        
        # Text-bearing columns, for the whitespace check below.
        # Tested per-column rather than via select_dtypes(include=['object']):
        # pandas 3 splits the new 'str' dtype out of 'object', so that call
        # warns now and would silently stop matching string columns later.
        object_cols = {
            col for col in self.df.columns
            if self.df[col].dtype == object
            or pd.api.types.is_string_dtype(self.df[col])
        }
        
        for col in self.df.columns:
            # Report the type of the *values*, not the storage. Large files hold
            # low-cardinality columns as `category` to save memory, which is an
            # internal optimisation - showing the user "category" where another
            # file shows "str" would just look like an inconsistency.
            dtype = self.df[col].dtype
            if isinstance(dtype, pd.CategoricalDtype):
                dtype = dtype.categories.dtype

            col_info = {
                'name': col,
                'dtype': str(dtype),
                'sample': None,
                'has_whitespace': False
            }
            
            # Get sample value (use iloc for speed)
            non_null_mask = self.df[col].notna()
            if non_null_mask.any():
                first_idx = non_null_mask.idxmax()
                col_info['sample'] = str(self.df.loc[first_idx, col])[:50]
            
            # Check for whitespace issues only on object columns
            if col in object_cols:
                col_data = self.df[col].dropna()
                if len(col_data) > 0:
                    # Sample-based check for large columns (faster)
                    sample_size = min(1000, len(col_data))
                    sample = col_data.iloc[:sample_size].astype(str)
                    has_ws = (sample.str.strip() != sample).any()
                    col_info['has_whitespace'] = bool(has_ws)
            
            profile['columns'].append(col_info)
        
        return profile
    
    # ==================== Monthly Inflow ====================
    
    def get_monthly_inflow(self) -> Dict[str, Any]:
        """
        Get monthly inflow data for charting.
        Returns: {'date_col': str, 'data': {period: count}, 'min_date': str, 'max_date': str}
        """
        date_col = self.created_column
        if not date_col:
            return {'error': 'No date column found', 'data': {}}
        
        date_series = self._get_created_dates()
        if date_series is None:
            return {'error': 'Could not parse dates', 'data': {}}
        
        valid_dates = date_series.dropna()
        if valid_dates.empty:
            return {'error': 'No valid dates', 'data': {}}
        
        monthly_counts = valid_dates.dt.to_period('M').value_counts().sort_index()
        
        return {
            'date_col': date_col,
            'data': {str(k): int(v) for k, v in monthly_counts.items()},
            'min_date': valid_dates.min().strftime('%Y-%m-%d'),
            'max_date': valid_dates.max().strftime('%Y-%m-%d')
        }
    
    # ==================== Description Quality ====================

    # COLUMN_KEYWORDS categories that hold free text describing the issue
    _DESCRIPTION_CATEGORIES = ("Short Description", "Description")

    @staticmethod
    def _pct(count: int, total: int) -> float:
        """Percentage of total, guarded against empty data."""
        return round((count / total) * 100, 2) if total else 0.0

    @property
    def description_columns(self) -> List[Tuple[str, str]]:
        """
        Free-text description columns as (category, column) pairs.
        Reuses the scored column_map built at construction time rather than
        re-implementing column detection.
        """
        column_map = self._cache.get('column_map', {})
        pairs = []
        for category in self._DESCRIPTION_CATEGORIES:
            matches = column_map.get(category, [])
            if matches:
                pairs.append((category, matches[0]))
        return pairs

    def _normalize_text_series(self, series: pd.Series,
                               col: Optional[str] = None) -> pd.Series:
        """
        Reduce free text to a comparable key so that two tickets describing the
        same issue collapse together. Strips ticket IDs, dates, numbers, URLs,
        emails and paths (via LanguageChecker's shared pattern list), then
        lowercases, drops punctuation and collapses whitespace.

        Cached per column and shared between columns holding identical text.
        This matters because it is ~94% of the findings register, and Overview
        builds that register on every file open: measured at 100k x 25 the
        register was 1.90s of which this was 1.77s. ServiceNow's
        short_description and description very often agree exactly, so without
        the sharing the same regex pass runs twice for one answer.

        Equality is tested against the raw columns the DataFrame already holds -
        exact, and retaining nothing beyond the results themselves, unlike a
        content hash which would trade a rare silent wrong answer for the same
        saving. A wrong reuse here would merge two columns' recurring-issue
        groups, which is a wrong answer rather than a slow one.

        There is deliberately no separate `col in cache` fast path: the loop
        below already finds the column's own entry, since a column equals
        itself. Two branches that cannot disagree are one branch and a place for
        them to drift.

        Both guards earn their place, and neither is reachable through today's
        single call site - see the tests that drive them directly:
          - `other in self.df.columns` - an entry naming a column this frame no
            longer has cannot be compared, so it must be skipped rather than
            trusted.
          - the length check - identical raw values imply identical populated
            masks, so a caller passing a differently filtered subset is the one
            case where the comparison would otherwise hand back the wrong
            length.
        """
        cache = self._cache.setdefault('normalized_text', {})

        reused = None
        if col is not None and col in self.df.columns:
            for other, cached in cache.items():
                if (other in self.df.columns and len(cached) == len(series)
                        and self.df[other].equals(self.df[col])):
                    reused = cached
                    break

        if reused is None:
            cleaned = LanguageChecker.clean_series(series).str.lower()
            cleaned = cleaned.str.replace(r'[^a-z0-9\s]', ' ', regex=True)
            reused = cleaned.str.replace(r'\s+', ' ', regex=True).str.strip()

        if col is not None:
            cache[col] = reused
        return reused

    # Tuning for the "reworded variants" pass
    _SIMILARITY_THRESHOLD = 0.6   # Jaccard overlap required to call two rows the same issue
    _SIMILAR_MIN_TOKENS = 3       # below this there is not enough signal to compare
    _SIMILAR_MAX_ROWS = 20000     # cap the candidate scan on very large files
    _SIMILAR_MAX_CANDIDATES = 300 # cap comparisons per row

    @staticmethod
    def _significant_tokens(normalized: str) -> set:
        """Topical words of a normalized description, minus filler."""
        return {
            t for t in normalized.split()
            if len(t) >= 4 and t not in DESCRIPTION_STOPWORDS
        }

    def _cluster_similar_issues(self, normalized: pd.Series) -> Tuple[List[List[Any]], bool]:
        """
        Group rows whose topical words overlap above the Jaccard threshold, so
        that differently-worded reports of one issue land together.

        Uses an inverted index over distinctive tokens to generate candidates,
        so rows are never compared pairwise across the whole column.
        Returns (clusters, truncated).
        """
        tokens_by_label: Dict[Any, set] = {}
        truncated = False
        for label, text in normalized.items():
            tokens = self._significant_tokens(text)
            if len(tokens) >= self._SIMILAR_MIN_TOKENS:
                tokens_by_label[label] = tokens
                if len(tokens_by_label) >= self._SIMILAR_MAX_ROWS:
                    truncated = True
                    break

        if len(tokens_by_label) < 2:
            return [], truncated

        # Index only reasonably distinctive tokens - words shared by most rows
        # generate huge candidate lists without discriminating between them.
        doc_freq: Dict[str, int] = {}
        for tokens in tokens_by_label.values():
            for token in tokens:
                doc_freq[token] = doc_freq.get(token, 0) + 1

        freq_cap = max(50, int(len(tokens_by_label) * 0.05))
        postings: Dict[str, List[Any]] = {}
        for label, tokens in tokens_by_label.items():
            for token in tokens:
                if doc_freq[token] <= freq_cap:
                    postings.setdefault(token, []).append(label)

        # Union-find over candidate pairs
        parent = {label: label for label in tokens_by_label}

        def find(x):
            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]
            return x

        for label, tokens in tokens_by_label.items():
            seen = set()
            exhausted = False
            for token in tokens:
                for other in postings.get(token, ()):
                    if other == label or other in seen:
                        continue
                    seen.add(other)
                    if len(seen) > self._SIMILAR_MAX_CANDIDATES:
                        exhausted = True
                        break
                    root_a, root_b = find(label), find(other)
                    if root_a == root_b:
                        continue
                    other_tokens = tokens_by_label[other]
                    union_size = len(tokens | other_tokens)
                    if union_size and len(tokens & other_tokens) / union_size >= self._SIMILARITY_THRESHOLD:
                        parent[root_b] = root_a
                if exhausted:
                    break

        clusters: Dict[Any, List[Any]] = {}
        for label in tokens_by_label:
            clusters.setdefault(find(label), []).append(label)

        return [m for m in clusters.values() if len(m) > 1], truncated

    def _sample_ids(self, index_labels, id_col: Optional[str],
                    limit: int = 5) -> List[str]:
        """Ticket IDs for a group of rows, so findings are actionable."""
        labels = list(index_labels)[:limit]
        if not id_col:
            return [f"row {label}" for label in labels]
        return [str(self.df.loc[label, id_col]) for label in labels]

    def _find_repeated_issues(self, col: str, text: pd.Series,
                              populated_mask: pd.Series, id_col: Optional[str],
                              top_n: int) -> Dict[str, Any]:
        """
        Group rows that describe the same underlying issue, in two tiers:
          - exact:   identical once ticket-specific noise is normalized away
          - similar: same significant-token signature but different wording
        Both tiers are O(n) hash groupbys - no pairwise similarity matrix.
        """
        result = {
            'column': col,
            'exact_group_count': 0,
            'exact_repeated_rows': 0,
            'exact_repeated_pct': 0.0,
            'similar_group_count': 0,
            'similar_grouped_rows': 0,
            'similar_grouped_pct': 0.0,
            'similar_truncated': False,
            'top_groups': []
        }

        populated = text[populated_mask]
        if populated.empty:
            return result

        normalized = self._normalize_text_series(populated, col)
        normalized = normalized[normalized != '']
        if normalized.empty:
            return result

        total = self.total_rows
        groups: List[Dict[str, Any]] = []

        # --- Tier 1: identical after normalization ---
        exact_members = normalized.groupby(normalized).groups
        exact_counts = normalized.value_counts()
        exact_repeats = exact_counts[exact_counts > 1]

        result['exact_group_count'] = int(len(exact_repeats))
        result['exact_repeated_rows'] = int(exact_repeats.sum())
        result['exact_repeated_pct'] = self._pct(result['exact_repeated_rows'], total)

        # Only materialize the largest groups - value_counts is already sorted desc
        for key, count in exact_repeats.head(top_n).items():
            members = exact_members[key]
            groups.append({
                'kind': 'exact',
                'count': int(count),
                'percentage': self._pct(int(count), total),
                'text': str(text.loc[members[0]])[:200],
                'sample_ids': self._sample_ids(members, id_col)
            })

        # --- Tier 2: overlapping wording (catches what tier 1 cannot) ---
        clusters, truncated = self._cluster_similar_issues(normalized)
        result['similar_truncated'] = truncated

        similar_groups = []
        for members in clusters:
            # A cluster is only news if it spans wording tier 1 kept apart
            if normalized.loc[members].nunique() < 2:
                continue
            result['similar_group_count'] += 1
            result['similar_grouped_rows'] += len(members)
            similar_groups.append({
                'kind': 'similar',
                'count': len(members),
                'percentage': self._pct(len(members), total),
                'text': str(text.loc[members[0]])[:200],
                'sample_ids': self._sample_ids(members, id_col)
            })

        result['similar_grouped_pct'] = self._pct(result['similar_grouped_rows'], total)

        similar_groups.sort(key=lambda g: g['count'], reverse=True)
        groups.extend(similar_groups[:top_n])

        groups.sort(key=lambda g: g['count'], reverse=True)
        result['top_groups'] = groups[:top_n]
        return result

    def _description_row_frame(self, col: str, min_length: int,
                               max_length: int) -> DescriptionRowFrame:
        """
        Build (and cache) the per-row arrays for one description column.

        Cached per (column, thresholds) because both check_description_quality
        and get_description_row_scores need it, and the second is called by the
        report for a sheet the first has usually already populated.
        """
        cache_key = f'desc_rows_{col}_{min_length}_{max_length}'
        if cache_key in self._cache:
            return self._cache[cache_key]

        # astype before fillna, not after: filling a *categorical* column
        # with '' raises TypeError unless '' is already a category, so the
        # original order crashed on a low-cardinality description column
        # that contained nulls. Converting to str first drops the
        # categorical constraint; the result is identical either way.
        text = self.df[col].astype(str).fillna('').str.strip()
        lengths = text.str.len()

        empty_mask = lengths == 0
        populated_mask = ~empty_mask

        # group_sizes: how many rows share the same raw text value.
        # Raw value_counts (no normalization) is deliberately used here -
        # it's O(n) and fast, and for scoring purposes the difference
        # between normalized and unnormalized matching is small: a template
        # pasted identically still groups, while a slight number variation
        # ("INC001 ..." vs "INC002 ...") scores individually, which is fine.
        raw_counts = text[populated_mask].value_counts()
        group_sizes = text.map(raw_counts).fillna(0).astype(int)
        group_sizes[empty_mask] = 0

        frame = DescriptionRowFrame(
            text=text,
            lengths=lengths,
            empty_mask=empty_mask,
            populated_mask=populated_mask,
            short_mask=populated_mask & (lengths < min_length),
            long_mask=lengths > max_length,
            group_sizes=group_sizes,
        )
        self._cache[cache_key] = frame
        return frame

    def get_description_row_scores(self, min_length: int = 20,
                                   max_length: int = 5000
                                   ) -> Dict[str, Dict[str, pd.Series]]:
        """
        Per-row quality score and tier for each description column.

        Returns {column: {'score': Series[int], 'tier': Series[str]}}, aligned
        to df.index.

        Deliberately does NOT call check_description_quality: that method's cost
        is the repetition clustering (~3.3s at 300k rows), which a per-row score
        does not need. Both read the same cached DescriptionRowFrame and the same
        _score_description_rows, so the tiers here and the score_distribution
        there are the same arrays rather than two implementations that could
        drift.

        Measured at ~0.31s for two description columns over 300k rows.
        """
        cache_key = f'desc_row_scores_{min_length}_{max_length}'
        if cache_key in self._cache:
            return self._cache[cache_key]

        scores: Dict[str, Dict[str, pd.Series]] = {}
        for _category, col in self.description_columns:
            frame = self._description_row_frame(col, min_length, max_length)
            score, tier = _score_description_rows(
                frame.lengths, frame.empty_mask, frame.short_mask,
                frame.long_mask, frame.group_sizes, min_length)
            scores[col] = {'score': score, 'tier': tier}

        self._cache[cache_key] = scores
        return scores

    def check_description_quality(self, min_length: int = 20,
                                  max_length: int = 5000,
                                  top_n: int = 25) -> Dict[str, Any]:
        """
        Assess whether free-text description fields are usable for analysis.

        Two dimensions:
          1. Length - entries too short to be meaningful or too long to review.
          2. Repetition - tickets describing the same recurring issue, so an
             analytics SME can see the real drivers behind ticket volume.

        Returns: {
            'columns':    [{category, column, total, empty, too_short, too_long,
                            *_pct, min/max/avg/median_length, samples}],
            'repetition': [{column, exact_*, similar_*, top_groups}],
            'issues':     [human-readable strings],
            'min_length': int, 'max_length': int
        }
        """
        cache_key = f'desc_quality_{min_length}_{max_length}_{top_n}'
        if cache_key in self._cache:
            return self._cache[cache_key]

        pairs = self.description_columns
        if not pairs:
            result = {
                'columns': [],
                'repetition': [],
                'issues': [],
                'min_length': min_length,
                'max_length': max_length,
                'error': 'Could not identify a Short Description or Description column'
            }
            self._cache[cache_key] = result
            return result

        total = self.total_rows
        id_col = self.id_column
        columns_out: List[Dict[str, Any]] = []
        repetition_out: List[Dict[str, Any]] = []
        issues: List[str] = []

        row_scores = self.get_description_row_scores(min_length, max_length)

        for category, col in pairs:
            frame = self._description_row_frame(col, min_length, max_length)
            text = frame.text
            lengths = frame.lengths
            empty_mask = frame.empty_mask
            populated_mask = frame.populated_mask
            short_mask = frame.short_mask
            long_mask = frame.long_mask
            populated_lengths = lengths[populated_mask]
            has_text = len(populated_lengths) > 0

            metrics = {
                'category': category,
                'column': col,
                'total': total,
                'empty': int(empty_mask.sum()),
                'too_short': int(short_mask.sum()),
                'too_long': int(long_mask.sum()),
                'min_length': int(populated_lengths.min()) if has_text else 0,
                'max_length': int(populated_lengths.max()) if has_text else 0,
                'avg_length': round(float(populated_lengths.mean()), 1) if has_text else 0.0,
                'median_length': int(populated_lengths.median()) if has_text else 0,
                'samples': {
                    'too_short': text[short_mask].head(5).tolist(),
                    'too_long': [t[:120] + '...' for t in text[long_mask].head(3)]
                }
            }
            metrics['empty_pct'] = self._pct(metrics['empty'], total)
            metrics['too_short_pct'] = self._pct(metrics['too_short'], total)
            metrics['too_long_pct'] = self._pct(metrics['too_long'], total)

            if metrics['empty']:
                issues.append(f"{metrics['empty']} rows have an empty {col}")
            if metrics['too_short']:
                issues.append(
                    f"{metrics['too_short']} rows where {col} is under "
                    f"{min_length} characters"
                )
            if metrics['too_long']:
                issues.append(
                    f"{metrics['too_long']} rows where {col} exceeds "
                    f"{max_length} characters"
                )

            repetition = self._find_repeated_issues(
                col, text, populated_mask, id_col, top_n
            )
            repetition_out.append(repetition)

            # --- Overall quality score ---
            # The very same tier array the report writes per row, so the
            # distribution here and the Tier column there cannot disagree.
            tiers = row_scores[col]['tier']
            distribution = {
                tier: int((tiers == tier).sum())
                for tier in _SCORE_TIER_ORDER
            }
            # Weighted average of per-row tier weights → single 0-100 field score.
            metrics['overall_score'] = int(
                sum(distribution[t] * _TIER_WEIGHTS[t] for t in _SCORE_TIER_ORDER) // max(total, 1)
            )
            metrics['score_distribution'] = distribution

            columns_out.append(metrics)

            if repetition['exact_group_count']:
                issues.append(
                    f"{repetition['exact_repeated_rows']} rows in {col} repeat "
                    f"{repetition['exact_group_count']} recurring issues "
                    f"({repetition['exact_repeated_pct']}% of all rows)"
                )
            if repetition['similar_group_count']:
                issues.append(
                    f"{repetition['similar_grouped_rows']} rows in {col} cluster into "
                    f"{repetition['similar_group_count']} differently-worded issues"
                )

        result = {
            'columns': columns_out,
            'repetition': repetition_out,
            'issues': issues,
            'min_length': min_length,
            'max_length': max_length
        }
        self._cache[cache_key] = result
        return result

    # ==================== Summary / Overview ====================

    def get_summary(self, null_threshold: float = 20.0) -> Dict[str, Any]:
        """
        Get a complete summary of all checks.
        """
        column_check = self.check_columns()
        null_check = self.check_nulls(null_threshold)
        date_check = self.check_date_logic()
        dup_check = self.check_duplicates()
        
        issues = []
        
        # Missing columns
        missing = [r for r, d in column_check.items() if not d['found']]
        if missing:
            issues.append(f"{len(missing)} missing required columns")
        
        # High nulls
        high_nulls = [c for c, d in null_check.items() if d['above_threshold']]
        if high_nulls:
            issues.append(f"{len(high_nulls)} columns exceed {null_threshold}% null threshold")
        
        # Duplicates
        if dup_check['count'] > 0:
            issues.append(f"{dup_check['count']} duplicate records found")
        
        # Date issues
        if date_check['issues']:
            issues.extend(date_check['issues'])
        
        return {
            'total_rows': self.total_rows,
            'total_columns': self.total_columns,
            'missing_columns': missing,
            'high_null_columns': high_nulls,
            'duplicate_count': dup_check['count'],
            'date_issues': date_check['issues'],
            'issues': issues,
            'has_issues': len(issues) > 0
        }
