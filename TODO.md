# TODO

Planned work, in order. Two features plus the fallout found while planning them.
Each item records *why*, because the reasoning is the part that gets lost.

Not committed deliberately — this is a working list, not a shipped document.

---

## A. Per-row language columns

**The feature already half-exists, which is the main thing to know before
starting.** `_row_language_columns()` (`core/reporter.py:1391`) already puts a
`{col} Language` column on the report's Ticket Detail sheet for every column
Language Check has been run on, at **zero extra detection cost** —
`row_languages` is a free by-product of the check that already ran, and
`tests/test_reporter.py:832` spies `detect_language_series` to assert it is called
**zero** times while building the detail. It already distinguishes four states:
`Not checked` / `Too short to check` / `English` / the language.

So the gaps are narrower than "build per-row language output":

- [ ] **A1 — Widen the answer from one column to three.** Confidence and tier are
  computed inside `check_column` and thrown away (`samples` keeps 5-tuples for the
  top ten rows only; `row_languages` keeps `{index: lang}`). Emitting
  `{col} Confidence` and `{col} Detected by` means widening `row_languages` to
  carry the tuple — **not** re-detecting. `analyzed_mask` and the existing key set
  must stay backward-compatible: `_row_language_columns` is deliberately defensive
  about hand-built dicts (`tests/test_reporter.py:819`).
  - Converge on the **shipped** naming, `{col} Language` with a space. The dead
    `add_language_columns_to_df` uses `{col}_Language`; two conventions for one
    thing is how they drift.
  - Move the tier labels into `core/language.py`. There are two vocabularies today
    (`_METHOD_REPORT_LABELS` at `reporter.py:23`, `_LANG_METHOD_LABELS` at
    `app_pyside.py:2103`) and Extract would make a third.
  - Add a **core-side** confidence renderer. `_lang_confidence_text` is GUI-only,
    and the rule that *confidence is only comparable within a tier* has to hold in
    a written file too: a script hit renders `definitive`, a keyword hit
    `keyword match` — never `1.00` and `0.90` sitting beside Lingua's `94%`.
  - Delete `_build_data_with_languages` (`reporter.py:816`) and
    `get_language_columns` (`language.py:546`). Both are unreachable, and
    `_build_data_with_languages` is actively wrong: it recomputes, **ignores
    `lang_mode`** (silently `mode='all'`), and uses `min_confidence=0.75` where
    the live path uses `0.50`, so reviving it would produce output that disagrees
    with the screen.

- [ ] **A2 — Two latent defects in detection.** Pre-existing, not caused by the
  above, cheap to fix while in here:
  - `_detection_cache` keys on **`text[:200]`** (`language.py:288`). Two different
    descriptions sharing a 200-character prefix — routine in templated ticket
    text — collide, and the second silently inherits the first's verdict.
  - The cache is capped at 10,000 entries and **stops inserting** past that rather
    than evicting, so a high-cardinality column gets no hits beyond the cap.
  - **`check_column` does not dedupe.** It iterates every analysed row relying
    solely on that capped cache, while `detect_language_series` dedupes by
    distinct text and is pinned for it (`tests/test_language.py:397`). Above 10k
    distinct texts `check_column` pays the ~1.2ms-per-text cost *per row* instead
    of per distinct value.

- [ ] **A3 — The Extract option.** `core/extract.py` has no language awareness at
  all; this is the actual ask. A column picker limited to what
  `self.language_results` already holds, **naming the unchecked columns as
  unavailable and why** — a missing column in a written file is
  indistinguishable from a clean one. The three columns append to the output; the
  mode goes on the info sheet and the CSV sidecar (which tiers *ran*, not which
  are installed). No new detection, so no progress bar — that is the whole point
  of reusing the results.
  - `lingua_only` with Lingua absent must **not** write a column of blanks.
    `check_column` already returns a `warning` for it and `_row_language_columns`
    already maps that to `Not checked` for every row; Extract must do the same
    rather than omitting the column.

**Deliberately out of scope: filtering the output to non-English rows only.** It
would make Extract's row count unknowable until detection had run, and that count
is exact and synchronous by design (`_update_extract_plan` is sync precisely so it
cannot paint a number for a selection already abandoned). The written Language
column filters in Excel in two clicks.

---

## B. Combine — append many sheets and files into one table

ITSM exports arrive as several sheets in one workbook or several monthly
workbooks, and they do not agree: the same field is `Number` on one and
`Ticket Number` on another, some sheets lack columns others have, some carry
extra custom fields, the same field is `"3 - Moderate"` here and `3` there, and
some sheets have a title block above the real header. Excel cannot do this join.
TicketAudit reads **one sheet at a time** today (`read_table(path, sheet_name)`,
`getOpenFileName` singular, no `pd.concat` on the load path).

**It is a cached dialog, not a fifteenth nav view.** The nav column is full: the
row and caption heights are tuned so all fourteen views fit the 14-inch laptop in
465px, and `test_every_view_is_reachable_without_scrolling_on_a_14_inch_laptop`
fails on a fifteenth. So it follows `Settings`'s `NON_NAV_VIEWS` pattern — which
also means it works with **no file open**, since `show_view`'s `df is None`
early-return never applies.

- [ ] **B1 — `core/loader.py`: three backward-compatible additions.** All
  keyword-only, all defaulting to today's behaviour so the existing `TestExcel` /
  `TestSheetListing` assertions are untouched.
  - `read_table(path, sheet_name=None, *, header=0, optimize=True)` — `header`
    threaded through `_read_excel_with_fallback:184`, `read_excel_streaming:115`
    and the `read_csv` branch; `optimize` gates `optimize_dtypes`.
  - `peek_rows(path, sheet_name=None, *, limit=25)` — cheap rows for header
    detection. **Must use the openpyxl read-only path**, not calamine: `nrows=`
    does not make calamine cheap (it parses the whole sheet — 691 MB for 5000
    rows), while `read_excel_streaming` already streams genuinely.
  - `estimate_rows(path, sheet_name=None)` — from the sheet dimension, `None` when
    unknown. Explicitly an estimate.

- [ ] **B2 — `core/combine.py` + `tests/test_combine.py`.** No GUI. `NamedTuple`s
  (`Source`, `ColumnGroup`, `CombinePlan`) matching `ExtractPlan`'s style, and
  `CombineTooLarge` / `CombineCancelled` as **siblings, not a hierarchy**.
  - **Column alignment, three tiers, conservative.** The destructive error is
    merging two different fields into one column — unrecoverable from the output,
    where keeping them apart just gives two sparse columns you can see. So:
    normalised-exact acts **automatically** (case, separators); keyword-category
    via `config.COLUMN_KEYWORDS` only **proposes** (catches `Number` vs `INC No`,
    which share almost no characters, and stays silent when a source offers two
    candidates); fuzzy `difflib` is **off by default** (`Assigned To` vs
    `Assignment Group` scores high and they are different fields).
  - **Streaming write** — read one source, reindex to the output columns, add
    provenance, dedupe, ceiling-check, write, `del`. Never holds the combined
    table; concatenation roughly doubles peak.
  - **`optimize=False` per source.** Concatenating `category` columns with
    differing category sets silently degrades to `object`, so per-sheet
    optimisation both wastes work and produces mixed dtypes.
  - **Dedupe holds only the id column's values** — a set of strings, affordable at
    any scale. Keeps the **first** occurrence in source order and says so, which
    is why the source list is reorderable. "Keep last" is not offered: it needs a
    second full read of every source, and reversing the list is free.
  - **Header-row detection returns a row *and a reason*** — "row 4: 22 of 22 cells
    are distinct text, and row 5 holds dates" — because a count nobody can trace
    is not actionable. Every decision shown and overridable.
  - Only `table`-classified sources ticked by default; `empty` and `not-a-table`
    listed **with their reason** rather than hidden.
  - `write_only=True` is available (no PivotTable), so `freeze_panes` and column
    widths must be set **before the first `append`**, and the temp-file cleanup
    must close `worksheet._rows` *before* `writer.cleanup()`.

- [ ] **B3 — The dialog.** Sources card, Columns card beside it, preview full
  width, Result card **outside** the scroll area so the write button and the size
  it is about to write stay visible. Progress and cancel via
  `run_in_background(..., report_progress=True)`; **not** `_refresh_async`, which
  cannot report progress.

- [ ] **B4 — Merge review, dedupe, language columns on Combine.** Suggestion
  review modelled on `show_column_mapping_dialog` (`app_pyside.py:4633`). Persist
  `combine_format`, `combine_dedupe`, `combine_fuzzy`.

- [ ] **B5 — Measure, then document.** Scan cost for 12 sheets, write rate, and
  **peak RSS against a combined table that would not fit** — the one load-bearing
  performance claim here. Then the CLAUDE.md section.

### Honesty surfaces (the ones that are easy to get wrong)

- **A blank from an absent column is not a blank cell.** Sheet B lacking
  `Resolved` gives every B row `NaN`, indistinguishable from "not yet resolved".
  Mandatory `Source File` / `Source Sheet` columns make it recoverable, and the
  info sheet carries a per-source × per-column presence grid.
- **Type conflicts are reported, not coerced.** `"3 - Moderate"` beside `3` is a
  real defect in the source; coercing hides the finding this tool exists to find,
  and Excel has no column types so writing as-is loses nothing.
- **The live row count is an estimate and says so** — it comes from the sheet
  dimension, which openpyxl can report stale. The plan reads `~N rows
  (estimated)`; the completion dialog reports the exact figure.
- **Never truncate.** Over Excel's ceiling, raise and keep CSV offered.
- **Cancelling deletes the part-written file.**

---

## C. Batch QA across files

- [ ] Run `_build_findings_df()` over several files into one comparison table.
  Shares only the multi-file picker with Combine, so it ships independently.

---

## D. CLAUDE.md is wrong in three places

Found while planning the above. Each sends a reader after something that is not
there:

- [ ] *"`config.json`'s `report_inline_max_rows` (100k)"* — **no such config
  key.** It is `DEFAULT_INLINE_MAX_ROWS` (`excel_report.py:103`), a constructor
  parameter (`inline_max_rows`) that `ReportGenerator.export_to_excel`
  (`reporter.py:593`) **never forwards**. The 100k CSV-sidecar threshold is
  currently unreachable from the GUI or config — a hard-coded default, not a
  setting. Either plumb it or stop calling it one.
- [ ] *"`Data with Languages` stays gated on `lang_columns`"* — **that sheet does
  not exist.** `_build_data_with_languages` has no callers and the sheet list in
  `excel_report.py:571` never mentions it. Per-row languages ship on **Ticket
  Detail**.
- [ ] `lang_columns` is credited as the per-row opt-in. It is reachable only from
  Python and tests (zero hits in `gui/app_pyside.py`), and `lang_results` **wins
  over it** in `_language_results()` (`reporter.py:124`) — passing both makes
  `lang_columns` a no-op.
