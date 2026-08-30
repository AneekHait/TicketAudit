"""
Tests for core/loader.py — file format, encoding and delimiter handling.
"""
import os

import pytest
import pandas as pd
from core import loader


ROWS = [("INC0001", "Password reset for user"), ("INC0002", "VPN keeps dropping")]


def _write(path, text, encoding="utf-8"):
    path.write_bytes(text.encode(encoding))
    return str(path)


def _csv_text(delimiter=",", header=("number", "short_description")):
    lines = [delimiter.join(header)]
    lines += [delimiter.join(row) for row in ROWS]
    return "\n".join(lines) + "\n"


# ──────────────────────────────────────────────────────────────
# Delimiters
# ──────────────────────────────────────────────────────────────

class TestDelimiters:
    @pytest.mark.parametrize("delimiter,suffix", [
        (",", ".csv"),
        (";", ".csv"),   # regional / European ITSM exports
        ("\t", ".tsv"),
        ("|", ".txt"),
    ])
    def test_delimiter_detected(self, tmp_path, delimiter, suffix):
        path = _write(tmp_path / f"tickets{suffix}", _csv_text(delimiter))
        df = loader.read_table(path)
        assert list(df.columns) == ["number", "short_description"]
        assert len(df) == 2
        assert df.iloc[0]["number"] == "INC0001"

    def test_commas_inside_quoted_text_do_not_break_semicolon_file(self, tmp_path):
        text = (
            "number;short_description\n"
            'INC0001;"Reset password, then reboot, please"\n'
            'INC0002;"Network down, users affected"\n'
        )
        path = _write(tmp_path / "quoted.csv", text)
        df = loader.read_table(path)
        assert list(df.columns) == ["number", "short_description"]
        assert df.iloc[0]["short_description"] == "Reset password, then reboot, please"

    def test_single_column_file_falls_back_to_comma(self, tmp_path):
        path = _write(tmp_path / "one.csv", "number\nINC0001\nINC0002\n")
        df = loader.read_table(path)
        assert list(df.columns) == ["number"]
        assert len(df) == 2


# ──────────────────────────────────────────────────────────────
# Encodings
# ──────────────────────────────────────────────────────────────

class TestEncodings:
    ACCENTED = "number,short_description\nINC0001,café Müller naïve\n"

    def test_plain_utf8(self, tmp_path):
        path = _write(tmp_path / "u8.csv", self.ACCENTED, "utf-8")
        df = loader.read_table(path)
        assert df.iloc[0]["short_description"] == "café Müller naïve"

    def test_utf8_with_bom_strips_bom_from_first_column_name(self, tmp_path):
        """A BOM left in place corrupts the first column name and breaks column matching."""
        path = _write(tmp_path / "bom.csv", self.ACCENTED, "utf-8-sig")
        df = loader.read_table(path)
        assert list(df.columns)[0] == "number", "BOM was not stripped"
        assert df.iloc[0]["short_description"] == "café Müller naïve"

    def test_cp1252(self, tmp_path):
        path = _write(tmp_path / "cp.csv", self.ACCENTED, "cp1252")
        df = loader.read_table(path)
        assert df.iloc[0]["short_description"] == "café Müller naïve"

    def test_detect_encoding_prefers_utf8(self, tmp_path):
        path = _write(tmp_path / "u8.csv", self.ACCENTED, "utf-8")
        assert loader.detect_encoding(path) == "utf-8-sig"

    def test_undecodable_bytes_still_load(self, tmp_path):
        """latin-1 is the terminal fallback: the app must never refuse to open a file."""
        path = tmp_path / "weird.csv"
        path.write_bytes(b"number,short_description\nINC0001,\x81\x8d\x8f raw bytes\n")
        df = loader.read_table(path)
        assert len(df) == 1


# ──────────────────────────────────────────────────────────────
# Excel
# ──────────────────────────────────────────────────────────────

class TestExcel:
    def _workbook(self, tmp_path):
        path = tmp_path / "book.xlsx"
        with pd.ExcelWriter(path) as writer:
            pd.DataFrame({"number": ["INC1", "INC2"]}).to_excel(
                writer, sheet_name="Incidents", index=False)
            pd.DataFrame({"number": ["CHG1"]}).to_excel(
                writer, sheet_name="Changes", index=False)
        return str(path)

    def test_list_sheets_returns_names(self, tmp_path):
        assert loader.list_sheets(self._workbook(tmp_path)) == ["Incidents", "Changes"]

    def test_read_named_sheet(self, tmp_path):
        df = loader.read_table(self._workbook(tmp_path), sheet_name="Changes")
        assert len(df) == 1
        assert df.iloc[0]["number"] == "CHG1"

    def test_read_without_sheet_name_returns_first_sheet_not_a_dict(self, tmp_path):
        """read_excel(sheet_name=None) would return a dict of all sheets."""
        df = loader.read_table(self._workbook(tmp_path))
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 2

    def test_is_excel(self, tmp_path):
        assert loader.is_excel(self._workbook(tmp_path))


class TestExcelEngine:
    """
    .xlsx reading dominates load time, so loader prefers calamine (~6x faster
    than openpyxl on a 100k-row export). It must stay an accelerator: a missing
    or failing engine may only cost speed, never break opening a file.
    """

    def _workbook(self, tmp_path, name="book.xlsx"):
        path = tmp_path / name
        pd.DataFrame({
            "number": ["INC1", "INC2"],
            "opened_at": pd.to_datetime(["2024-01-01", "2024-02-01"]),
            "count": [1, 2],
            "text": ["café Müller", "plain"],
        }).to_excel(path, index=False)
        return str(path)

    @pytest.fixture(autouse=True)
    def _clear_engine_cache(self):
        loader.excel_engine.cache_clear()
        yield
        loader.excel_engine.cache_clear()

    def test_engine_is_none_when_calamine_missing(self, monkeypatch):
        import builtins
        real_import = builtins.__import__

        def no_calamine(name, *args, **kwargs):
            if name == "python_calamine":
                raise ImportError("not installed")
            return real_import(name, *args, **kwargs)

        monkeypatch.setattr(builtins, "__import__", no_calamine)
        assert loader.excel_engine() is None, \
            "must fall back to pandas' default engine, not a bogus name"

    def test_file_still_loads_without_the_fast_engine(self, tmp_path, monkeypatch):
        path = self._workbook(tmp_path)
        monkeypatch.setattr(loader, "excel_engine", lambda: None)
        df = loader.read_table(path)
        assert len(df) == 2
        assert loader.list_sheets(path) == ["Sheet1"]

    def test_a_broken_engine_does_not_stop_the_file_opening(self, tmp_path, monkeypatch):
        """A file the accelerator cannot parse must be retried, not rejected."""
        path = self._workbook(tmp_path)
        monkeypatch.setattr(loader, "excel_engine", lambda: "no_such_engine")
        df = loader.read_table(path)
        assert len(df) == 2
        assert loader.list_sheets(path) == ["Sheet1"]

    def test_engine_choice_does_not_change_the_data(self, tmp_path, monkeypatch):
        """Speed must not cost fidelity — values and dtypes must match."""
        path = self._workbook(tmp_path)

        fast = loader.read_table(path)
        monkeypatch.setattr(loader, "excel_engine", lambda: None)
        default = loader.read_table(path)

        assert list(fast.columns) == list(default.columns)
        assert fast.shape == default.shape
        assert [str(d) for d in fast.dtypes] == [str(d) for d in default.dtypes]
        assert fast.astype(str).equals(default.astype(str))


# ──────────────────────────────────────────────────────────────
# Low-memory streaming fallback
# ──────────────────────────────────────────────────────────────

class TestExcelStreamingFallback:
    """
    Every whole-sheet Excel reader holds its own full copy before pandas sees
    it: 300k x 25 peaked at ~1050 MB to produce a 40 MB frame. The streaming
    reader halves that (~500 MB) for ~5x the time, so it is a last resort for
    a file that otherwise cannot be opened at all.

    Its one hard requirement is fidelity - reader choice may change speed and
    memory but never what the file appears to contain.
    """

    def _write(self, tmp_path, frame, name="book.xlsx"):
        path = tmp_path / name
        frame.to_excel(path, index=False)
        return str(path)

    def _fast(self, path, sheet=0):
        return pd.read_excel(path, sheet_name=sheet,
                             engine=loader.excel_engine() or None)

    @pytest.mark.parametrize("frame,label", [
        (pd.DataFrame({"n": [f"INC{i:04d}" for i in range(6)],
                       "p": ["P1", "P2"] * 3,
                       "when": pd.date_range("2024-01-01", periods=6),
                       "num": range(6),
                       "flag": [True, False] * 3,
                       "cost": [1.5, 2.25] * 3}), "mixed types"),
        (pd.DataFrame({"a": [1, None, 3, None],
                       "b": ["x", None, None, "y"]}), "interior nulls"),
        (pd.DataFrame({"id": [1, 2, 3], "empty": [None, None, None]}), "all-null column"),
        (pd.DataFrame({"t": ["café Müller", "  padded  ", "密码重置"]}), "unicode"),
        (pd.DataFrame({"only": [1]}), "single cell"),
        (pd.DataFrame({"code": ["007", "010", "0900"]}), "numeric-looking text"),
        (pd.DataFrame({f"c{i}": range(4) for i in range(25)}), "wide"),
    ])
    def test_matches_the_normal_reader(self, tmp_path, frame, label):
        path = self._write(tmp_path, frame)
        fast = self._fast(path)
        streamed = loader.read_excel_streaming(path, 0)

        assert list(streamed.columns) == list(fast.columns), label
        assert streamed.shape == fast.shape, label
        assert [str(d) for d in streamed.dtypes] == [str(d) for d in fast.dtypes], label
        assert streamed.astype(str).equals(fast.astype(str)), label

    def test_interior_empty_row_is_kept(self, tmp_path):
        """
        Dropping every all-empty row silently shortens the frame. Only trailing
        ones (from a stale sheet dimension) may go.
        """
        frame = pd.DataFrame({"a": [1, None, 3], "b": ["x", None, "z"]})
        path = self._write(tmp_path, frame)
        assert len(loader.read_excel_streaming(path, 0)) == 3

    def test_header_only_sheet(self, tmp_path):
        path = self._write(tmp_path, pd.DataFrame({"a": [], "b": []}))
        streamed = loader.read_excel_streaming(path, 0)
        assert list(streamed.columns) == ["a", "b"]
        assert len(streamed) == 0

    def test_named_sheet_is_selected(self, tmp_path):
        path = tmp_path / "multi.xlsx"
        with pd.ExcelWriter(path) as writer:
            pd.DataFrame({"a": [1, 2]}).to_excel(writer, sheet_name="First", index=False)
            pd.DataFrame({"b": ["x"]}).to_excel(writer, sheet_name="Second", index=False)
        streamed = loader.read_excel_streaming(str(path), "Second")
        assert list(streamed.columns) == ["b"]
        assert len(streamed) == 1

    def test_first_sheet_when_unspecified(self, tmp_path):
        path = tmp_path / "multi.xlsx"
        with pd.ExcelWriter(path) as writer:
            pd.DataFrame({"a": [1, 2]}).to_excel(writer, sheet_name="First", index=False)
            pd.DataFrame({"b": ["x"]}).to_excel(writer, sheet_name="Second", index=False)
        assert list(loader.read_excel_streaming(str(path)).columns) == ["a"]

    def test_out_of_memory_triggers_the_fallback(self, tmp_path, monkeypatch):
        path = self._write(tmp_path, pd.DataFrame({"a": [1, 2, 3]}))
        calls = []

        def oom(*a, **k):
            calls.append("read_excel")
            raise MemoryError()

        monkeypatch.setattr(loader.pd, "read_excel", oom)
        df = loader.read_table(path)
        assert calls, "pd.read_excel should have been attempted first"
        assert list(df["a"]) == [1, 2, 3], "streaming reader should have supplied the data"

    def test_message_is_actionable_when_streaming_also_fails(self, tmp_path, monkeypatch):
        path = self._write(tmp_path, pd.DataFrame({"a": [1]}))
        monkeypatch.setattr(loader.pd, "read_excel",
                            lambda *a, **k: (_ for _ in ()).throw(MemoryError()))
        monkeypatch.setattr(loader, "read_excel_streaming",
                            lambda *a, **k: (_ for _ in ()).throw(MemoryError()))
        with pytest.raises(MemoryError, match="CSV"):
            loader.read_table(path)

    def test_file_handle_is_released(self, tmp_path):
        """read_only workbooks keep the file open unless closed explicitly."""
        path = self._write(tmp_path, pd.DataFrame({"a": [1, 2]}))
        loader.read_excel_streaming(path, 0)
        os.replace(path, path + ".moved")   # fails on Windows if still open


# ──────────────────────────────────────────────────────────────
# Memory
# ──────────────────────────────────────────────────────────────

class TestOptimizeDtypes:
    """
    Low-cardinality string columns become `category` on large frames: a 300k-row
    export went 341 MB -> 109 MB. The cardinality guard is the whole safety
    mechanism — a category dtype on a unique key is *bigger* than the strings.
    """

    def _frame(self, rows):
        return pd.DataFrame({
            "number": [f"INC{i:07d}" for i in range(rows)],          # unique
            "priority": [["P1", "P2", "P3"][i % 3] for i in range(rows)],
            "state": [["Open", "Closed"][i % 2] for i in range(rows)],
            "count": [i % 7 for i in range(rows)],                    # int
            "opened": pd.to_datetime(["2024-01-01"] * rows),
        })

    def _big(self):
        return self._frame(loader.MIN_ROWS_TO_OPTIMIZE + 10)

    def test_low_cardinality_columns_become_category(self):
        out = loader.optimize_dtypes(self._big())
        assert isinstance(out["priority"].dtype, pd.CategoricalDtype)
        assert isinstance(out["state"].dtype, pd.CategoricalDtype)

    def test_unique_key_column_is_left_alone(self):
        """A category dtype on a unique column costs more memory than it saves."""
        out = loader.optimize_dtypes(self._big())
        assert not isinstance(out["number"].dtype, pd.CategoricalDtype)

    def test_non_string_columns_untouched(self):
        out = loader.optimize_dtypes(self._big())
        assert out["count"].dtype == "int64"
        assert str(out["opened"].dtype).startswith("datetime64")

    def test_small_frames_are_not_converted(self):
        """Keeps the saving where it matters and small fixtures predictable."""
        out = loader.optimize_dtypes(self._frame(50))
        assert not any(isinstance(out[c].dtype, pd.CategoricalDtype)
                       for c in out.columns)

    def test_it_actually_saves_memory(self):
        before = self._big()
        after = loader.optimize_dtypes(before.copy())
        assert after.memory_usage(deep=True).sum() < \
            before.memory_usage(deep=True).sum()

    def test_values_are_unchanged(self):
        before = self._big()
        after = loader.optimize_dtypes(before.copy())
        for col in before.columns:
            assert list(after[col].astype(str)) == list(before[col].astype(str))

    def test_nulls_survive_conversion(self):
        rows = loader.MIN_ROWS_TO_OPTIMIZE + 10
        df = pd.DataFrame({"state": [None if i % 5 == 0 else "Open"
                                     for i in range(rows)]})
        out = loader.optimize_dtypes(df)
        assert isinstance(out["state"].dtype, pd.CategoricalDtype)
        assert out["state"].isna().sum() == len(range(0, rows, 5))

    def test_idempotent(self):
        once = loader.optimize_dtypes(self._big())
        twice = loader.optimize_dtypes(once.copy())
        assert isinstance(twice["priority"].dtype, pd.CategoricalDtype)

    def test_none_frame_is_tolerated(self):
        assert loader.optimize_dtypes(None) is None

    def test_read_table_applies_it(self, tmp_path):
        rows = loader.MIN_ROWS_TO_OPTIMIZE + 10
        path = tmp_path / "big.csv"
        lines = ["number,state"] + [
            f"INC{i:07d},{'Open' if i % 2 else 'Closed'}" for i in range(rows)]
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")

        df = loader.read_table(path)
        assert len(df) == rows
        assert isinstance(df["state"].dtype, pd.CategoricalDtype)
        assert not isinstance(df["number"].dtype, pd.CategoricalDtype)


# ──────────────────────────────────────────────────────────────
# Sheets vs single-table
# ──────────────────────────────────────────────────────────────

class TestSheetListing:
    def test_csv_reports_no_sheets(self, tmp_path):
        path = _write(tmp_path / "t.csv", _csv_text())
        assert loader.list_sheets(path) == []
        assert loader.is_excel(path) is False

    def test_extension_that_lies_is_detected_by_content(self, tmp_path):
        """An .xlsx renamed to .csv should still be recognised as Excel."""
        real = tmp_path / "real.xlsx"
        pd.DataFrame({"number": ["INC1"]}).to_excel(real, index=False)
        renamed = tmp_path / "actually_excel.csv"
        renamed.write_bytes(real.read_bytes())
        # .csv is a known text extension, so content is not consulted - document that
        assert loader.is_excel(renamed) is False

        no_ext = tmp_path / "mystery"
        no_ext.write_bytes(real.read_bytes())
        assert loader.is_excel(no_ext) is True
        assert len(loader.read_table(no_ext)) == 1


# ──────────────────────────────────────────────────────────────
# Degenerate input
# ──────────────────────────────────────────────────────────────

class TestEmptyAndOdd:
    def test_empty_file_raises_clear_message(self, tmp_path):
        path = _write(tmp_path / "empty.csv", "")
        with pytest.raises(ValueError, match="empty"):
            loader.read_table(path)

    def test_header_only_file_loads_with_zero_rows(self, tmp_path):
        path = _write(tmp_path / "head.csv", "number,short_description\n")
        df = loader.read_table(path)
        assert len(df) == 0
        assert list(df.columns) == ["number", "short_description"]

    def test_missing_file_raises(self, tmp_path):
        with pytest.raises((OSError, ValueError)):
            loader.read_table(tmp_path / "nope.csv")


# ──────────────────────────────────────────────────────────────
# Unreadable workbooks
# ──────────────────────────────────────────────────────────────


def _truncate(path, keep):
    """A file cut off after `keep` bytes, like an interrupted download."""
    data = path.read_bytes()
    assert len(data) > keep, "fixture is too small to truncate meaningfully"
    path.write_bytes(data[:keep])
    return str(path)


def _workbook(path):
    pd.DataFrame({"number": ["INC0001"], "state": ["Open"]}).to_excel(
        path, index=False)
    return path


class TestUnreadableWorkbooks:
    """
    An .xlsx *is* a zip, and a zip keeps its index at the end. So a download
    that stops early leaves a file with a perfect header, a readable first
    member, and no index - and zipfile reports that as "File is not a zip
    file", which reads as "wrong format" to anyone who does not know that.

    Reproduced from two real interrupted downloads of a 230MB ServiceNow
    export, at 31.6MB and 47.8MB. The format was never the problem; there was
    simply less of the file than there should have been.
    """

    def test_a_truncated_workbook_is_called_incomplete(self, tmp_path):
        path = _workbook(tmp_path / "cut.xlsx")
        full = path.stat().st_size
        _truncate(path, full // 2)

        with pytest.raises(ValueError) as caught:
            loader.list_sheets(str(path))
        message = str(caught.value)
        assert "incomplete" in message.lower()
        assert "not a zip" not in message.lower(), \
            "the raw zipfile wording is what this replaced"

    def test_the_message_names_the_size_it_found(self, tmp_path):
        """
        The size is the actionable part: it is what the user compares against
        their source. A count nobody can trace is not actionable.
        """
        path = _workbook(tmp_path / "cut.xlsx")
        _truncate(path, path.stat().st_size // 2)
        with pytest.raises(ValueError, match=r"MB"):
            loader.list_sheets(str(path))

    def test_read_table_also_explains_itself(self, tmp_path):
        """read_table is a public entry point and is reachable without
        list_sheets, so it cannot rely on that call having run first."""
        path = _workbook(tmp_path / "cut.xlsx")
        _truncate(path, path.stat().st_size // 2)
        with pytest.raises(ValueError) as caught:
            loader.read_table(str(path))
        assert "not a zip" not in str(caught.value).lower()

    def test_an_empty_workbook_says_it_is_empty(self, tmp_path):
        path = tmp_path / "nothing.xlsx"
        path.write_bytes(b"")
        with pytest.raises(ValueError, match="empty"):
            loader.list_sheets(str(path))

    def test_html_named_xlsx_is_identified(self, tmp_path):
        """What a web app's "export to Excel" often actually produces."""
        path = tmp_path / "export.xlsx"
        path.write_bytes(b"<html><body><table><tr><td>INC0001</td></tr>"
                         b"</table></body></html>")
        with pytest.raises(ValueError, match="HTML or XML"):
            loader.list_sheets(str(path))

    def test_legacy_xls_named_xlsx_is_identified(self, tmp_path):
        path = tmp_path / "old.xlsx"
        path.write_bytes(loader._OLE2_MAGIC + b"\x00" * 4096)
        with pytest.raises(ValueError, match="older Excel format"):
            loader.list_sheets(str(path))

    def test_random_bytes_are_not_called_a_spreadsheet(self, tmp_path):
        path = tmp_path / "junk.xlsx"
        path.write_bytes(b"\x7fELF" + b"\x00" * 2048)
        with pytest.raises(ValueError, match="not a spreadsheet"):
            loader.list_sheets(str(path))

    def test_a_good_workbook_is_untouched(self, tmp_path):
        """
        The guard must not become a new way to refuse a valid file - which is
        the same rule TestExcelEngine pins for the accelerator.
        """
        path = _workbook(tmp_path / "fine.xlsx")
        assert loader.list_sheets(str(path))
        frame = loader.read_table(str(path))
        assert list(frame.columns) == ["number", "state"]
        assert len(frame) == 1

    def test_the_zip_index_probe_agrees_with_zipfile(self, tmp_path):
        """
        _has_zip_index is the whole diagnosis, so it is checked against the
        standard library rather than trusted.
        """
        import zipfile

        good = _workbook(tmp_path / "good.xlsx")
        assert loader._has_zip_index(str(good)) is True
        assert zipfile.is_zipfile(str(good)) is True

        cut = _workbook(tmp_path / "cut2.xlsx")
        _truncate(cut, cut.stat().st_size // 2)
        assert loader._has_zip_index(str(cut)) is False
        assert zipfile.is_zipfile(str(cut)) is False
