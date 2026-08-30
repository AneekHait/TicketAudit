"""
Guards the theme invariants in gui/app_pyside.py.

These are cheap text/QSS checks rather than rendering tests, because both
failure modes they cover are *silent*: Qt drops an unknown QSS property without
a warning, and styling a subcontrol removes its glyph without any error.
"""
import os
import os
import re

import pytest

# Importing the GUI module is safe headless - it only builds strings at import
# time. No QApplication is required for DARK_THEME_QSS or the token values.
from gui import app_pyside as ui


GUI_SOURCE = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "gui", "app_pyside.py",
)

HEX_RE = re.compile(r"#[0-9a-fA-F]{3,8}\b")


def _strip_qss_comments(qss: str) -> str:
    """QSS comments legitimately *mention* forbidden subcontrols to explain why."""
    return re.sub(r"/\*.*?\*/", "", qss, flags=re.S)


def _token_colours() -> set:
    """Every colour value defined in the token block, including derived shades."""
    return {
        value.lower()
        for name, value in vars(ui).items()
        if name.startswith("COLOR_") and isinstance(value, str)
    }


class TestTokenSubstitution:
    def test_no_unsubstituted_placeholders(self):
        leftover = re.findall(r"\$[A-Za-z_]+", ui.DARK_THEME_QSS)
        assert leftover == [], f"unsubstituted tokens: {leftover}"

    def test_template_and_output_differ(self):
        """Sanity check that substitution actually ran."""
        assert "$COLOR_" in ui._QSS_TEMPLATE
        assert "$COLOR_" not in ui.DARK_THEME_QSS

    def test_a_missing_token_fails_loudly(self):
        """A typo must raise at import rather than silently dropping a rule."""
        from string import Template
        with pytest.raises(KeyError):
            Template("color: $COLOR_DOES_NOT_EXIST;").substitute(
                {k: v for k, v in vars(ui).items() if k.startswith("COLOR_")}
            )


class TestEveryQssColourIsAToken:
    def test_no_hardcoded_colour_in_stylesheet(self):
        qss = _strip_qss_comments(ui.DARK_THEME_QSS)
        allowed = _token_colours()
        found = {h.lower() for h in HEX_RE.findall(qss)}
        stray = found - allowed
        assert stray == set(), (
            f"colours in the QSS that are not design tokens: {sorted(stray)}"
        )

    def test_template_itself_has_no_literal_colours(self):
        """The template must reference tokens, never spell a colour out."""
        template = _strip_qss_comments(ui._QSS_TEMPLATE)
        assert HEX_RE.findall(template) == []

    def test_semantic_roles_are_single_valued(self):
        """One colour per role - the sprawl this replaced had 4 greens and 3 reds."""
        for role in ("COLOR_SUCCESS", "COLOR_ERROR", "COLOR_WARNING", "COLOR_ACCENT"):
            assert isinstance(getattr(ui, role), str)
        # Selection must be one colour everywhere; it used to be two.
        assert ui.COLOR_PRIMARY in ui.DARK_THEME_QSS


class TestSubcontrolsAreNotStyled:
    """
    Styling a Qt subcontrol makes the stylesheet responsible for painting it,
    and QSS cannot draw a glyph without an `image:`. Doing so previously made
    the combo arrow, the checkmark and the spin box buttons vanish. Keep them
    unstyled and let the palette supply contrast.
    """

    FORBIDDEN = [
        "::drop-down",
        "::down-arrow",
        "::up-arrow",
        "::indicator",
        "::up-button",
        "::down-button",
    ]

    @pytest.mark.parametrize("subcontrol", FORBIDDEN)
    def test_not_present_in_stylesheet(self, subcontrol):
        qss = _strip_qss_comments(ui.DARK_THEME_QSS)
        assert subcontrol not in qss, (
            f"{subcontrol} is styled - this erases the widget's glyph. "
            "Use apply_dark_palette() for contrast instead."
        )

    def test_spinbox_is_not_given_a_border(self):
        """
        A border/padding on QSpinBox squeezes its native up/down buttons to an
        unusable sliver. Verified by rendering; keep spin boxes out of the
        styled-input selector.
        """
        qss = _strip_qss_comments(ui.DARK_THEME_QSS)
        for match in re.finditer(r"([^{}]*)\{([^}]*)\}", qss):
            selector, body = match.group(1), match.group(2)
            if "SpinBox" in selector and ("border" in body or "padding" in body):
                pytest.fail(f"spin box given border/padding by selector: {selector.strip()}")


class TestPaletteMatchesTokens:
    def test_palette_uses_token_values(self):
        """apply_dark_palette must not reintroduce its own colours."""
        source = open(GUI_SOURCE, encoding="utf-8").read()
        start = source.index("def apply_dark_palette")
        end = source.index("class PandasModel")
        body = source[start:end]
        assert HEX_RE.findall(body) == [], (
            "apply_dark_palette contains hardcoded colours; use the tokens"
        )


class TestNothingBypassesTheTokens:
    """
    The token block's comment has always claimed that no hex literal appears
    outside it. That was not true: the old checks only read DARK_THEME_QSS, so
    two literals in widget code and twenty-eight in the Help/About markup sat
    there for months - a whole second theme that a token change could not move.
    """

    def _source(self):
        path = os.path.join(os.path.dirname(os.path.dirname(
            os.path.abspath(__file__))), "gui", "app_pyside.py")
        with open(path, encoding="utf-8") as handle:
            return handle.read()

    def test_no_hex_literal_outside_the_token_block(self):
        source = self._source()
        # The block ends where the stylesheet template begins; everything after
        # it - QSS, widget code, help markup - must name a token instead.
        _tokens, rest = source.split("_QSS_TEMPLATE = ", 1)
        rest = _strip_qss_comments(rest)
        found = {}
        for match in HEX_RE.finditer(rest):
            line = rest[:match.start()].count(chr(10)) + 1
            found.setdefault(match.group(0), []).append(line)
        allowed = _token_colours()
        offenders = {colour: lines for colour, lines in found.items()
                     if colour.lower() not in allowed}
        assert not offenders, (
            "raw colours outside the token block: " + repr(offenders))

    def test_the_help_markup_uses_placeholders(self):
        """
        Those dialogs are rich text, so QSS cannot reach them. They go through
        html(), which is the only other place a colour may be named.
        """
        source = self._source()
        assert "def html(markup: str)" in source
        for dialog in ("about_text", "guide_text", "shortcuts_text"):
            assert "html(" + dialog + ")" in source, dialog

    def test_no_emoji_in_the_interface(self):
        """
        Emoji render differently on every machine, and every one of these sat
        next to a colour or a word already carrying the meaning. One had even
        been made load-bearing once, by code that parsed the prefix back out.
        """
        source = self._source()
        ranges = ((0x1F300, 0x1FAFF), (0x2600, 0x27BF), (0x2190, 0x21FF),
                  (0xFE0F, 0xFE0F))
        found = sorted({char for char in source
                        if any(low <= ord(char) <= high for low, high in ranges)})
        # The arrow is content, not decoration: it separates a date range and
        # points at a menu path in the guide.
        found = [char for char in found if char != chr(0x2192)]
        assert not found, "emoji in the UI source: " + repr(found)


class TestTheThemeIsAppliedOnce:
    def test_one_stylesheet_call(self):
        """
        Applied to the main window, in one place. A second setStyleSheet on the
        QApplication re-polishes every live widget, which is slow enough on a few
        hundred widgets to look like a hang.
        """
        source = TestNothingBypassesTheTokens()._source()
        assert source.count("setStyleSheet(DARK_THEME_QSS)") == 1

    def test_the_proxy_style_is_built_from_a_name(self):
        """
        QProxyStyle takes ownership of the style it is given, and setStyle
        deletes the outgoing one - so handing it app.style() leaves it holding a
        freed pointer, which presented as the test suite hanging.
        """
        source = TestNothingBypassesTheTokens()._source()
        assert "TicketAuditStyle(current.objectName()" in source
        assert "TicketAuditStyle(app.style())" not in source
