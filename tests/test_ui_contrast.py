"""
Contrast and colour-separation guards for the theme.

The theme is a hybrid: dark chrome (sidebar, top bar, menu bar) around a light
workspace. The two families never mix, and that is the thing most easily broken
here - workspace ink on a chrome surface measures about 1.1:1, which is exactly
the invisibility this palette was rebuilt to remove. So each family is measured
against its own text, and TestChromeIsItsOwnFamily pins that they are separate.

These exist because "it looks fine" is not checkable and the failures were
invisible in review. Every number here was a real defect before the overhaul:

* An unchecked checkbox was drawn at **1.10:1** against a card - Fusion derives
  the frame from QPalette::Window darkened, and 153 of the indicator's 196
  pixels were simply the card showing through. Nobody could see it.
* A card sat at **1.10:1** against the page behind it, so panels had no edge.
* COLOR_PRIMARY measured **1.64:1** as text, which is why a second blue existed
  only for headings.

Two of these are measured from *rendered pixels* rather than from the token
values, because the tokens were never the problem - what Qt actually painted
was. Colour comes out right under the offscreen plugin; only text does not, and
nothing here measures text. For screenshots you actually look at, see
scripts/ui_shots.py, which needs the native plugin for fonts.
"""
import colorsys
import os
import sys

import pytest

# Offscreen, like the other GUI tests. These measure painted *colour*, which
# the offscreen plugin renders correctly - it only lacks fonts, and no assertion
# here depends on a glyph. Popping this variable instead (to get fonts) would
# change the platform for the whole pytest process, and every window.show() in
# test_gui_smoke would then open a real window on the developer's desktop.
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

pytest.importorskip("PySide6")

from PySide6.QtCore import Qt                                       # noqa: E402
from PySide6.QtWidgets import (                                     # noqa: E402
    QApplication, QCheckBox, QLineEdit, QStyle, QStyleOptionButton, QVBoxLayout,
    QWidget,
)

import gui.app_pyside as ui                                         # noqa: E402

# WCAG 2.1: 4.5 for body text, 3.0 for large text and for the boundary of a
# control you have to be able to find.
AA_TEXT = 4.5
AA_UI = 3.0
# Two hues encoding different things need to be separable at a glance. 30 deg is
# the floor below which two status colours start reading as one.
MIN_HUE_SEPARATION = 30.0


def _channels(colour: str):
    colour = colour.lstrip("#")
    return tuple(int(colour[i:i + 2], 16) / 255 for i in (0, 2, 4))


def luminance(colour: str) -> float:
    def linear(channel: float) -> float:
        return (channel / 12.92 if channel <= 0.03928
                else ((channel + 0.055) / 1.055) ** 2.4)

    red, green, blue = (linear(c) for c in _channels(colour))
    return 0.2126 * red + 0.7152 * green + 0.0722 * blue


def contrast(one: str, two: str) -> float:
    first, second = luminance(one), luminance(two)
    high, low = max(first, second), min(first, second)
    return (high + 0.05) / (low + 0.05)


def hue_degrees(colour: str) -> float:
    return colorsys.rgb_to_hls(*_channels(colour))[0] * 360


def hue_separation(one: str, two: str) -> float:
    gap = abs(hue_degrees(one) - hue_degrees(two))
    return min(gap, 360 - gap)


@pytest.fixture(scope="module")
def qapp():
    """
    The app, with the theme installed but *no application-level stylesheet*.

    Calling setStyleSheet on the QApplication re-polishes every live widget in
    the process. Harmless in isolation, but this module runs last, after
    test_gui_smoke has built a few hundred windows, and the repolish storm made
    the suite appear to hang. The stylesheet goes on the host widget instead,
    where it still cascades to the widget under test.
    """
    app = QApplication.instance() or QApplication(sys.argv[:1])
    if not isinstance(app.style(), ui.TicketAuditStyle):
        app.setStyle("Fusion")
    ui.apply_dark_palette(app)
    return app


def _on_card(qapp, widget):
    """Render `widget` on a card surface and hand back the image."""
    host = QWidget()
    host.setObjectName("CardHost")
    host.setStyleSheet(
        ui.DARK_THEME_QSS
        + "\nQWidget#CardHost { background: %s; }" % ui.COLOR_SURFACE_RAISED)
    layout = QVBoxLayout(host)
    layout.setContentsMargins(8, 8, 8, 8)
    layout.addWidget(widget)
    host.resize(320, 48)
    host.grab()                      # force a layout and paint pass
    return host, host.grab().toImage()


def _strongest_contrast(image, rect_origin, size, against):
    """The most visible pixel in a region, as a contrast ratio."""
    best = 1.0
    for x in range(rect_origin[0], rect_origin[0] + size[0]):
        for y in range(rect_origin[1], rect_origin[1] + size[1]):
            best = max(best, contrast(image.pixelColor(x, y).name(), against))
    return best


class TestTextOnEverySurface:
    """Body and secondary text has to be readable on every surface it lands on."""

    SURFACES = ("COLOR_SURFACE_SUNKEN", "COLOR_SURFACE_DEEP",
                "COLOR_SURFACE_BASE", "COLOR_SURFACE_RAISED",
                "COLOR_SURFACE_OVERLAY")

    @pytest.mark.parametrize("surface", SURFACES)
    def test_body_text(self, surface):
        assert contrast(ui.COLOR_TEXT, getattr(ui, surface)) >= AA_TEXT

    @pytest.mark.parametrize("surface", SURFACES)
    def test_muted_text(self, surface):
        """Captions and help text are still text, so they get the text bar."""
        assert contrast(ui.COLOR_TEXT_MUTED, getattr(ui, surface)) >= AA_TEXT

    @pytest.mark.parametrize("surface", SURFACES)
    def test_disabled_text_is_dim_but_findable(self, surface):
        ratio = contrast(ui.COLOR_TEXT_DISABLED, getattr(ui, surface))
        assert ratio >= AA_UI, "disabled is not the same as illegible"
        assert ratio < contrast(ui.COLOR_TEXT, getattr(ui, surface)), \
            "disabled text must read as quieter than body text"

    def test_accent_works_as_text(self):
        """
        The old COLOR_PRIMARY was 1.64:1 as text, which is why a second blue
        existed purely for headings. One accent has to do both jobs.
        """
        assert contrast(ui.COLOR_ACCENT, ui.COLOR_SURFACE_RAISED) >= AA_TEXT
        assert contrast(ui.COLOR_ACCENT, ui.COLOR_SURFACE_BASE) >= AA_TEXT

    def test_text_on_selection_fill(self):
        assert contrast(ui.COLOR_TEXT_STRONG, ui.COLOR_PRIMARY) >= AA_TEXT


class TestChromeIsItsOwnFamily:
    """
    The dark shell has its own text and its own surfaces, and they have to clear
    the same bars as the light workspace does.

    The failure this guards is not subtle once seen and completely invisible
    until then: a chrome widget that forgets to name a COLOR_CHROME_* colour
    inherits COLOR_TEXT from the base QWidget rule, which is near-black ink on a
    near-black panel. The assertion at the bottom is what makes that a *test*
    rather than a convention - if the two families ever converge, the mistake
    stops being catchable.
    """

    SURFACES = ("COLOR_CHROME", "COLOR_CHROME_RAISED", "COLOR_CHROME_SUNKEN",
                "COLOR_CHROME_HOVER", "COLOR_CHROME_ACTIVE")

    @pytest.mark.parametrize("surface", SURFACES)
    def test_chrome_body_text(self, surface):
        assert contrast(ui.COLOR_CHROME_TEXT, getattr(ui, surface)) >= AA_TEXT

    @pytest.mark.parametrize("surface", SURFACES)
    def test_chrome_muted_text(self, surface):
        assert contrast(ui.COLOR_CHROME_MUTED, getattr(ui, surface)) >= AA_TEXT

    @pytest.mark.parametrize("surface", SURFACES)
    def test_chrome_faint_text_is_findable(self, surface):
        """Group captions and row numbers. Quiet, not illegible."""
        ratio = contrast(ui.COLOR_CHROME_FAINT, getattr(ui, surface))
        assert ratio >= AA_UI
        assert ratio < contrast(ui.COLOR_CHROME_TEXT, getattr(ui, surface))

    def test_the_accent_has_a_weight_that_works_on_chrome(self):
        """
        COLOR_ACCENT is tuned for 12px type on white, so on the chrome it is
        nearly black. ACCENT_BRIGHT is the same hue lifted for that surface, and
        it is why two teals exist.
        """
        assert contrast(ui.COLOR_ACCENT_BRIGHT, ui.COLOR_CHROME) >= AA_TEXT
        assert hue_separation(ui.COLOR_ACCENT_BRIGHT, ui.COLOR_ACCENT) < 15.0

    @pytest.mark.parametrize("role", ("COLOR_ERROR_BRIGHT",
                                      "COLOR_WARNING_BRIGHT"))
    def test_a_badge_reads_on_the_chrome(self, role):
        """A navigation badge is dark text on a light pill, so measure both."""
        assert contrast(getattr(ui, role), ui.COLOR_CHROME) >= AA_TEXT
        assert contrast(ui.COLOR_CHROME, getattr(ui, role)) >= AA_UI

    def test_chrome_edges_are_visible_against_the_workspace(self):
        """The one place the families meet: the sidebar's right-hand edge."""
        assert contrast(ui.COLOR_CHROME, ui.COLOR_SURFACE_BASE) >= AA_UI

    def test_the_two_families_do_not_overlap(self):
        assert contrast(ui.COLOR_TEXT, ui.COLOR_CHROME) < AA_UI, \
            "workspace ink is supposed to be unusable on chrome - if it stops " \
            "being so, the family split stops being enforceable"
        assert contrast(ui.COLOR_CHROME_TEXT, ui.COLOR_SURFACE_RAISED) < AA_UI


class TestStatusColours:
    ROLES = ("COLOR_SUCCESS", "COLOR_WARNING", "COLOR_ERROR", "COLOR_ACCENT")

    @pytest.mark.parametrize("role", ROLES)
    def test_readable_on_a_card(self, role):
        assert contrast(getattr(ui, role), ui.COLOR_SURFACE_RAISED) >= AA_UI

    @pytest.mark.parametrize("role,fill", [
        ("COLOR_SUCCESS", "COLOR_SUCCESS_FILL"),
        ("COLOR_WARNING", "COLOR_WARNING_FILL"),
        ("COLOR_ERROR", "COLOR_ERROR_FILL"),
        ("COLOR_ACCENT", "COLOR_ACCENT_FILL"),
    ])
    def test_readable_on_its_own_severity_fill(self, role, fill):
        """A severity card tints its background; the text sits on that tint."""
        assert contrast(getattr(ui, role), getattr(ui, fill)) >= AA_TEXT

    @pytest.mark.parametrize("one,two", [
        ("COLOR_ACCENT", "COLOR_WARNING"),
        ("COLOR_ACCENT", "COLOR_ERROR"),
        ("COLOR_SUCCESS", "COLOR_WARNING"),
        ("COLOR_SUCCESS", "COLOR_ERROR"),
    ])
    def test_hues_are_separable(self, one, two):
        """
        Meanings that co-occur have to stay separable at a glance - a warning
        beside an error, the accent beside either.

        The accent-vs-"pass" pair is deliberately absent: TicketAudit's brand is
        forest green, so COLOR_ACCENT and COLOR_SUCCESS share a hue on purpose
        (a selected row and an OK badge are meant to read as the same family
        here). Every other pair still has to clear MIN_HUE_SEPARATION.
        """
        assert hue_separation(getattr(ui, one), getattr(ui, two)) >= \
            MIN_HUE_SEPARATION


class TestBorderWeights:
    """
    Three weights, three jobs. Only the divider may be invisible-ish; a control
    boundary that cannot be found is a usability bug, not a subtle style.
    """

    def test_control_border_is_perceivable(self):
        assert contrast(ui.COLOR_BORDER, ui.COLOR_SURFACE_RAISED) >= AA_UI
        assert contrast(ui.COLOR_BORDER, ui.COLOR_SURFACE_SUNKEN) >= AA_UI

    def test_the_weights_are_ordered(self):
        card = ui.COLOR_SURFACE_RAISED
        divider = contrast(ui.COLOR_BORDER_DIVIDER, card)
        outline = contrast(ui.COLOR_BORDER_SUBTLE, card)
        control = contrast(ui.COLOR_BORDER, card)
        assert divider < outline < control

    def test_a_card_has_an_edge_because_tone_cannot_do_it(self):
        """
        Documenting the reason in a test: the tone step is imperceptible, so if
        the outline ever leaves the stylesheet, cards disappear again.
        """
        assert contrast(ui.COLOR_SURFACE_RAISED, ui.COLOR_SURFACE_BASE) < 1.3
        assert "QFrame#Card" in ui.DARK_THEME_QSS
        card_rule = ui.DARK_THEME_QSS.split("QFrame#Card {")[1].split("}")[0]
        assert "border" in card_rule, \
            "a card is 1.1:1 against the page - the outline is what makes it read"


class TestRenderedIndicators:
    """
    Measured from pixels, not tokens. The tokens were never wrong here - Fusion
    was painting a frame derived from QPalette::Window, which no token controls.
    """

    def _indicator(self, qapp, checked):
        box = QCheckBox("Exclude cancelled tickets")
        box.setChecked(checked)
        host, image = _on_card(qapp, box)
        option = QStyleOptionButton()
        box.initStyleOption(option)
        rect = box.style().subElementRect(
            QStyle.SubElement.SE_CheckBoxIndicator, option, box)
        origin = box.mapTo(host, rect.topLeft())
        return _strongest_contrast(image, (origin.x(), origin.y()),
                                   (rect.width(), rect.height()),
                                   ui.COLOR_SURFACE_RAISED)

    def test_an_unchecked_box_can_be_seen(self, qapp):
        """The bug that started the overhaul: this measured 1.10:1."""
        assert self._indicator(qapp, False) >= AA_UI

    def test_a_checked_box_can_be_seen(self, qapp):
        assert self._indicator(qapp, True) >= AA_UI

    def test_the_two_states_are_distinguishable(self, qapp):
        assert self._indicator(qapp, True) > self._indicator(qapp, False)

    def test_an_input_has_a_findable_edge(self, qapp):
        """Scan the whole field: the border is somewhere on its perimeter."""
        field = QLineEdit()
        host, image = _on_card(qapp, field)
        origin = field.mapTo(host, field.rect().topLeft())
        best = _strongest_contrast(image, (origin.x(), origin.y()),
                                   (field.width(), field.height()),
                                   ui.COLOR_SURFACE_RAISED)
        assert best >= AA_UI, "an input must not blend into the card"


class TestThemeIsAppliedThroughOneFunnel:
    def test_apply_dark_palette_installs_the_style(self, qapp):
        """
        Every entry point - main.py, the __main__ block, the test fixtures -
        already calls apply_dark_palette right after setStyle. Installing the
        style there is what keeps the indicator fix from depending on which one
        started the app.
        """
        ui.apply_dark_palette(qapp)
        assert isinstance(qapp.style(), ui.TicketAuditStyle)

    def test_installing_twice_does_not_nest_proxies(self, qapp):
        """
        A long-lived process calls this more than once; each call used to add
        another layer to the style chain.
        """
        ui.apply_dark_palette(qapp)
        first = qapp.style()
        ui.apply_dark_palette(qapp)
        assert qapp.style() is first

    def test_the_style_does_not_touch_anything_else(self, qapp):
        """
        Only the check indicators are ours; everything else must fall through to
        Fusion, or we own the whole widget set by accident.
        """
        import inspect
        source = inspect.getsource(ui.TicketAuditStyle)
        assert source.count("PE_Indicator") == 3
        assert "super().drawPrimitive" in source
