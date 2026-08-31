"""
Modern PySide6 GUI for TicketAudit - ITSM ticket data QA.
Features:
- Native look and feel with dark theme
- Model/View architecture for tables
- Asynchronous operations with QThread
- Interactive charts with QChart
"""
import sys
import os
import logging
import math
import threading
import datetime
from string import Template
from typing import Optional, Dict, Any, List

import pandas as pd
import numpy as np

# Nothing configures logging, so logging's "last resort" handler prints
# WARNING and above to stderr while debug tracing stays silent. That gives a
# console record of swallowed exceptions without narrating every file open,
# which is what the old print("DEBUG: ...") calls did on every load.
log = logging.getLogger("ticketaudit")

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QTabWidget, QLabel, QPushButton, QComboBox, QFileDialog, QProxyStyle,
    QMessageBox, QFrame, QScrollArea, QSplitter, QTableView,
    QHeaderView, QAbstractItemView, QLineEdit, QListWidget, 
    QListWidgetItem, QProgressBar, QTextEdit, QPlainTextEdit, QSizePolicy,
    QStatusBar, QCheckBox, QSlider, QDockWidget, QStyle, QDialog,
    QSpinBox, QStackedWidget, QDoubleSpinBox, QMenu,
    QStyledItemDelegate, QStyleOptionViewItem, QButtonGroup
)
from PySide6.QtCore import (
    Qt, QThread, Signal, Slot, QSize, QAbstractTableModel,
    QModelIndex, QSortFilterProxyModel, QRunnable, QThreadPool,
    QObject, QRect, QUrl, QRectF, QPointF, QMargins
)
from PySide6.QtGui import (
    QAction, QIcon, QFont, QStandardItemModel, QStandardItem,
    QColor, QPalette, QBrush, QTextDocument, QTextCursor, QPainter, QPixmap,
    QKeySequence, QDesktopServices, QPen, QPainterPath
)

from PySide6.QtCharts import (
    QChart, QChartView, QBarSeries, QBarSet, QBarCategoryAxis, 
    QValueAxis, QAbstractBarSeries
)

# Core imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import ConfigManager, COLUMN_KEYWORDS, ID_KEYWORDS
from core import loader
from core.analyzer import COLUMN_ABSENT, SanityAnalyzer
from core.language import LanguageChecker

from core.excel_report import DETAIL_MAX_DATA_ROWS
from core import extract as extract_core
from core.reporter import ReportGenerator

# ---------------------------------------------------------------------------
# Design tokens
#
# Every colour, radius, spacing and font size in the UI is defined here and
# nowhere else. The stylesheet below is a string.Template that references these
# by name, so a value exists in exactly one place. tests/test_design_tokens.py
# enforces that no hex literal appears outside this block.
#
# Template (not f-strings/.format) because the QSS is full of { } braces that
# would all need doubling; the QSS contains no '$', so $NAME needs no escaping.
# .substitute() raises KeyError on a typo - a loud failure, not a dropped rule.
# ---------------------------------------------------------------------------

# Two families of surface, because this is a hybrid shell: the navigation and
# the top bar are dark chrome, and the workspace they frame is light. That split
# is deliberate. The people who use this tool live in Excel and ServiceNow all
# day, and a large dark panel holding a sparse result reads as unfinished, where
# the same panel on paper reads as a document. The chrome stays dark so "where
# am I" never competes with "what am I looking at" for the same tone.
#
# The families never mix. A widget is either chrome or workspace, and its text
# colour follows from that: COLOR_CHROME_TEXT on the former, COLOR_TEXT on the
# latter. Putting workspace ink on a chrome surface measures 1.1:1 - which is
# the same invisibility this palette was rebuilt to remove, so
# tests/test_ui_contrast.py measures each family against its own text.

# --- Chrome: the dark shell ---
COLOR_CHROME = "#0f1518"            # sidebar, top bar, menu bar
COLOR_CHROME_RAISED = "#161e22"     # a panel sitting on the chrome
COLOR_CHROME_SUNKEN = "#0a1012"     # an input sitting on the chrome
COLOR_CHROME_HOVER = "#1c262b"      # hover on a chrome row or button
COLOR_CHROME_ACTIVE = "#123430"     # the selected navigation row
COLOR_CHROME_LINE = "#29353a"       # chrome edges and hairlines
COLOR_CHROME_TEXT = "#eff5f3"       # body text on chrome          16.7:1
COLOR_CHROME_MUTED = "#91a19e"      # captions on chrome            6.8:1
COLOR_CHROME_FAINT = "#6d7d7a"      # group labels, row numbers     4.3:1

# --- Workspace: the light page the chrome frames ---
#
# Elevation still cannot be carried by tone: a card is 1.14:1 against the page,
# imperceptible at any hue. A card reads as a card because of its 1px
# COLOR_BORDER_SUBTLE outline - if that leaves the QSS, cards disappear again.
# SUNKEN and RAISED are the same white on purpose: an input and the card under
# it are separated by a border, not a fill, and inventing a fourth off-white
# only makes the page look grubby.
COLOR_SURFACE_DEEP = "#e8edeb"      # scrollbar troughs, status bar
COLOR_SURFACE_BASE = "#edf1ef"      # the page behind the cards
COLOR_SURFACE_OVERLAY = "#f1f4f3"   # buttons, hover, table headers, alt rows
COLOR_SURFACE_RAISED = "#ffffff"    # cards and panes
COLOR_SURFACE_SUNKEN = "#ffffff"    # inputs and table bodies

# Three border weights, because they do three different jobs and only one of
# them may be subtle:
#   DIVIDER  1.2:1  hairlines between rows and sections - decoration
#   SUBTLE   1.5:1  card and panel edges - reads as an edge, not a line
#   BORDER   3.5:1  checkbox, radio and input frames - a *control* boundary
# BORDER is markedly darker than the mockup's line colour, and that is not a
# transcription error: a control boundary has to clear 3:1, and the mockup's
# 1.5:1 hairline is the bug that started the whole overhaul.
COLOR_BORDER_DIVIDER = "#e3e9e7"
COLOR_BORDER_SUBTLE = "#ccd6d3"
COLOR_BORDER = "#7f8c89"

# Text, most to least prominent. Ratios are against COLOR_SURFACE_RAISED.
COLOR_TEXT_STRONG = "#ffffff"       # on accent/selection fills only
COLOR_TEXT = "#17211f"              # body                     16.5:1
COLOR_TEXT_MUTED = "#5c6a67"        # captions, secondary info  5.7:1
COLOR_TEXT_DISABLED = "#7b8885"     #                           3.7:1

# Semantic roles - exactly one value each for the workspace, plus a lifted
# variant for the few places the same meaning has to land on dark chrome (a
# navigation badge). WARNING means "a finding that is not fatal"; transient
# states such as "Analyzing..." are NOT warnings and use COLOR_TEXT_MUTED.
#
# These are darker than the mockup's, and each one is set by the bar it has to
# clear: a status colour is read as *text on its own tint*, so SUCCESS must make
# 4.5:1 against SUCCESS_FILL, not merely against white. The mockup's coral and
# amber measured 3.8 and 4.0 on their own fills.
COLOR_SUCCESS = "#1d7a3c"           # 5.4:1 on white, 4.8:1 on its fill
COLOR_SUCCESS_FILL = "#e6f5ec"
COLOR_ERROR = "#b4322d"             # 6.1:1 on white, 5.2:1 on its fill
COLOR_ERROR_FILL = "#fbe9e7"
COLOR_WARNING = "#8f6100"           # 5.4:1 on white, 4.9:1 on its fill
COLOR_WARNING_FILL = "#fff2da"
COLOR_ERROR_BRIGHT = "#f2938e"      # the same meaning, on chrome
COLOR_WARNING_BRIGHT = "#efae47"

# One accent, in four weights. ACCENT is the text/icon weight and has to work
# as 12px type on white, so it is a dark forest green; ACCENT_BRIGHT is the same
# hue lifted for chrome, where a 6:1-on-white green would be nearly black.
# PRIMARY is the fill under white text, and FILL is the tint behind ACCENT text.
#
# TicketAudit's brand is forest green, so the accent shares a hue with the
# "pass" colour by design - the selection fill and a success badge are meant to
# read as the same family here. There is deliberately no separate info hue -
# severity="info" uses the accent itself.
COLOR_ACCENT = "#15733f"            # text, icons, chips           5.9:1
COLOR_ACCENT_BRIGHT = "#3fc16e"     # icons and marks on chrome    7.9:1
COLOR_ACCENT_FILL = "#e7f5ec"       # selected row, info card fill
COLOR_PRIMARY = "#177a44"           # primary buttons and selection fills
                                    # white on it: 5.4:1

# Metrics. Spacing is a 4-based scale so every gap in the UI is a multiple of
# the same unit; the old 5/10/15/20 could not subdivide evenly.
RADIUS_SM = 4
RADIUS_MD = 8
RADIUS_LG = 10
SPACE_XS = 4
SPACE_SM = 8
SPACE_MD = 12
SPACE_LG = 16
SPACE_XL = 24
SPACE_XXL = 32
# The scale is deliberately *wide*: small dense body text under a large title,
# which is what makes a sparse result read as a document rather than a poster.
# The mockup runs 7px micro-labels against a 24px page title - a 3.4x spread -
# where this was 11 against 20, a 1.8x spread, and the flatness is why every
# view read as equally loud. These are the mockup's proportions at sizes that
# stay legible on a 14-inch 1080p panel at 150% scaling, which is the machine
# this tool is actually used on.
FONT_XS = 10                        # micro: table headers, group captions, badge
FONT_SM = 11                        # captions, secondary, a nav row
FONT_MD = 12                        # body
FONT_LG = 14                        # card and section titles
FONT_XL = 16                        # an inspector heading
FONT_DISPLAY = 22                   # the one view title in the chrome
FONT_KPI = 24                       # the number on a summary tile
# The About dialog's natural width. Wide enough that its two-column feature
# table stays two columns; the dialog is still capped to the screen on top of
# this, because a fixed width that exceeds the display is the bug this replaced.
ABOUT_WIDTH = 560
# The height to *ask* for, not a fixed one: it is capped to the screen below.
# Asking for the content's sizeHint instead does not work once the content is in
# a scroll area - the hint then describes the scroll area's own minimum, which
# came out at 463px and made the dialog shorter than the display it had room in,
# scrolling content that would have fitted.
ABOUT_HEIGHT = 640
# Kept clear of the screen edges so the dialog never opens flush against them.
ABOUT_SCREEN_MARGIN = 80

# The brand mark in the sidebar. 26 matches the wordmark's cap height beside it;
# the .ico carries a 32px face, so QIcon scales from that rather than from 256.
BRAND_MARK_SIZE = 26
SIDEBAR_WIDTH = 244
SIDEBAR_MAX_WIDTH = 420
SIDEBAR_ROW_HEIGHT = 25
# A group caption is one small word, so it is sized for the word and the gap
# above it rather than given a full row. These were 34 and 22, which spent 192px
# of a 584px column on six captions and pushed three views off a 14-inch screen.
# The first caption needs no gap above it - the dataset block already separates
# it from the brand.
NAV_GROUP_HEIGHT = 22
NAV_GROUP_FIRST_HEIGHT = 14
TOPBAR_HEIGHT = 46
# The evidence panel. 360 fits an ID, a state and a timestamp on one row of the
# evidence sample without eliding; below INSPECTOR_MIN_WIDTH it is hidden
# instead of squeezed, because a truncated evidence row is worse than none.
INSPECTOR_WIDTH = 360
INSPECTOR_MIN_WIDTH = 300
SCROLLBAR_SIZE = 10
PROGRESS_HEIGHT = 4
# A row in a combo's drop-down. Stated rather than left to the font, so the
# popup's height is predictable - see the QComboBox QAbstractItemView rule.
COMBO_ITEM_HEIGHT = 24
# Tracks FONT_MD: at 12px body, 26 keeps the same ratio 30 had at 14px. The
# rule it encodes is that a row must not be tighter than the text it holds -
# below that the table reads as a spreadsheet dump.
TABLE_ROW_HEIGHT = 26

# Rows shown in the Extract preview. Enough to see whether the right columns
# and the right months came out, few enough to build in about a millisecond.
EXTRACT_PREVIEW_ROWS = 15
# ...and how many of them are visible without scrolling the preview itself.
# Pinning all fifteen made the preview 480px - taller than the entire content
# viewport on a 14-inch laptop, so it pushed the pickers above it off screen.
# The table still holds EXTRACT_PREVIEW_ROWS and scrolls to them.
EXTRACT_PREVIEW_VISIBLE = 6

# One line per view, shown under the title in the content header. They live
# here rather than inside each builder because three views used to print their
# own heading as well, in a different style, and the page then said the same
# thing twice. A view with no entry simply gets no subtitle.
VIEW_SUBTITLES = {
    "Overview": "Every check this file has been through, worst first, with what "
                "was never looked at.",
    "Column Check": "Which required fields this export has, and what each one "
                    "was matched to.",
    "Null Analysis": "How much of each column is empty, against the threshold "
                     "in Settings.",
    "Logic Checks": "Contradictions between fields - impossible dates, states "
                    "that disagree with their resolution date.",
    "Language Check": "Non-English text in a chosen column, and which detection "
                      "tiers actually ran.",
    "Description Quality": "Descriptions too short to act on, too long to read, "
                           "or repeated across tickets.",
    "Monthly Inflow": "Ticket volume by month, from the created date.",
    "Pivot Tables": "How the values of one column are distributed, largest "
                    "first.",
    "Data Profile": "Type, fill and cardinality for every column.",
    "Duplicate Check": "Records sharing an ID, and the rows involved.",
    "Raw Data": "The sheet as loaded, for looking things up.",
    "Extract": "Write a smaller file: a month range and the columns you need.",
    "Report": "The whole review as a text document, to file or to paste into a "
              "ticket.",

}

# The guided review walks the views in the order a reviewer actually needs
# them, grouped into four stages. It is a *mode*, not the default: outside it
# the ribbon is hidden and nothing on the page mentions a step number, because
# a progress bar you did not ask for is a claim that you are behind.
#
# The names on the left are the stages; the lists are view names from tab_defs,
# so a view that is renamed or removed drops out of the walk instead of leaving
# a dead step. REVIEW_STAGES is the authority for both the ribbon and the
# "Step n of m" eyebrow.
REVIEW_STAGES = [
    ("Map", ["Column Check"]),
    ("Validate", ["Null Analysis", "Logic Checks"]),
    ("Investigate", ["Language Check", "Description Quality", "Duplicate Check"]),
    ("Deliver", ["Overview", "Report"]),
]

# The two ways to choose months in the Extract view. Stored as item data on the
# mode combo, never read from its display text.
EXTRACT_MONTHS_RANGE = "range"
EXTRACT_MONTHS_PICK = "pick"

# How many months the Extract range combos show before scrolling. Qt's default
# is 10, which hides most of a two-year export behind a scrollbar.
EXTRACT_MONTHS_VISIBLE = 18
TABLE_MAX_COL_WIDTH = 400

# Three faces, as the mockup has. Body stays Segoe UI: it is the proven
# Windows UI face, and the mockup's first choice (Aptos) is not installed here -
# what was actually missing is the other two, which is where its character sits.
#
# DISPLAY is for the wordmark, the page title and any big number. Bahnschrift is
# a condensed DIN, so it sets a large title in less width than Segoe UI and
# reads as drawn rather than typed. MONO is for identifiers only - the nav row
# numbers, the brand mark, a ticket ID - because a column of proportional digits
# does not line up and a ticket number is a code, not a word.
#
# Every family in each stack was checked as installed on the target platform;
# the last entry of each is a generic so a missing face degrades rather than
# falling back to whatever Qt picks first.
FONT_FAMILY = "'Segoe UI', Arial, sans-serif"
FONT_FAMILY_DISPLAY = ("'Bahnschrift', 'Segoe UI Variable Display', "
                       "'Segoe UI Semibold', 'Segoe UI', sans-serif")
FONT_FAMILY_MONO = "'Cascadia Mono', Consolas, 'Courier New', monospace"
# The same stack QFont wants: a list of names, no quotes, no generic. Derived
# from the QSS string rather than retyped, so the two cannot drift.
NAV_INDEX_FAMILIES = [name.strip().strip("'")
                      for name in FONT_FAMILY_MONO.split(",")
                      if "monospace" not in name]


def _lighter(hex_color: str, pct: int = 115) -> str:
    """Derive a hover shade so state variants add no new colour literals."""
    return QColor(hex_color).lighter(pct).name()


def _darker(hex_color: str, pct: int = 130) -> str:
    """Derive a deeper shade, for large fills where the base hue is too vivid."""
    return QColor(hex_color).darker(pct).name()


# On a light workspace a hover is *darker*, which is the opposite of the dark
# theme this replaced - hence _darker here where it used to be _lighter.
COLOR_PRIMARY_HOVER = _darker(COLOR_PRIMARY, 118)
COLOR_OVERLAY_HOVER = _darker(COLOR_SURFACE_OVERLAY, 106)
# ...except on the chrome, where the surrounding surface is dark and the same
# button has to lift to read as hovered.
COLOR_PRIMARY_LIFT = _lighter(COLOR_PRIMARY, 122)

# --- Stylesheet ---
_QSS_TEMPLATE = """
/* =========================================================================
   The workspace: everything is light unless it is named as chrome below.
   ========================================================================= */
QMainWindow, QWidget {
    background-color: $COLOR_SURFACE_BASE;
    color: $COLOR_TEXT;
    font-family: $FONT_FAMILY;
    font-size: $FONT_MD;
}
/* The rule above matches subclasses, so labels would otherwise paint an opaque
   base-coloured patch on top of a card. Keep them transparent. */
QLabel, QCheckBox, QRadioButton {
    background: transparent;
}

/* =========================================================================
   The chrome: the dark sidebar, top bar and menu bar.

   Every chrome widget is named. The QWidget rule above matches subclasses, so
   an unnamed container inside the sidebar paints a light patch onto it - and a
   descendant selector is not the fix: `QFrame#Sidebar QWidget` outranks
   `QFrame#DatasetBlock` on Qt's specificity and would flatten the panel it is
   meant to leave alone.
   ========================================================================= */
QFrame#Sidebar, QFrame#TopBar, QWidget#SidebarBrand, QWidget#SidebarFooter,
QWidget#TopBarLeft, QWidget#TopBarRight {
    background-color: $COLOR_CHROME;
    color: $COLOR_CHROME_TEXT;
}
QFrame#Sidebar { border-right: 1px solid $COLOR_CHROME_LINE; }
QFrame#TopBar  { border-bottom: 1px solid $COLOR_CHROME_LINE; }

/* The brand mark carries its own field and its own rounded edge, so it takes no
   background here - a plate behind it would show through the transparent
   corners of the artwork as four teal notches. The colour and font are the
   fallback's: they only apply when the asset is missing and the label falls
   back to the letters "TA". */
QLabel#BrandMark {
    color: $COLOR_ACCENT_BRIGHT;
    font-family: $FONT_FAMILY_MONO;
    font-size: $FONT_SM;
    font-weight: bold;
}
QLabel#Wordmark {
    font-family: $FONT_FAMILY_DISPLAY;
    font-size: $FONT_XL;
    font-weight: 600;
    color: $COLOR_CHROME_TEXT;
}
QFrame#DatasetBlock {
    background-color: $COLOR_CHROME_RAISED;
    border: 1px solid $COLOR_CHROME_LINE;
    border-radius: $RADIUS_MD;
}
QLabel#DatasetKicker {
    font-size: $FONT_XS;
    font-weight: bold;
    color: $COLOR_CHROME_FAINT;
}
QLabel#DatasetName {
    font-size: $FONT_SM;
    font-weight: bold;
    color: $COLOR_CHROME_TEXT;
}
QLabel#DatasetMeta {
    font-size: $FONT_XS;
    color: $COLOR_CHROME_MUTED;
}
QLabel#Breadcrumb {
    font-size: $FONT_SM;
    color: $COLOR_CHROME_MUTED;
}

/* Chrome buttons. Two weights: the export action, and the quiet ones beside it. */
QPushButton#ChromePrimary {
    background-color: $COLOR_PRIMARY;
    border: 1px solid $COLOR_PRIMARY;
    border-radius: $RADIUS_SM;
    color: $COLOR_TEXT_STRONG;
    font-weight: bold;
    padding: 7px $SPACE_MD;
    min-height: 22px;
}
QPushButton#ChromePrimary:hover {
    background-color: $COLOR_PRIMARY_LIFT;
    border-color: $COLOR_PRIMARY_LIFT;
}
QPushButton#ChromePrimary:disabled {
    background-color: $COLOR_CHROME_RAISED;
    border-color: $COLOR_CHROME_LINE;
    color: $COLOR_CHROME_FAINT;
}
QPushButton#ChromeButton {
    background-color: $COLOR_CHROME_RAISED;
    border: 1px solid $COLOR_CHROME_LINE;
    border-radius: $RADIUS_SM;
    color: $COLOR_CHROME_TEXT;
    padding: 5px $SPACE_MD;
}
QPushButton#ChromeButton:hover {
    background-color: $COLOR_CHROME_HOVER;
    border-color: $COLOR_ACCENT_BRIGHT;
}
QPushButton#ChromeButton:disabled {
    color: $COLOR_CHROME_FAINT;
}
QPushButton#ChromePrimary:focus, QPushButton#ChromeButton:focus {
    border: 1px solid $COLOR_ACCENT_BRIGHT;
}
QLineEdit#TopSearch {
    background-color: $COLOR_CHROME_SUNKEN;
    border: 1px solid $COLOR_CHROME_LINE;
    border-radius: $RADIUS_SM;
    padding: 4px $SPACE_SM;
    color: $COLOR_CHROME_TEXT;
    selection-background-color: $COLOR_PRIMARY;
    selection-color: $COLOR_TEXT_STRONG;
}
QLineEdit#TopSearch:focus {
    border: 1px solid $COLOR_ACCENT_BRIGHT;
    padding: 4px $SPACE_SM;
}

/* Sidebar navigation - flat, borderless, part of the chrome surface. The row
   text, index and badge are painted by NavDelegate; these rules own only the
   background states, because a delegate cannot read a QSS `color`. */
QListWidget#NavList {
    background: transparent;
    border: none;
    outline: 0;
}
QListWidget#NavList::item {
    border: none;
    border-radius: $RADIUS_SM;
}
QListWidget#NavList::item:selected {
    background-color: $COLOR_CHROME_ACTIVE;
}
QListWidget#NavList::item:hover:!selected {
    background-color: $COLOR_CHROME_HOVER;
}
QListWidget#NavList QScrollBar:vertical {
    background: transparent;
}

/* The guided-review ribbon. Hidden unless that mode is on, so it can afford to
   be chrome-toned: when it is visible it is a mode indicator, not content. */
QFrame#ReviewRibbon {
    background-color: $COLOR_CHROME_RAISED;
    border-bottom: 1px solid $COLOR_CHROME_LINE;
}
QLabel#RibbonTitle {
    font-size: $FONT_SM;
    font-weight: bold;
    color: $COLOR_CHROME_TEXT;
}
QLabel#RibbonSub {
    font-size: $FONT_XS;
    color: $COLOR_CHROME_MUTED;
}
QLabel#StageName {
    font-size: $FONT_SM;
    color: $COLOR_CHROME_MUTED;
}
QLabel#StageName[state="ok"]   { color: $COLOR_ACCENT_BRIGHT; }
QLabel#StageName[state="busy"] { color: $COLOR_CHROME_TEXT; font-weight: bold; font-style: normal; }

/* =========================================================================
   Workspace widgets
   ========================================================================= */

/* Tabs */
QTabWidget::pane {
    border: 1px solid $COLOR_BORDER_SUBTLE;
    background: $COLOR_SURFACE_RAISED;
}
QTabBar::tab {
    background: $COLOR_SURFACE_BASE;
    color: $COLOR_TEXT;
    padding: 8px $SPACE_LG;
    border-top-left-radius: $RADIUS_SM;
    border-top-right-radius: $RADIUS_SM;
    margin-right: 2px;
}
QTabBar::tab:selected {
    background: $COLOR_SURFACE_RAISED;
    border-bottom: 2px solid $COLOR_ACCENT;
    font-weight: bold;
}
QTabBar::tab:hover {
    background: $COLOR_SURFACE_OVERLAY;
}

/* Buttons */
QPushButton {
    background-color: $COLOR_SURFACE_RAISED;
    border: 1px solid $COLOR_BORDER;
    border-radius: $RADIUS_SM;
    padding: 6px $SPACE_MD;
    min-height: 25px;
}
QPushButton:hover {
    background-color: $COLOR_SURFACE_OVERLAY;
    border-color: $COLOR_ACCENT;
}
QPushButton:pressed {
    background-color: $COLOR_OVERLAY_HOVER;
}
QPushButton:disabled {
    background-color: $COLOR_SURFACE_OVERLAY;
    border-color: $COLOR_BORDER_SUBTLE;
    color: $COLOR_TEXT_DISABLED;
}
QPushButton#PrimaryButton {
    background-color: $COLOR_PRIMARY;
    border-color: $COLOR_PRIMARY;
    color: $COLOR_TEXT_STRONG;
    font-weight: bold;
}
QPushButton#PrimaryButton:hover {
    background-color: $COLOR_PRIMARY_HOVER;
    border-color: $COLOR_PRIMARY_HOVER;
}
QPushButton#PrimaryButton:disabled {
    background-color: $COLOR_SURFACE_OVERLAY;
    border-color: $COLOR_BORDER_SUBTLE;
    color: $COLOR_TEXT_DISABLED;
}
QPushButton#DangerButton {
    background-color: $COLOR_SURFACE_RAISED;
    border: 1px solid $COLOR_ERROR;
    color: $COLOR_ERROR;
}
QPushButton#DangerButton:hover {
    background-color: $COLOR_ERROR_FILL;
}
/* A quiet button for the actions that sit beside a primary one. */
QPushButton#GhostButton {
    background-color: transparent;
    border: 1px solid transparent;
    color: $COLOR_TEXT_MUTED;
}
QPushButton#GhostButton:hover {
    background-color: $COLOR_SURFACE_OVERLAY;
    color: $COLOR_TEXT;
    border-color: $COLOR_BORDER_SUBTLE;
}
/* The tab strip inside a panel: one row, one of them current. Checkable, so
   the state is the button's own and no code has to restyle its siblings. */
QPushButton#SegTab {
    background-color: transparent;
    border: 1px solid transparent;
    border-radius: $RADIUS_SM;
    color: $COLOR_TEXT_MUTED;
    font-weight: bold;
    padding: 5px $SPACE_MD;
    min-height: 20px;
}
QPushButton#SegTab:hover {
    background-color: $COLOR_SURFACE_OVERLAY;
    color: $COLOR_TEXT;
}
QPushButton#SegTab:checked {
    background-color: $COLOR_ACCENT_FILL;
    border-color: $COLOR_ACCENT;
    color: $COLOR_ACCENT;
}
QPushButton#SegTab:focus {
    border-color: $COLOR_ACCENT;
}

/* Inputs.
   QSpinBox / QDoubleSpinBox are deliberately absent from this rule. Giving a
   spin box a border and padding here makes the stylesheet own its layout and
   squeezes the native up/down buttons down to an unusable sliver - verified by
   rendering. They stay unstyled so Fusion draws them, with the palette
   supplying contrast. Same reason the subcontrols below are left alone. */
QLineEdit, QTextEdit, QComboBox {
    background-color: $COLOR_SURFACE_SUNKEN;
    border: 1px solid $COLOR_BORDER;
    border-radius: $RADIUS_SM;
    padding: $SPACE_XS;
    selection-background-color: $COLOR_PRIMARY;
    selection-color: $COLOR_TEXT_STRONG;
}
QLineEdit:hover, QTextEdit:hover, QComboBox:hover {
    border-color: $COLOR_ACCENT;
}
QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
    border: 2px solid $COLOR_ACCENT;
    padding: 3px;
}
/* The Report view's document pane. It is read-only rather than an input, so it
   takes a panel edge instead of a control frame - COLOR_BORDER is the weight
   reserved for something you type into. Mono because the report rules its
   headings with "=" and lines its counts up in columns: neither survives a
   proportional face. This is the only results surface in the app that is text
   rather than a table, because the text *is* the artifact. */
QPlainTextEdit#ReportText {
    background-color: $COLOR_SURFACE_SUNKEN;
    color: $COLOR_TEXT;
    border: 1px solid $COLOR_BORDER_SUBTLE;
    border-radius: $RADIUS_SM;
    padding: $SPACE_SM;
    font-family: $FONT_FAMILY_MONO;
    font-size: $FONT_SM;
    selection-background-color: $COLOR_PRIMARY;
    selection-color: $COLOR_TEXT_STRONG;
}
QPushButton:focus {
    border: 1px solid $COLOR_ACCENT;
}
QCheckBox {
    spacing: $SPACE_SM;
}
/* A combo's popup list. This is a child *widget* (a QListView), not a
   subcontrol, so styling it is safe - the rule below about ::drop-down does not
   apply here, and `QListWidget::item` is styled the same way further down.
   Unstyled it fell through to the base `QMainWindow, QWidget` rule and picked
   up COLOR_SURFACE_BASE, so the list was page-coloured while the closed combo
   beside it was white.

   The explicit item height is what gives the popup slack. Qt sizes the popup to
   the sum of the rows' size hints, and a font-derived hint left it at exactly
   2 x 24px for two items - no margin at all, so any rounding at 150% scaling
   takes a bite out of the last row. Same reasoning as fit_table_height: state
   the height rather than inferring it from the font. */
QComboBox QAbstractItemView {
    background-color: $COLOR_SURFACE_SUNKEN;
    border: 1px solid $COLOR_BORDER;
    border-radius: $RADIUS_SM;
    outline: 0;
    padding: 2px;
    selection-background-color: $COLOR_PRIMARY;
    selection-color: $COLOR_TEXT_STRONG;
}
QComboBox QAbstractItemView::item {
    min-height: $COMBO_ITEM_HEIGHT;
    padding: 0px $SPACE_XS;
    border-radius: $RADIUS_SM;
}
QComboBox QAbstractItemView::item:hover:!selected {
    background-color: $COLOR_ACCENT_FILL;
}

/* NOTE: QComboBox::drop-down, QComboBox::down-arrow, QCheckBox::indicator and
   the spin box ::up-button/::down-button are deliberately NOT styled here.
   Styling a Qt subcontrol makes the stylesheet responsible for drawing it, and
   a QSS rule cannot draw a glyph without an `image:` - border tricks render as
   rectangles. Styling them made the arrow and the checkmark disappear.
   Contrast is handled by the palette in apply_dark_palette() instead, which
   keeps the native arrow/checkmark and makes them independent of the user's
   Windows light/dark theme. */

/* Semantic label states. Set via set_state(); the property selector does
   nothing without the unpolish/polish pair in that helper. */
QLabel[state="ok"]    { color: $COLOR_SUCCESS; }
QLabel[state="error"] { color: $COLOR_ERROR; }
QLabel[state="warn"]  { color: $COLOR_WARNING; }
QLabel[state="busy"]  { color: $COLOR_TEXT_MUTED; font-style: italic; }
QLabel[state="muted"] { color: $COLOR_TEXT_MUTED; }

QLabel#SectionHeader {
    font-size: $FONT_LG;
    font-weight: bold;
    margin-top: $SPACE_MD;
    margin-bottom: $SPACE_XS;
    border-bottom: 1px solid $COLOR_BORDER_SUBTLE;
    padding-bottom: $SPACE_XS;
}
/* The heading of a card. Distinct from GroupLabel, which labels a group in the
   sidebar chrome - the same objectName was doing both jobs, so a card title and
   a sidebar caption could not be styled apart. */
QLabel#CardTitle {
    font-size: $FONT_SM;
    font-weight: bold;
    color: $COLOR_TEXT_MUTED;
}
QLabel#Caption {
    color: $COLOR_TEXT_MUTED;
    font-size: $FONT_XS;
}
QLabel#GroupLabel {
    color: $COLOR_TEXT_MUTED;
    font-size: $FONT_XS;
    font-weight: bold;
}
/* The line above a page title: where you are, and how far through. Never a
   colour on its own - it always spells the position out in words. */
QLabel#Eyebrow {
    font-size: $FONT_XS;
    font-weight: bold;
    color: $COLOR_TEXT_MUTED;
}
/* The page title carries the display face at a *lighter* weight than the body
   bold around it. Bold at this size shouts; the mockup sets it at 500 and the
   size alone does the work of establishing it as the heading. */
QLabel#ViewTitle {
    font-family: $FONT_FAMILY_DISPLAY;
    font-size: $FONT_DISPLAY;
    font-weight: 600;
}
/* One title per view, in the content header. Three views used to repeat it
   underneath in a second style, so a page read "Null Analysis" and then
   "Null Value Analysis". */
QLabel#ViewSubtitle {
    font-size: $FONT_SM;
    color: $COLOR_TEXT_MUTED;
}

/* Cards - one reusable surface. Padding comes from setContentsMargins in
   make_card(), because QSS padding does not move a laid-out QFrame's children.

   The 1px outline is not decoration. A card is 1.14:1 against the page behind
   it, which is imperceptible - tone cannot carry elevation at any hue, so the
   edge is what makes a card read as a card. */
QFrame#Card {
    background-color: $COLOR_SURFACE_RAISED;
    border: 1px solid $COLOR_BORDER_SUBTLE;
    border-radius: $RADIUS_LG;
}
QFrame#Card[severity="ok"]    { border-left: 3px solid $COLOR_SUCCESS; background-color: $COLOR_SUCCESS_FILL; }
QFrame#Card[severity="error"] { border-left: 3px solid $COLOR_ERROR;   background-color: $COLOR_ERROR_FILL; }
QFrame#Card[severity="warn"]  { border-left: 3px solid $COLOR_WARNING; background-color: $COLOR_WARNING_FILL; }
QFrame#Card[severity="info"]  { border-left: 3px solid $COLOR_ACCENT;  background-color: $COLOR_ACCENT_FILL; }

/* A summary tile: one number, one label. The value takes no colour here, so the
   QLabel[state=...] rules above can still reach it - an id selector would
   outrank them and the number would never turn red. */
QFrame#Kpi {
    background-color: $COLOR_SURFACE_RAISED;
    border: 1px solid $COLOR_BORDER_SUBTLE;
    border-radius: $RADIUS_MD;
}
QFrame#Kpi[severity="error"] { border-color: $COLOR_ERROR; }
QFrame#Kpi[severity="warn"]  { border-color: $COLOR_WARNING; }
QFrame#Kpi[severity="ok"]    { border-color: $COLOR_SUCCESS; }
/* A tile's number is the one thing on it worth reading at a glance, so it gets
   the display face. Sets no color, deliberately - see the note above. */
QLabel#KpiValue {
    font-family: $FONT_FAMILY_DISPLAY;
    font-size: $FONT_KPI;
    font-weight: 600;
}
QLabel#KpiLabel {
    font-size: $FONT_XS;
    color: $COLOR_TEXT_MUTED;
}

/* A status chip. The word inside it always says the state too, so the colour is
   reinforcement and never the only cue. */
QLabel#Chip {
    background-color: $COLOR_SURFACE_OVERLAY;
    border: 1px solid $COLOR_BORDER_SUBTLE;
    border-radius: $RADIUS_SM;
    color: $COLOR_TEXT_MUTED;
    font-size: $FONT_XS;
    font-weight: bold;
    padding: 2px $SPACE_SM;
}
QLabel#Chip[severity="ok"]    { background-color: $COLOR_SUCCESS_FILL; border-color: $COLOR_SUCCESS; color: $COLOR_SUCCESS; }
QLabel#Chip[severity="error"] { background-color: $COLOR_ERROR_FILL;   border-color: $COLOR_ERROR;   color: $COLOR_ERROR; }
QLabel#Chip[severity="warn"]  { background-color: $COLOR_WARNING_FILL; border-color: $COLOR_WARNING; color: $COLOR_WARNING; }
QLabel#Chip[severity="info"]  { background-color: $COLOR_ACCENT_FILL;  border-color: $COLOR_ACCENT;  color: $COLOR_ACCENT; }

/* The evidence inspector: a panel pinned to the right of a workbench view. */
QFrame#Inspector {
    background-color: $COLOR_SURFACE_RAISED;
    border-left: 1px solid $COLOR_BORDER_SUBTLE;
}
QFrame#InspectorHead {
    background-color: $COLOR_SURFACE_OVERLAY;
    border-bottom: 1px solid $COLOR_BORDER_SUBTLE;
}
QFrame#InspectorFoot {
    background-color: $COLOR_SURFACE_OVERLAY;
    border-top: 1px solid $COLOR_BORDER_SUBTLE;
}
QLabel#InspectorTitle {
    font-size: $FONT_LG;
    font-weight: bold;
}
QFrame#NoteBlock {
    background-color: $COLOR_ACCENT_FILL;
    border: none;
    border-left: 3px solid $COLOR_ACCENT;
    border-radius: $RADIUS_SM;
}
/* A toolbar strip inside a panel: the tab row, the filter row. */
QFrame#PanelBar {
    background-color: $COLOR_SURFACE_OVERLAY;
    border-bottom: 1px solid $COLOR_BORDER_SUBTLE;
}

/* Tables */
QTableView {
    background-color: $COLOR_SURFACE_SUNKEN;
    alternate-background-color: $COLOR_SURFACE_OVERLAY;
    gridline-color: $COLOR_BORDER_DIVIDER;
    selection-background-color: $COLOR_ACCENT_FILL;
    selection-color: $COLOR_TEXT;
    border: 1px solid $COLOR_BORDER_SUBTLE;
    border-radius: $RADIUS_SM;
    outline: 0;
}
/* A column head is a label, not content, so it sits a step below body size.
   The mockup sets these 7px uppercase against 9px body; Qt QSS supports neither
   text-transform nor letter-spacing, and uppercasing the strings themselves
   would put shouting headers into the TSV copy - which is the thing these
   tables exist to produce. Size and weight carry the distinction instead. */
QHeaderView::section {
    background-color: $COLOR_SURFACE_OVERLAY;
    padding: $SPACE_SM $SPACE_XS;
    border: none;
    font-size: $FONT_XS;
    font-weight: bold;
    color: $COLOR_TEXT_MUTED;
    border-bottom: 1px solid $COLOR_BORDER_SUBTLE;
}
QHeaderView::section:hover {
    color: $COLOR_TEXT;
}
QTableCornerButton::section {
    background-color: $COLOR_SURFACE_OVERLAY;
    border: none;
}

/* Lists */
QListWidget {
    background-color: $COLOR_SURFACE_SUNKEN;
    border: 1px solid $COLOR_BORDER;
    border-radius: $RADIUS_SM;
    outline: 0;
}
QListWidget::item {
    padding: 6px $SPACE_SM;
    border-bottom: 1px solid $COLOR_BORDER_DIVIDER;
}
QListWidget::item:selected {
    background-color: $COLOR_PRIMARY;
    color: $COLOR_TEXT_STRONG;
}
QListWidget::item:hover:!selected {
    background-color: $COLOR_SURFACE_OVERLAY;
}

/* Progress */
QProgressBar {
    height: $PROGRESS_HEIGHT;
    border: none;
    background: $COLOR_SURFACE_DEEP;
}
QProgressBar::chunk {
    background: $COLOR_ACCENT;
}

/* Scrollbars */
QScrollBar:vertical {
    border: none;
    background: transparent;
    width: $SCROLLBAR_SIZE;
    margin: 0px 0px 0px 0px;
}
QScrollBar::handle:vertical:hover {
    background: $COLOR_TEXT_DISABLED;
}
QScrollBar::handle:vertical {
    background: $COLOR_BORDER_SUBTLE;
    min-height: $SPACE_LG;
    border-radius: $RADIUS_SM;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
QScrollBar:horizontal {
    border: none;
    background: transparent;
    height: $SCROLLBAR_SIZE;
    margin: 0px 0px 0px 0px;
}
QScrollBar::handle:horizontal {
    background: $COLOR_BORDER_SUBTLE;
    min-width: $SPACE_LG;
    border-radius: $RADIUS_SM;
}
QScrollBar::handle:horizontal:hover {
    background: $COLOR_TEXT_DISABLED;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0px;
}
QScrollArea {
    border: none;
    background: transparent;
}

/* Splitter - it sits between the chrome and the workspace, so it takes the
   chrome's line colour rather than the page's. */
QSplitter::handle {
    background-color: $COLOR_CHROME_LINE;
}
QSplitter::handle:horizontal {
    width: 1px;
}
QSplitter::handle:vertical {
    height: 1px;
}

/* Menus and status bar. The menu bar is chrome; the popups it opens land on the
   workspace and are lit like it. */
QMenuBar {
    background-color: $COLOR_CHROME;
    color: $COLOR_CHROME_TEXT;
    border-bottom: 1px solid $COLOR_CHROME_LINE;
}
QMenuBar::item {
    padding: 5px $SPACE_SM;
    background: transparent;
}
QMenuBar::item:selected {
    background-color: $COLOR_CHROME_HOVER;
}
QMenu {
    background-color: $COLOR_SURFACE_RAISED;
    border: 1px solid $COLOR_BORDER_SUBTLE;
    padding: $SPACE_XS;
}
QMenu::item {
    padding: 6px $SPACE_LG;
    border-radius: $RADIUS_SM;
}
QMenu::item:selected {
    background-color: $COLOR_PRIMARY;
    color: $COLOR_TEXT_STRONG;
}
QMenu::separator {
    height: 1px;
    background: $COLOR_BORDER_SUBTLE;
    margin: $SPACE_XS 0px;
}
QStatusBar {
    background-color: $COLOR_SURFACE_DEEP;
    color: $COLOR_TEXT_MUTED;
    border-top: 1px solid $COLOR_BORDER_SUBTLE;
}
QStatusBar::item {
    border: none;
}
QToolTip {
    background-color: $COLOR_CHROME_RAISED;
    color: $COLOR_CHROME_TEXT;
    border: 1px solid $COLOR_CHROME_LINE;
    padding: $SPACE_XS;
}
QDialog {
    background-color: $COLOR_SURFACE_BASE;
}
"""


# ---------------------------------------------------------------------------
# Navigation
#
# NAV_GROUPS controls *placement and order only*. The authoritative registry of
# views remains SanityCheckApp.tab_defs, so adding a view there is still enough
# to make it appear - anything registered but not placed here is appended to a
# trailing "MORE" group rather than silently vanishing.
# ---------------------------------------------------------------------------
NAV_GROUPS = [
    ("REVIEW", ["Overview"]),
    ("STRUCTURE", ["Column Check", "Null Analysis", "Logic Checks"]),
    ("CONTENT", ["Language Check", "Description Quality"]),
    ("INSIGHTS", ["Monthly Inflow", "Pivot Tables", "Data Profile"]),
    ("DATA", ["Duplicate Check", "Raw Data", "Extract"]),
    ("DELIVER", ["Report"]),

]

# Registered views deliberately kept out of the nav list.
NON_NAV_VIEWS = {"Settings"}   # reached via the gear button and File > Settings


def _token_values() -> dict:
    """The design tokens, as strings, for template substitution."""
    values = {}
    for name, value in list(globals().items()):
        if name.startswith(("COLOR_", "RADIUS_", "SPACE_", "FONT_", "SCROLLBAR_",
                            "PROGRESS_", "TABLE_", "SIDEBAR_", "TOPBAR_",
                            "INSPECTOR_", "COMBO_")):
            values[name] = f"{value}px" if isinstance(value, int) else value
    return values


def _build_qss() -> str:
    """Substitute the design tokens into the stylesheet template."""
    return Template(_QSS_TEMPLATE).substitute(_token_values())


DARK_THEME_QSS = _build_qss()





def html(markup: str) -> str:
    """
    Fill $TOKEN placeholders in the Help and About markup.

    Those dialogs are rich text, so QSS cannot reach them - they carried their
    own inline CSS with hard-coded colours and font sizes, which meant the app
    had two themes and only one of them moved when a token changed. This is the
    same string.Template mechanism the stylesheet uses, for the same reason: the
    markup contains no '$', and .substitute raises on a typo instead of quietly
    leaving a placeholder in a user-visible dialog.
    """
    return Template(markup).substitute(_token_values())


class TicketAuditStyle(QProxyStyle):
    """
    Draws the check indicators, because neither QSS nor the palette can.

    Fusion derives a checkbox frame from QPalette::Window darkened, which on this
    theme lands at 1.10:1 against a card - an unchecked box was invisible, and
    measurably so: 153 of its 196 pixels were the card colour showing through a
    frame nobody could see. Raising a palette role does not help, because Fusion
    consults neither Mid nor Dark nor Shadow here.

    The QSS route is closed on purpose: `::indicator` cannot draw a glyph without
    an `image:`, and styling it erases the native one - the rule that
    tests/test_design_tokens.py enforces. A QProxyStyle is the supported way to
    change native painting, and it keeps every colour a design token.

    Everything not listed here falls through to Fusion untouched.
    """

    # Indicators are square; these are fractions of that square so the mark
    # scales with the widget rather than being pinned to one pixel size.
    _TICK = ((0.24, 0.53), (0.43, 0.72), (0.78, 0.28))

    def drawPrimitive(self, element, option, painter, widget=None):
        checkbox = QStyle.PrimitiveElement.PE_IndicatorCheckBox
        itemcheck = QStyle.PrimitiveElement.PE_IndicatorItemViewItemCheck
        radio = QStyle.PrimitiveElement.PE_IndicatorRadioButton
        if element in (checkbox, itemcheck, radio):
            self._draw_indicator(option, painter, round_shape=element == radio)
            return
        super().drawPrimitive(element, option, painter, widget)

    def _draw_indicator(self, option, painter, *, round_shape: bool) -> None:
        state = option.state
        on = bool(state & QStyle.StateFlag.State_On)
        partial = bool(state & QStyle.StateFlag.State_NoChange)
        enabled = bool(state & QStyle.StateFlag.State_Enabled)
        hover = bool(state & QStyle.StateFlag.State_MouseOver)

        if not enabled:
            edge, fill, mark = (COLOR_BORDER_SUBTLE, COLOR_SURFACE_SUNKEN,
                                COLOR_TEXT_DISABLED)
        elif on or partial:
            edge, fill, mark = COLOR_ACCENT, COLOR_PRIMARY, COLOR_TEXT_STRONG
        else:
            edge = COLOR_ACCENT if hover else COLOR_BORDER
            fill, mark = COLOR_SURFACE_SUNKEN, COLOR_TEXT
        # 1.5 inset keeps the 1.4px stroke inside option.rect, so the box does
        # not clip against the label baseline.
        box = QRectF(option.rect).adjusted(1.5, 1.5, -1.5, -1.5)

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setBrush(QColor(fill))
        painter.setPen(QPen(QColor(edge), 1.4))
        if round_shape:
            painter.drawEllipse(box)
        else:
            painter.drawRoundedRect(box, RADIUS_SM - 1, RADIUS_SM - 1)

        if on and round_shape:
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(mark))
            painter.drawEllipse(box.adjusted(box.width() * 0.28,
                                             box.height() * 0.28,
                                             -box.width() * 0.28,
                                             -box.height() * 0.28))
        elif partial:
            painter.setPen(QPen(QColor(mark), 1.8, Qt.PenStyle.SolidLine,
                                Qt.PenCapStyle.RoundCap))
            middle = box.center().y()
            painter.drawLine(QPointF(box.left() + box.width() * 0.26, middle),
                             QPointF(box.right() - box.width() * 0.26, middle))
        elif on:
            path = QPainterPath()
            first = True
            for fx, fy in self._TICK:
                point = QPointF(box.left() + box.width() * fx,
                                box.top() + box.height() * fy)
                if first:
                    path.moveTo(point)
                    first = False
                else:
                    path.lineTo(point)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.setPen(QPen(QColor(mark), 1.8, Qt.PenStyle.SolidLine,
                                Qt.PenCapStyle.RoundCap,
                                Qt.PenJoinStyle.RoundJoin))
            painter.drawPath(path)
        painter.restore()


def apply_dark_palette(app) -> None:
    """
    Give the Fusion style a dark palette to match DARK_THEME_QSS.

    Widget parts that Qt draws natively rather than from the stylesheet - the
    combo box drop-down arrow, check box checkmarks, spin box arrows - take
    their colours from the palette, not the QSS. Without this the app inherits
    whatever palette Windows supplies, so on a dark Windows theme those glyphs
    are drawn dark on dark and vanish. Setting the palette explicitly makes
    them render consistently regardless of the user's system theme.
    """
    # Installed here because every entry point - main.py, the __main__ block and
    # the test fixtures - already calls this function right after setStyle, so no
    # call site has to learn about the style.
    #
    # Constructed from the style *name*, never from app.style(): QProxyStyle takes
    # ownership of the instance it is given, while QApplication.setStyle deletes
    # the outgoing style. Handing it the live style therefore leaves the proxy
    # holding a freed pointer, which showed up as the test suite hanging rather
    # than as a clean crash. Re-wrapping is also skipped, since this function is
    # called more than once in a long-lived process and each call would add
    # another layer to the chain.
    current = app.style()
    if not isinstance(current, TicketAuditStyle):
        app.setStyle(TicketAuditStyle(current.objectName() or "Fusion"))

    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(COLOR_SURFACE_BASE))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(COLOR_TEXT))
    palette.setColor(QPalette.ColorRole.Base, QColor(COLOR_SURFACE_SUNKEN))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(COLOR_SURFACE_RAISED))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(COLOR_SURFACE_OVERLAY))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(COLOR_TEXT))
    palette.setColor(QPalette.ColorRole.Text, QColor(COLOR_TEXT))
    palette.setColor(QPalette.ColorRole.Button, QColor(COLOR_SURFACE_OVERLAY))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(COLOR_TEXT))
    palette.setColor(QPalette.ColorRole.BrightText, QColor(COLOR_ERROR))
    palette.setColor(QPalette.ColorRole.Link, QColor(COLOR_ACCENT))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(COLOR_PRIMARY))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(COLOR_TEXT_STRONG))
    # Unset by default, so placeholder contrast was unpredictable
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(COLOR_TEXT_MUTED))

    for role in (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text,
                 QPalette.ColorRole.ButtonText):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor(COLOR_TEXT_DISABLED))

    app.setPalette(palette)

    # Set on the QApplication, not the main window, so every dialog - Settings,
    # About, the four help windows, the error boxes - inherits it instead of
    # falling back to Qt's blank default. Same reason TicketAuditStyle is
    # installed here: every entry point already calls this right after setStyle.
    #
    # The .ico is the multi-resolution one from the brand kit, which is what
    # Windows wants: it picks 16px for the titlebar and 32/48 for the taskbar
    # and Alt-Tab out of the same file, rather than scaling one bitmap down and
    # turning the probe and the finding dot to mush.
    mark = app_icon()
    if not mark.isNull():
        app.setWindowIcon(mark)

    # Open a combo's list instantly, fully drawn.
    #
    # Qt animates the popup on Windows by revealing it downwards over ~100ms,
    # and every frame of that is a list with its last row sliced through
    # horizontally - which is indistinguishable from a clipped popup, and is
    # what it looks like in a screenshot. These pickers are how you choose a
    # date column or a month range, so the list is the answer to a question
    # rather than a flourish; there is nothing here worth animating, and the
    # animation's only visible product is a frame that looks broken.
    app.setEffectEnabled(Qt.UIEffect.UI_AnimateCombo, False)


def _asset(name: str) -> str:
    """Absolute path to a file in gui/assets."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", name)


def app_icon() -> QIcon:
    """
    The TicketAudit mark: a structured ticket list crossed by a green probe
    that ends in an audit check.

    Returns an empty QIcon rather than raising if the asset is missing. A
    missing icon is a cosmetic fault and must not stop the app opening - and it
    is not the integrity-checked logo, which is a separate file with a hash
    behind it (see show_about_dialog).
    """
    path = _asset("ticketaudit.ico")
    return QIcon(path) if os.path.exists(path) else QIcon()


def set_state(widget, state: Optional[str] = None) -> None:
    """
    Apply a semantic state to a widget, styled by the QLabel[state="..."] rules.

    state is one of "ok", "error", "warn", "busy", "muted", or None to clear.

    The unpolish/polish pair is required: Qt does not re-evaluate property
    selectors when a property changes, so without it the new colour never
    appears and the widget silently keeps the previous state's styling.
    """
    widget.setProperty("state", state)
    widget.style().unpolish(widget)
    widget.style().polish(widget)


def set_severity(card, sev: Optional[str] = None) -> None:
    """Update the severity property on a QFrame#Card, triggering its accent-bar style.

    sev is one of "ok", "error", "warn", "info", or None to show the plain surface.
    Requires the same unpolish/polish pair as set_state.
    """
    card.setProperty("severity", sev)
    card.style().unpolish(card)
    card.style().polish(card)


# Icons are drawn, not shipped. Qt's standard pixmaps (SP_DialogOpenButton and
# friends) are the platform's own colourful glyphs - a blue folder and a floppy
# disk - which look pasted on against a monochrome dark theme, and one of them
# was simply wrong: the Settings button used SP_FileDialogDetailedView, a
# "details view" glyph, not a gear.
#
# Painting them keeps every icon the same weight and the same colour as the text
# beside it, scales cleanly at 125% and 150%, and adds no binary assets to a
# repo whose one image is integrity-hashed. Each painter draws on a 24x24 grid
# and is scaled to the requested size.
ICON_STROKE = 1.9


def _icon_folder(painter):
    path = QPainterPath()
    path.moveTo(3, 7)
    path.lineTo(9.5, 7)
    path.lineTo(11.5, 9.5)
    path.lineTo(21, 9.5)
    path.lineTo(21, 19)
    path.lineTo(3, 19)
    path.closeSubpath()
    painter.drawPath(path)


def _icon_download(painter):
    painter.drawLine(QPointF(12, 4), QPointF(12, 14.5))
    head = QPainterPath()
    head.moveTo(7.5, 10.5)
    head.lineTo(12, 15.2)
    head.lineTo(16.5, 10.5)
    painter.drawPath(head)
    painter.drawLine(QPointF(4.5, 19.5), QPointF(19.5, 19.5))


def _icon_gear(painter):
    painter.drawEllipse(QRectF(9, 9, 6, 6))
    # Six teeth rather than eight: at 16px the eighth is a smudge.
    for index in range(6):
        angle = math.radians(index * 60)
        inner_x, inner_y = 12 + 8.2 * math.cos(angle), 12 + 8.2 * math.sin(angle)
        outer_x, outer_y = 12 + 10.5 * math.cos(angle), 12 + 10.5 * math.sin(angle)
        painter.drawLine(QPointF(inner_x, inner_y), QPointF(outer_x, outer_y))
    painter.drawEllipse(QRectF(4.6, 4.6, 14.8, 14.8))


def _icon_mail(painter):
    painter.drawRoundedRect(QRectF(3.5, 6, 17, 12), 1.5, 1.5)
    flap = QPainterPath()
    flap.moveTo(3.5, 7)
    flap.lineTo(12, 13)
    flap.lineTo(20.5, 7)
    painter.drawPath(flap)


def _icon_copy(painter):
    painter.drawRoundedRect(QRectF(8.5, 8.5, 12, 12), 2, 2)
    back = QPainterPath()
    back.moveTo(5.5, 15.5)
    back.lineTo(3.5, 15.5)
    back.lineTo(3.5, 3.5)
    back.lineTo(15.5, 3.5)
    back.lineTo(15.5, 5.5)
    painter.drawPath(back)


_ICON_PAINTERS = {
    "folder": _icon_folder,
    "download": _icon_download,
    "gear": _icon_gear,
    "mail": _icon_mail,
    "copy": _icon_copy,
}


def icon(name: str, colour: str = "", size: int = 16) -> QIcon:
    """
    A monochrome icon in the colour of the text it sits next to.

    Rendered at the screen's device pixel ratio, because a 16px pixmap upscaled
    to 150% is visibly soft next to text that is not.
    """
    painter_fn = _ICON_PAINTERS.get(name)
    if painter_fn is None:
        return QIcon()
    app = QApplication.instance()
    ratio = app.devicePixelRatio() if app is not None else 1.0
    pixels = max(1, int(round(size * ratio)))
    pixmap = QPixmap(pixels, pixels)
    pixmap.fill(Qt.GlobalColor.transparent)
    pixmap.setDevicePixelRatio(ratio)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    # size, not pixels: a pixmap carrying a devicePixelRatio already hands the
    # painter logical coordinates, so scaling by the device size applied the
    # ratio twice and the drawing overflowed its own pixmap - each icon came out
    # as the top-left fragment of itself.
    painter.scale(size / 24.0, size / 24.0)
    pen = QPen(QColor(colour or COLOR_TEXT), ICON_STROKE)
    pen.setCapStyle(Qt.PenCapStyle.RoundCap)
    pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    painter.setPen(pen)
    painter.setBrush(Qt.BrushStyle.NoBrush)
    painter_fn(painter)
    painter.end()
    return QIcon(pixmap)


def theme_chart(chart, *, category_axis=None, value_axis=None,
                top_value: float = 0) -> None:
    """
    Apply the app's chart style. Everything here removes something.

    What it takes away and why:

    * **Vertical gridlines.** A bar chart has categories, not a continuous x
      axis; ruling between them boxes every bar for no information. The old
      chart drew a full grid in near-white, which was the loudest thing in the
      window.
    * **The plot border and the shading behind it.** Two more rectangles that
      carry nothing.
    * **The chart title.** The view already has a title and a subtitle above it;
      a third line saying "Monthly Inflow - Created" is the same fact again.
    * **Non-round tick values.** Qt divides the data range, so a maximum of 496
      produced ticks at 124, 248 and 372. Rounded to a sensible step, the axis
      reads 0, 100, 200 and the bars can be compared against it by eye.

    Horizontal gridlines stay, at divider weight - they are what lets you read a
    bar's height at all.
    """
    chart.setBackgroundBrush(QBrush(QColor(COLOR_SURFACE_RAISED)))
    chart.setBackgroundPen(QPen(Qt.PenStyle.NoPen))
    chart.setPlotAreaBackgroundVisible(False)
    chart.setTitle("")
    chart.legend().setVisible(False)
    chart.setMargins(QMargins(SPACE_SM, SPACE_SM, SPACE_SM, SPACE_SM))

    if category_axis is not None:
        category_axis.setGridLineVisible(False)
        category_axis.setLineVisible(False)
        category_axis.setLabelsColor(QColor(COLOR_TEXT_MUTED))

    if value_axis is not None:
        value_axis.setLabelsColor(QColor(COLOR_TEXT_MUTED))
        value_axis.setGridLineColor(QColor(COLOR_BORDER_DIVIDER))
        value_axis.setLineVisible(False)
        value_axis.setLabelFormat("%d")
        if top_value > 0:
            step = _axis_step(top_value)
            # A whole step of headroom above the tallest bar. Without it the axis
            # top equals the data maximum, Qt has nowhere to put that bar's value
            # label, and it drops it *inside* the bar - grey on teal, and
            # inconsistent with the shorter bars whose labels sit outside.
            steps = int(math.ceil((top_value + step * 0.2) / step))
            value_axis.setRange(0, step * max(1, steps))
            value_axis.setTickInterval(step)
            value_axis.setTickType(QValueAxis.TickType.TicksDynamic)


def _axis_step(top_value: float) -> int:
    """
    A round tick interval for a value axis: 1, 2 or 5 times a power of ten.

    Aiming for four or five ticks - enough to read a height against, few enough
    that the labels do not become the busiest thing on the chart.
    """
    # Five, not four: /4 on a maximum of 496 chose a step of 200, which puts the
    # top of the axis at 600 and leaves three gridlines to read twelve bars
    # against. /5 chooses 100, so the axis runs 0..500 in five steps.
    rough = max(top_value, 1) / 5
    magnitude = 10 ** int(math.floor(math.log10(rough)))
    for multiple in (1, 2, 5, 10):
        step = magnitude * multiple
        if step >= rough:
            # Counts are integers, and int() truncates: a maximum of 1 gave a
            # step of 0.2 and therefore 0, which is not a usable interval.
            return max(1, int(round(step)))
    return max(1, int(round(magnitude * 10)))


def make_view(parent, *, scroll: bool = False) -> QVBoxLayout:
    """
    The standard skeleton for a view page. Returns the layout to fill.

    Every view used to set its own margins and spacing, or none at all: nine of
    the fourteen left both at Qt's defaults, so the vertical rhythm changed as
    you moved between them. This is the one place those numbers live.

    `scroll=True` wraps the content in a QScrollArea. Any view that pins a
    height - a fit_table_height table, a setMinimumHeight chart - needs it: a
    QVBoxLayout with less room than its size hints ask for takes the space from
    whatever can shrink, and lays the widgets that cannot on top of each other.
    That is not hypothetical, it is the bug that was reported in Language Check.

    Either way a layout is installed on `parent`, which show_view relies on as
    its not-yet-built sentinel.
    """
    if not scroll:
        layout = QVBoxLayout(parent)
        layout.setContentsMargins(SPACE_XS, SPACE_XS, SPACE_XS, SPACE_XS)
        layout.setSpacing(SPACE_SM)
        return layout

    outer = QVBoxLayout(parent)
    outer.setContentsMargins(0, 0, 0, 0)
    outer.setSpacing(0)
    area = QScrollArea()
    area.setWidgetResizable(True)
    area.setFrameShape(QFrame.Shape.NoFrame)
    inner = QWidget()
    area.setWidget(inner)
    outer.addWidget(area)
    layout = QVBoxLayout(inner)
    layout.setContentsMargins(SPACE_XS, SPACE_XS, SPACE_XS, SPACE_XS)
    layout.setSpacing(SPACE_SM)
    return layout


def add_intro(layout, text: str) -> QLabel:
    """
    The one-line explanation some views carry above their content.

    Muted, wrapped, and added the same way everywhere - it was a bare QLabel with
    no state in some views, a muted one in others, and absent in seven.
    """
    intro = QLabel(text)
    intro.setWordWrap(True)
    set_state(intro, "muted")
    layout.addWidget(intro)
    return intro


def make_card(parent_layout=None, horizontal: bool = False, margin: int = SPACE_MD):
    """
    Create a QFrame#Card with its layout, the app's one reusable panel surface.

    Padding comes from setContentsMargins rather than QSS `padding`, because
    Qt does not apply stylesheet padding to a QFrame that has a layout.
    Returns (frame, layout).
    """
    frame = QFrame()
    frame.setObjectName("Card")
    layout = QHBoxLayout(frame) if horizontal else QVBoxLayout(frame)
    layout.setContentsMargins(margin, margin, margin, margin)
    layout.setSpacing(SPACE_SM)
    if parent_layout is not None:
        parent_layout.addWidget(frame)
    return frame, layout


# Item data roles carried by a sidebar row. The name lives in UserRole, as it
# always has; the rest is drawn by NavDelegate and must not be part of the text,
# or the index and the badge would sort, elide and copy as part of the name.
NAV_ROLE_INDEX = Qt.ItemDataRole.UserRole + 1
NAV_ROLE_BADGE = Qt.ItemDataRole.UserRole + 2
NAV_ROLE_SEVERITY = Qt.ItemDataRole.UserRole + 3


class NavDelegate(QStyledItemDelegate):
    """
    Paints a sidebar row: a two-digit index, the view name, and a count badge.

    The background, selection and hover still come from the QSS `::item` rules -
    this blanks the text, hands the row to the style so those rules run, then
    draws the three pieces itself. A delegate cannot read a QSS `color`, so the
    text colours are named from the tokens here instead; that is the whole
    reason `QListWidget#NavList::item` sets no colour.

    A badge is a *count of findings*, so it is only ever set from a real number
    and is absent rather than zero - a grey "0" beside every clean view reads as
    a score, and there is deliberately no score in this app.
    """

    INDEX_WIDTH = 24
    BADGE_PAD = 6

    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        name = index.data(Qt.ItemDataRole.UserRole)
        text = opt.text
        opt.text = ""
        style = opt.widget.style() if opt.widget is not None else QApplication.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, opt, painter,
                          opt.widget)

        left = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        painter.save()
        if name is None:
            # A group caption: smaller, letter-spaced and bottom-aligned in its
            # taller row, so it cannot be mistaken for something clickable.
            font = QFont(opt.font)
            font.setPixelSize(FONT_XS)
            font.setBold(True)
            font.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 1.2)
            painter.setFont(font)
            painter.setPen(QColor(COLOR_CHROME_FAINT))
            painter.drawText(
                opt.rect.adjusted(SPACE_SM, 0, -SPACE_SM, -SPACE_XS),
                int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignBottom),
                text)
            painter.restore()
            return

        selected = bool(opt.state & QStyle.StateFlag.State_Selected)
        enabled = bool(opt.state & QStyle.StateFlag.State_Enabled)
        hover = bool(opt.state & QStyle.StateFlag.State_MouseOver)
        if not enabled:
            ink, dim = COLOR_CHROME_FAINT, COLOR_CHROME_FAINT
        elif selected:
            ink, dim = COLOR_CHROME_TEXT, COLOR_ACCENT_BRIGHT
        elif hover:
            ink, dim = COLOR_CHROME_TEXT, COLOR_CHROME_FAINT
        else:
            ink, dim = COLOR_CHROME_MUTED, COLOR_CHROME_FAINT

        rect = opt.rect.adjusted(SPACE_SM, 0, -SPACE_SM, 0)

        number = index.data(NAV_ROLE_INDEX) or ""
        if number:
            # Mono, so the two-digit numbers form a straight column. In a
            # proportional face "01" and "11" are different widths and the
            # labels beside them stop aligning.
            small = QFont(opt.font)
            small.setFamilies(NAV_INDEX_FAMILIES)
            small.setPixelSize(FONT_XS)
            painter.setFont(small)
            painter.setPen(QColor(dim))
            painter.drawText(QRect(rect.x(), rect.y(), self.INDEX_WIDTH,
                                   rect.height()), int(left), str(number))

        badge = index.data(NAV_ROLE_BADGE)
        body = QRect(rect.x() + (self.INDEX_WIDTH if number else 0), rect.y(),
                     rect.width() - (self.INDEX_WIDTH if number else 0),
                     rect.height())

        if badge:
            severity = index.data(NAV_ROLE_SEVERITY) or "warn"
            fill = {"error": COLOR_ERROR_BRIGHT,
                    "warn": COLOR_WARNING_BRIGHT}.get(severity, COLOR_CHROME_MUTED)
            small = QFont(opt.font)
            small.setPixelSize(FONT_XS)
            small.setBold(True)
            painter.setFont(small)
            width = painter.fontMetrics().horizontalAdvance(str(badge)) + \
                self.BADGE_PAD * 2
            height = FONT_MD + 2
            pill = QRectF(body.right() - width, body.y() + (body.height() - height) / 2,
                          width, height)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(fill))
            painter.drawRoundedRect(pill, RADIUS_SM, RADIUS_SM)
            painter.setPen(QColor(COLOR_CHROME))
            painter.drawText(pill, int(Qt.AlignmentFlag.AlignCenter), str(badge))
            body.setRight(int(pill.left()) - SPACE_XS)

        font = QFont(opt.font)
        font.setPixelSize(FONT_SM)
        font.setBold(selected)
        painter.setFont(font)
        painter.setPen(QColor(ink))
        painter.drawText(body, int(left), painter.fontMetrics().elidedText(
            text, Qt.TextElideMode.ElideRight, body.width()))
        painter.restore()


def make_kpi(parent_layout, label: str, value: str = "\u2014"):
    """
    A summary tile: one number, one label. Returns (frame, value_label).

    The headline of a view is a count, not a score - there is no cross-check
    score anywhere in core/, and inventing a weighting for one would mean
    inventing the judgement behind it. So a tile carries a number the analyzer
    actually produced, and set_state() on the returned label is what colours it.
    """
    frame = QFrame()
    frame.setObjectName("Kpi")
    inner = QVBoxLayout(frame)
    inner.setContentsMargins(SPACE_MD, SPACE_SM, SPACE_MD, SPACE_SM)
    inner.setSpacing(0)
    number = QLabel(value)
    number.setObjectName("KpiValue")
    inner.addWidget(number)
    caption = QLabel(label)
    caption.setObjectName("KpiLabel")
    inner.addWidget(caption)
    if parent_layout is not None:
        parent_layout.addWidget(frame)
    return frame, number


def make_chip(text: str, severity: Optional[str] = None) -> QLabel:
    """
    A status chip. The word inside always says the state as well, so the colour
    reinforces it and is never the only cue.
    """
    chip = QLabel(text)
    chip.setObjectName("Chip")
    chip.setAlignment(Qt.AlignmentFlag.AlignCenter)
    set_severity(chip, severity)
    return chip


def make_seg_tabs(parent_layout, labels: List[str], on_change) -> Dict[str, QPushButton]:
    """
    The tab strip inside a panel. Checkable buttons in an exclusive group, so
    the current one is the button's own state and no code has to restyle its
    siblings. `on_change` is called with the label when a different tab is
    chosen; the first is checked without firing it.
    """
    group = QButtonGroup(parent_layout.parentWidget() or None)
    group.setExclusive(True)
    buttons = {}
    for position, label in enumerate(labels):
        button = QPushButton(label)
        button.setObjectName("SegTab")
        button.setCheckable(True)
        button.setCursor(Qt.CursorShape.PointingHandCursor)
        group.addButton(button)
        parent_layout.addWidget(button)
        buttons[label] = button
        if position == 0:
            button.setChecked(True)
        button.clicked.connect(
            lambda _checked=False, name=label: on_change(name))
    # Held on the first button so the group is not garbage-collected: a
    # QButtonGroup with no Python reference stops enforcing exclusivity, and the
    # symptom is two tabs looking current at once.
    buttons[labels[0]]._group = group
    return buttons


class NumericAlignDelegate(QStyledItemDelegate):
    """
    Right-aligns cells that hold a number.

    Done in a delegate rather than per item because the models are built in a
    dozen places from plain QStandardItems, and a column of counts that is
    left-aligned cannot be compared down its own length - the digits do not line
    up. Alignment already set by the model is respected.
    """

    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        if index.data(Qt.ItemDataRole.TextAlignmentRole) is not None:
            return
        text = option.text
        if text and _looks_numeric(text):
            option.displayAlignment = (Qt.AlignmentFlag.AlignRight
                                       | Qt.AlignmentFlag.AlignVCenter)


def _looks_numeric(text: str) -> bool:
    """
    Whether a cell should be read as a number.

    Deliberately narrow: digits with the separators this app actually produces -
    thousands commas, a decimal point, a percent sign, a leading minus. Not a
    general parser, because "2025-08-01" and "1 - Critical" are not numbers and
    right-aligning them would be worse than leaving them alone.
    """
    stripped = text.strip().rstrip("%").replace(",", "")
    if stripped.startswith("-"):
        stripped = stripped[1:]
    if not stripped:
        return False
    return stripped.replace(".", "", 1).isdigit()


def configure_table(view, *, sortable: bool = True, select_rows: bool = True,
                    stretch_last: bool = False, row_numbers: bool = False,
                    resize: str = "contents", title: Optional[str] = None):
    """
    Apply the app's standard QTableView behaviour. Call at creation time.

    resize: "contents" | "interactive" | "stretch".
    title:  section name used in the report header when the table is copied.

    setEditTriggers is the important one - it is set on the *view*, which makes
    all tables read-only regardless of whether their model happens to expose
    editable items (QStandardItem does by default).

    Copy support is installed here rather than per table so that every table in
    the app is copyable by the same keys and menu. Read-only does not mean
    unselectable: results exist to be pasted into a ticket or a report.
    """
    view._copy_title = title
    view._copy_extra = {}
    _install_table_copy(view)
    view.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    view.setSelectionBehavior(
        QAbstractItemView.SelectionBehavior.SelectRows if select_rows
        else QAbstractItemView.SelectionBehavior.SelectItems
    )
    view.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
    view.setAlternatingRowColors(True)
    # The alternating fill separates rows; a full grid on top of it draws a box
    # around every cell, which is the single biggest reason these tables looked
    # like a spreadsheet paste rather than a report.
    view.setShowGrid(False)
    view.setItemDelegate(NumericAlignDelegate(view))
    view.setSortingEnabled(sortable)
    view.setCornerButtonEnabled(False)
    view.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
    view.setWordWrap(False)

    vheader = view.verticalHeader()
    vheader.setVisible(row_numbers)
    vheader.setDefaultSectionSize(TABLE_ROW_HEIGHT)

    hheader = view.horizontalHeader()
    # Left by default, and right for the numeric columns (set per column in
    # set_table_model). Centred headers over left-aligned text read as floating
    # above the column rather than belonging to it, and the sort indicator ends
    # up at the far edge, a long way from the word it applies to.
    hheader.setDefaultAlignment(Qt.AlignmentFlag.AlignLeft
                                | Qt.AlignmentFlag.AlignVCenter)
    hheader.setStretchLastSection(stretch_last)
    modes = {
        "contents": QHeaderView.ResizeMode.Interactive,
        "interactive": QHeaderView.ResizeMode.Interactive,
        "stretch": QHeaderView.ResizeMode.Stretch,
    }
    hheader.setSectionResizeMode(modes.get(resize, QHeaderView.ResizeMode.Interactive))
    # Sample rows rather than scanning every cell when sizing columns
    hheader.setResizeContentsPrecision(50)
    view._auto_resize = (resize == "contents")
    return view


def _table_headers(view) -> List[str]:
    model = view.model()
    return [str(model.headerData(c, Qt.Orientation.Horizontal,
                                 Qt.ItemDataRole.DisplayRole) or "")
            for c in range(model.columnCount())]


def _table_rows(view, selected_only: bool) -> List[List[str]]:
    """
    Read the table as text, in the order the user is looking at it.

    Goes through the model's DisplayRole rather than any source data, so a
    sorted table copies in its sorted order and a cell showing "72/100" copies
    as "72/100" - what was on screen is what lands in the report.
    """
    model = view.model()
    if model is None:
        return []

    cols = range(model.columnCount())
    if selected_only and view.selectionModel() and view.selectionModel().hasSelection():
        rows = sorted({i.row() for i in view.selectionModel().selectedIndexes()})
    else:
        rows = range(model.rowCount())

    out = []
    for r in rows:
        out.append([
            str(model.data(model.index(r, c), Qt.ItemDataRole.DisplayRole) or "")
            for c in cols
        ])
    return out


def _copy_table(view, *, selected_only: bool, with_headers: bool, with_context: bool):
    """
    Put the table on the clipboard as tab-separated text.

    TSV because the destination is Excel or an email: tabs paste as real
    columns, where a comma-separated line would land in one cell and a
    hand-drawn ASCII grid would have to be re-typed.
    """
    rows = _table_rows(view, selected_only)
    if not rows:
        return 0

    lines = []
    if with_context:
        window = view.window()
        provider = getattr(window, "table_copy_context", None)
        if callable(provider):
            lines.extend(provider(view))
            lines.append("")
    if with_headers:
        lines.append("\t".join(_table_headers(view)))
    lines.extend("\t".join(cells) for cells in rows)

    QApplication.clipboard().setText("\n".join(lines))
    return len(rows)


def _install_table_copy(view):
    """Ctrl+C and a right-click menu on a table view."""
    def copy_plain():
        _copy_table(view, selected_only=True, with_headers=False, with_context=False)

    def copy_headers():
        _copy_table(view, selected_only=True, with_headers=True, with_context=False)

    def copy_report():
        _copy_table(view, selected_only=False, with_headers=True, with_context=True)

    # QAction on the view, not a QShortcut on the window: several tables can be
    # visible at once and the copy must go to the one with focus.
    act_copy = QAction("Copy", view)
    act_copy.setShortcut(QKeySequence.StandardKey.Copy)
    act_copy.setShortcutContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
    act_copy.triggered.connect(copy_plain)
    view.addAction(act_copy)

    act_headers = QAction("Copy with column headers", view)
    act_headers.triggered.connect(copy_headers)
    view.addAction(act_headers)

    act_report = QAction("Copy whole table for report", view)
    act_report.setToolTip("Includes the file name, sheet and settings this result came from")
    act_report.triggered.connect(copy_report)
    view.addAction(act_report)

    view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
    view.setToolTip("Ctrl+C copies the selection · right-click to copy the "
                    "whole table with its file, sheet and settings")

    def show_menu(pos):
        menu = QMenu(view)
        menu.addAction(act_copy)
        menu.addAction(act_headers)
        menu.addSeparator()
        menu.addAction(act_report)
        menu.exec(view.viewport().mapToGlobal(pos))

    view.customContextMenuRequested.connect(show_menu)


def set_copy_context(view, **fields):
    """
    Attach result-specific provenance to a table or a findings label, shown when
    it is copied for a report. Call from the populate step, since the values
    describe that run (which column, which mode, how many rows were in scope).
    """
    view._copy_extra = {k: v for k, v in fields.items() if v not in (None, "")}


def _label_plain_text(label) -> str:
    """A label's text as plain text, with any markup resolved."""
    text = label.text()
    if label.textFormat() == Qt.TextFormat.RichText or "<" in text:
        document = QTextDocument()
        document.setHtml(text)
        return document.toPlainText()
    return text


def _copy_label(label, *, with_context: bool) -> None:
    """
    Put a label's text on the clipboard: the selection if there is one, the
    whole thing otherwise - the same rule the tables follow.
    """
    selected = label.selectedText()
    # Qt uses U+2029 for a line break inside selected label text.
    body = selected.replace(" ", "\n") if selected else _label_plain_text(label)
    if not body.strip():
        return

    lines = []
    if with_context:
        provider = getattr(label.window(), "table_copy_context", None)
        if callable(provider):
            lines.extend(provider(label))
            lines.append("")
    lines.append(body)
    QApplication.clipboard().setText("\n".join(lines))


def make_text_selectable(root) -> None:
    """
    Make every label under `root` - or `root` itself, if it is one - selectable
    and copyable.

    Read-only means non-editable, not non-extractable - the same reasoning that
    put Ctrl+C on every table. A finding the user cannot select is one they have
    to retype into a ticket, and the summary cards carry exactly the sentence
    worth pasting ("672 of 6,913 analyzed entries (9.7%) contain non-English
    text"). Applied to whole view pages from show_view, so no view can omit it
    and nothing has to be wired per label.

    TextSelectableByMouse without TextSelectableByKeyboard is deliberate: Qt
    raises focusPolicy to StrongFocus for the keyboard flag, which would put all
    ~70 labels into the tab order. Click focus is enough for Ctrl+C to reach the
    label the user is pointing at.
    """
    labels = list(root.findChildren(QLabel))
    if isinstance(root, QLabel):
        labels.append(root)

    for label in labels:
        if getattr(label, "_copy_installed", False):
            continue
        label._copy_installed = True

        flags = Qt.TextInteractionFlag.TextSelectableByMouse
        if label.openExternalLinks():
            flags |= Qt.TextInteractionFlag.LinksAccessibleByMouse
        label.setTextInteractionFlags(flags)

        act_copy = QAction("Copy", label)
        act_copy.setShortcut(QKeySequence.StandardKey.Copy)
        act_copy.setShortcutContext(
            Qt.ShortcutContext.WidgetWithChildrenShortcut)
        act_copy.triggered.connect(
            lambda _checked=False, w=label: _copy_label(w, with_context=False))
        label.addAction(act_copy)

        act_report = QAction("Copy for report", label)
        act_report.setToolTip(
            "Includes the file name, sheet and settings this result came from")
        act_report.triggered.connect(
            lambda _checked=False, w=label: _copy_label(w, with_context=True))
        label.addAction(act_report)

        label.setContextMenuPolicy(Qt.ContextMenuPolicy.ActionsContextMenu)


def fit_table_height(view, *, max_rows: Optional[int] = None) -> None:
    """
    Size a small table to its content, so it does not leave a large empty panel
    below one row.

    Measured from the vertical header's **actual** section sizes rather than
    arithmetic on TABLE_ROW_HEIGHT. setDefaultSectionSize is a request: Qt
    raises it to minimumSectionSize when the font is larger, which is what
    happens on a scaled display - so the constant under-measures and clips the
    last rows on exactly the machines where the text is biggest.

    A pinned height cannot shrink, so any view using this must be scrollable.
    A QVBoxLayout with less room than its size hints need takes the space from
    whatever *can* shrink, and lays the neighbouring widgets out on top of the
    table that cannot.

    There is no "cap only" variant, and the reason is worth recording because it
    looks like the obvious way to avoid needing a scroll area: setting only a
    maximum leaves the layout free to use the widget's *size hint*, which for a
    QTableView is a fixed ~192px regardless of content. So a two-row table still
    got 192px and an eight-row one was clipped to five - the opposite of fitting
    the content, in both directions.
    """
    model = view.model()
    rows = model.rowCount() if model is not None else 0
    shown = rows if max_rows is None else min(rows, max_rows)
    vheader = view.verticalHeader()
    if shown > 0:
        body = sum(vheader.sectionSize(r) for r in range(shown))
    else:
        body = vheader.defaultSectionSize()
    # The larger of the header's current and preferred height. height() alone is
    # whatever the header happens to be *right now*, which is its pre-stylesheet
    # size when this runs during a populate - so giving the header padding in the
    # QSS silently clipped the last row by exactly that padding.
    header = view.horizontalHeader()
    chrome = max(header.height(), header.sizeHint().height())
    # Reserve the horizontal scrollbar's height too. Whether one appears is not
    # knowable here - the viewport has not been laid out when a populate runs -
    # and if it does appear inside a fixed height it takes ~10px from the bottom
    # and clips the last row. Costing a 10px gap when no scrollbar shows is the
    # cheaper mistake.
    if view.horizontalScrollBarPolicy() != Qt.ScrollBarPolicy.ScrollBarAlwaysOff:
        chrome += view.horizontalScrollBar().sizeHint().height()
    view.setFixedHeight(body + chrome + 2 * view.frameWidth())


def set_table_model(view, model):
    """
    Attach a model and size the columns, clamped so one wide free-text column
    cannot push everything else off screen, then spread any slack.

    The spread is the point. Sizing to contents alone left a three-column table
    occupying a quarter of a maximised window with a void beside it, which is
    what made several views look unfinished. Slack is shared in proportion to
    the width each column already asked for, so a narrow code column does not
    end up as wide as a description.
    """
    view.setModel(model)
    if getattr(view, "_auto_resize", True):
        view.resizeColumnsToContents()
        header = view.horizontalHeader()
        for col in range(header.count()):
            if header.sectionSize(col) > TABLE_MAX_COL_WIDTH:
                header.resizeSection(col, TABLE_MAX_COL_WIDTH)
        _spread_table_slack(view)
    _align_numeric_headers(view, model)
    return model


def _align_numeric_headers(view, model) -> None:
    """
    Right-align the header of any column whose cells are numbers.

    Sampled from the first few rows rather than the whole column: this runs on
    every populate, including a 300k-row detail table, and a column does not
    change type a thousand rows down.
    """
    rows = min(model.rowCount(), 8)
    if rows == 0:
        return
    right = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
    for col in range(model.columnCount()):
        values = [model.data(model.index(row, col), Qt.ItemDataRole.DisplayRole)
                  for row in range(rows)]
        texts = [str(v) for v in values if v not in (None, "")]
        if texts and all(_looks_numeric(text) for text in texts):
            try:
                model.setHeaderData(col, Qt.Orientation.Horizontal, right,
                                    Qt.ItemDataRole.TextAlignmentRole)
            except (AttributeError, TypeError):
                # A read-only model - PandasModel among them - simply keeps the
                # default alignment. Not worth failing a populate over.
                return


def _spread_table_slack(view) -> None:
    """Give the columns any width the viewport has left over."""
    header = view.horizontalHeader()
    count = header.count()
    if count == 0:
        return
    widths = [header.sectionSize(col) for col in range(count)]
    used = sum(widths)
    # Short of the viewport by a scrollbar's width, deliberately.
    #
    # The width measured here is not always the final one: a table inside a
    # scroll area is populated before that area knows whether it needs its own
    # scrollbar, and when one appears the table narrows. Filling the width
    # exactly then pushed the columns over the edge, raised a *horizontal*
    # scrollbar inside the table, and that scrollbar took 10px off the viewport
    # height - clipping the last row of every table whose height is pinned by
    # fit_table_height. Leaving that much slack unused costs a thin gap on the
    # right and removes the whole failure mode.
    available = view.viewport().width() - SCROLLBAR_SIZE - 2
    # Only ever widens, and only when it is worth doing. If the content is
    # already wider than the viewport the table scrolls sideways, which is right
    # for a 27-column detail sheet.
    if used <= 0 or available - used < SPACE_MD:
        return
    slack = available - used
    for col in range(count):
        share = round(slack * widths[col] / used)
        header.resizeSection(col, widths[col] + share)


_LANG_METHOD_LABELS = {
    LanguageChecker.METHOD_CHARACTERS: "Non-English script",
    LanguageChecker.METHOD_KEYWORDS: "Keyword pattern",
    LanguageChecker.METHOD_LINGUA: "Lingua model",
}


def _lang_confidence_text(method: str, conf: float) -> str:
    """
    Describe a detection's confidence in terms its method can support.

    Only Lingua produces a real probability. A script match is definitive and a
    keyword match is a fixed 0.90 heuristic, so printing either as a percentage
    invited the reader to compare numbers that do not mean the same thing.
    """
    if method == LanguageChecker.METHOD_CHARACTERS:
        return "Definitive"
    if method == LanguageChecker.METHOD_KEYWORDS:
        return "Heuristic"
    return f"{conf * 100:.0f}%"


class PandasModel(QAbstractTableModel):
    """A model to interface a Qt view with pandas dataframe."""

    def __init__(self, df: pd.DataFrame = pd.DataFrame()):
        QAbstractTableModel.__init__(self)
        self._df = df

    def rowCount(self, parent=QModelIndex()) -> int:
        return self._df.shape[0]

    def columnCount(self, parent=QModelIndex()) -> int:
        return self._df.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None

        if role == Qt.ItemDataRole.DisplayRole:
            value = self._df.iloc[index.row(), index.column()]
            # str(NaN) is the literal "nan", which reads as data
            return "" if pd.isna(value) else str(value)

        return None

    def flags(self, index):
        """Explicitly read-only; results are not user-editable."""
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable

    def headerData(self, col, orientation, role):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return self._df.columns[col]
        # Show the original row label, which identifies the source row
        return str(self._df.index[col])

    def sort(self, column, order=Qt.SortOrder.AscendingOrder):
        """
        Required for setSortingEnabled to do anything. Without it the header
        shows a sort indicator that silently does nothing.
        """
        if self._df.empty or column < 0 or column >= self._df.shape[1]:
            return
        self.beginResetModel()
        try:
            self._df = self._df.sort_values(
                self._df.columns[column],
                ascending=(order == Qt.SortOrder.AscendingOrder),
                kind="stable",
                na_position="last",
            )
        except TypeError:
            # Mixed types in one column - fall back to string comparison
            self._df = self._df.reindex(
                self._df.iloc[:, column].astype(str).sort_values(
                    ascending=(order == Qt.SortOrder.AscendingOrder),
                    kind="stable",
                ).index
            )
        finally:
            self.endResetModel()

    def setDataFrame(self, df):
        self.beginResetModel()
        self._df = df
        self.endResetModel()


# The register's four outcomes, worst first, and the state each maps to. This
# is the same vocabulary reporter.py sorts by, so the screen and the workbook
# cannot disagree about what outranks what.
RESULT_STATES = {"Fail": "error", "Not checked": "warn", "Review": "warn",
                 "Pass": "ok"}

# The severity words a logic rule can carry, plus the two outcomes that take a
# severity column's place when there is no finding to rank. "Not checked" is
# amber and never green, here as everywhere.
SEVERITY_STATES = {"Critical": "error", "High": "error", "Medium": "warn",
                   "Low": "warn", "Review": "warn", "Not checked": "warn",
                   "Pass": "ok"}

_STATE_COLOURS = {"error": COLOR_ERROR, "warn": COLOR_WARNING,
                  "ok": COLOR_SUCCESS}


class FindingsModel(PandasModel):
    """
    A results table whose outcome column carries a colour.

    A colour, not *only* a colour: the cell already says "Fail", "Review",
    "Pass", "Not checked" or a severity word, so the tint reinforces something
    legible without it. Done in the model rather than a delegate because
    configure_table installs NumericAlignDelegate on every table, and that
    delegate already honours ForegroundRole.
    """

    def __init__(self, df=pd.DataFrame(), colour_column="Result",
                 colours=None):
        super().__init__(df)
        self._colour_column = colour_column
        self._colours = colours if colours is not None else RESULT_STATES

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        # The last column gets the leftover width, which on a small window is
        # not enough for a sentence. Eliding is fine as long as the whole thing
        # is still reachable - and it is also in the copy, which is what would
        # land in a ticket.
        if role == Qt.ItemDataRole.ToolTipRole and index.isValid():
            value = self._df.iloc[index.row(), index.column()]
            return None if pd.isna(value) else str(value)
        if (role == Qt.ItemDataRole.ForegroundRole and index.isValid()
                and self._colour_column in self._df.columns):
            column = self._df.columns[index.column()]
            if column == self._colour_column:
                value = str(self._df.iloc[index.row(), index.column()])
                state = self._colours.get(value)
                if state:
                    return QBrush(QColor(_STATE_COLOURS[state]))
            return None
        return super().data(index, role)


class WorkerSignals(QObject):
    """Define the signals available from a running worker thread."""
    finished = Signal()
    error = Signal(str)
    result = Signal(object)
    progress = Signal(str)


class Worker(QRunnable):
    """Worker thread for running tasks in background."""

    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    @Slot()
    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs)
            self.signals.result.emit(result)
        except Exception as e:
            self.signals.error.emit(str(e))
        finally:
            self.signals.finished.emit()


class SanityCheckApp(QMainWindow):
    """Main Application Window."""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("TicketAudit - ITSM Ticket Data QA")
        self.setGeometry(100, 100, 1400, 900)
        
        # Load Config
        self.config_mgr = ConfigManager()
        self.config = self.config_mgr.config
        
        # Data & Core Objects
        self.df: Optional[pd.DataFrame] = None
        self.filepath: Optional[str] = None
        self.analyzer: Optional[SanityAnalyzer] = None
        self.lang_checker = LanguageChecker()
        # Completed Language Check results, keyed by column, so the exported
        # report can include them without paying for detection again (~26s per
        # 50k rows with Lingua). Cleared when a new sheet is loaded.
        self.language_results: Dict[str, Any] = {}
        # The Extract column ticks for this sheet, once the user has touched
        # them. None means "fall back to the remembered or standard set".
        self._extract_ticks: Optional[List[str]] = None
        # Day coverage per month for the chosen date column. Rebuilt when that
        # column changes rather than on every keystroke - see month_coverage.
        self._extract_coverage: List[Any] = []
        # Months ticked by hand this session, once the user has touched them.
        self._extract_month_ticks: Optional[List[str]] = None
        # Per-column fill shares for this sheet, measured once on first use.
        self._extract_fill: Optional[Dict[str, float]] = None
        # Set by the Cancel button and read by the writing worker.
        self._extract_cancelled = False
        # Cancelled-state values and their row mask for this sheet.
        self._extract_cancelled_info: Optional[Dict[str, Any]] = None
        # Where the last extract went, for the Open folder button.
        self._extract_last_path: Optional[str] = None
        # Whether the user has asked for the findings register on this sheet.
        # False means Overview shows its prompt: the register runs every check
        # in the app, so opening a file must not start it. See refresh_overview.
        self._overview_started = False
        # The same gate for the Report view, and for the same reason: both its
        # generators run every check in the app. See refresh_report.
        self._report_started = False
        # Which flavour the strip is on. make_seg_tabs checks the first button,
        # so this has to be the first entry or the two would disagree before the
        # user has touched anything.
        self._report_flavour = self.REPORT_FLAVOURS[0]
        # The flavour actually in the pane, and the report text itself - the
        # Copy and Save actions read these rather than re-deriving anything.
        self._report_flavour_shown: Optional[str] = None
        self._report_text: Optional[str] = None
        # Thread Pool
        self.threadpool = QThreadPool()
        self._active_workers = []  # Keep references to prevent GC

        # Async-refresh bookkeeping - see _refresh_async
        self._data_seq = 0                      # bumped on every sheet load
        self._refresh_inflight: set = set()     # one pending refresh per view
        # Guards one analyzer's lazily-built _cache. Replaced whenever the
        # analyzer is, so a slow refresh on the previous sheet cannot block the
        # new sheet's refresh - they share no cache and must not contend.
        self._analyzer_lock = threading.Lock()
        
        # Guided review is a mode and starts off. The set is per-session on
        # purpose: "I have looked at this" is a statement about the sitting, and
        # a persisted tick would carry a judgement onto a file it never saw.
        self._review_mode = False
        self._reviewed: set = set()

        # UI Setup
        self.init_ui()
        self.setup_views()
        
        # Apply Stylesheet
        self.setStyleSheet(DARK_THEME_QSS)

        # Restore size/position/sidebar width from the last session
        self._restore_window_state()

    def init_ui(self):
        """
        Build the shell: dark chrome wrapped around a light workspace.

        The sidebar, top bar and menu bar are one surface family and the page
        between them is the other - see the token block for why. Nothing here
        paints a result; every view owns a page in self.stack.
        """
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Sidebar is user-resizable, but must never collapse to zero - that
        # would hide the navigation entirely with no way back.
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setHandleWidth(1)
        self.splitter.addWidget(self._build_sidebar())
        self.splitter.addWidget(self._build_workspace())
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)
        main_layout.addWidget(self.splitter)
        self.setCentralWidget(main_widget)

        # App status bar - the single surface for transient status. It repeats
        # the dataset facts on purpose: on a maximised window the sidebar block
        # is a long way from the table being read.
        self.status_bar = QStatusBar()
        # The grip draws a diagonal scratch into the corner of a maximised
        # window, where it cannot do anything anyway.
        self.status_bar.setSizeGripEnabled(False)
        self.dataset_badge = QLabel("No data loaded")
        set_state(self.dataset_badge, "muted")
        self.status_bar.addPermanentWidget(self.dataset_badge)
        self.setStatusBar(self.status_bar)

        self.setup_menu_bar()

    # --- Shell pieces -----------------------------------------------------

    def _build_sidebar(self) -> QFrame:
        self.sidebar = QFrame()
        self.sidebar.setObjectName("Sidebar")
        self.sidebar.setMinimumWidth(200)
        self.sidebar.setMaximumWidth(SIDEBAR_MAX_WIDTH)
        column = QVBoxLayout(self.sidebar)
        column.setContentsMargins(SPACE_MD, SPACE_MD, SPACE_MD, SPACE_MD)
        column.setSpacing(SPACE_MD)

        # The brand mark beside the wordmark, at BRAND_MARK_SIZE.
        #
        # This is not the logo that used to be here. That PNG was a dark-blue
        # mark on transparency, so it needed a near-white plate to be visible -
        # making the brightest object in the window a decoration - and cost
        # 130px of sidebar height, which is what clipped the last nav item. It
        # still lives on the About dialog with its integrity check.
        #
        # This one is drawn for a dark shell: its own field is COLOR_CHROME and
        # its edge COLOR_CHROME_LINE, the two tokens this panel is already made
        # of, so it needs no plate and reads as part of the surface. Taken from
        # the .ico via QIcon.pixmap so Qt picks the nearest authored size and
        # scales for the device pixel ratio, rather than resampling one bitmap.
        brand = QWidget()
        brand.setObjectName("SidebarBrand")
        brand_row = QHBoxLayout(brand)
        brand_row.setContentsMargins(0, 0, 0, 0)
        brand_row.setSpacing(SPACE_SM)
        mark = QLabel()
        mark.setObjectName("BrandMark")
        mark.setFixedSize(BRAND_MARK_SIZE, BRAND_MARK_SIZE)
        mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # Not named `icon`: that is the module-level glyph helper, which this
        # same function calls further down for the export and settings buttons.
        brand_icon = app_icon()
        if brand_icon.isNull():
            # Never a blank square: the wordmark is beside it either way, but an
            # empty plate reads as a failed image where two letters read as a
            # mark. Same fallback the icon() helper takes for a missing glyph.
            mark.setText("TA")
        else:
            mark.setPixmap(
                brand_icon.pixmap(QSize(BRAND_MARK_SIZE, BRAND_MARK_SIZE)))
        brand_row.addWidget(mark)
        wordmark = QLabel("TicketAudit")
        wordmark.setObjectName("Wordmark")
        brand_row.addWidget(wordmark)
        brand_row.addStretch()
        column.addWidget(brand)

        # What is loaded, in the place you look for it.
        self.dataset_block = QFrame()
        self.dataset_block.setObjectName("DatasetBlock")
        block = QVBoxLayout(self.dataset_block)
        block.setContentsMargins(SPACE_SM, SPACE_SM, SPACE_SM, SPACE_SM)
        block.setSpacing(2)
        kicker = QLabel("ACTIVE DATASET")
        kicker.setObjectName("DatasetKicker")
        block.addWidget(kicker)
        self.file_label = QLabel("No file selected")
        self.file_label.setObjectName("DatasetName")
        # One line, elided, rather than wrapped to two. A long export name cost
        # 15px of the nav column here to say what the breadcrumb and the status
        # bar both also say - and set_dataset_badge writes all three, so nothing
        # is lost. The full name stays reachable as the tooltip.
        self.file_label.setWordWrap(False)
        block.addWidget(self.file_label)
        self.dataset_meta = QLabel("Open a file to begin")
        self.dataset_meta.setObjectName("DatasetMeta")
        self.dataset_meta.setWordWrap(True)
        block.addWidget(self.dataset_meta)
        # Deliberately left with the workspace input styling. A chrome-toned
        # combo would need its drop-down arrow to be light, and that arrow is
        # painted natively from the palette - the one thing a per-widget style
        # cannot reach without styling ::down-arrow, which erases it.
        self.combo_sheets = QComboBox()
        self.combo_sheets.setPlaceholderText("Select Sheet")
        self.combo_sheets.setEnabled(False)
        # Hidden until a workbook supplies sheets. An empty, disabled picker on
        # the start screen reads as a control the user has failed to use, and
        # load_file_async hides it again for delimited text, which has no sheets.
        self.combo_sheets.setVisible(False)
        self.combo_sheets.currentTextChanged.connect(self.load_sheet)
        block.addSpacing(SPACE_XS)
        block.addWidget(self.combo_sheets)
        column.addWidget(self.dataset_block)

        # Navigation takes the flexible space, and at 1280x720 - a 14-inch
        # 1080p panel at 150% scaling, which is the machine this is used on -
        # all fourteen views now fit without it scrolling. It is still a list,
        # and still scrolls, because a shorter window or a wider font can
        # always exceed the column; what changed is that the common case no
        # longer hides three views behind a scrollbar nobody looks for.
        self.nav_list = QListWidget()
        self.nav_list.setObjectName("NavList")
        self.nav_list.setFrameShape(QFrame.Shape.NoFrame)
        self.nav_list.setTextElideMode(Qt.TextElideMode.ElideRight)
        self.nav_list.setWordWrap(False)
        self.nav_list.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.nav_list.setItemDelegate(NavDelegate(self.nav_list))
        self.nav_list.currentItemChanged.connect(self.on_nav_changed)
        column.addWidget(self.nav_list, 1)

        footer = QWidget()
        footer.setObjectName("SidebarFooter")
        footer_row = QHBoxLayout(footer)
        footer_row.setContentsMargins(0, 0, 0, 0)
        footer_row.setSpacing(SPACE_SM)
        self.btn_export = QPushButton("Export report")
        self.btn_export.setObjectName("ChromePrimary")
        self.btn_export.setIcon(icon("download", COLOR_TEXT_STRONG))
        self.btn_export.setEnabled(False)
        self.btn_export.setToolTip("Write the full Excel report (Ctrl+E)")
        self.btn_export.clicked.connect(self.export_report)
        footer_row.addWidget(self.btn_export, 1)

        self.btn_settings = QPushButton()
        self.btn_settings.setObjectName("ChromeButton")
        self.btn_settings.setIcon(icon("gear", COLOR_CHROME_MUTED))
        self.btn_settings.setToolTip("Settings (Ctrl+,)")
        self.btn_settings.setFixedWidth(36)
        self.btn_settings.clicked.connect(self.show_settings_dialog)
        footer_row.addWidget(self.btn_settings)
        column.addWidget(footer)
        return self.sidebar

    def _build_workspace(self) -> QWidget:
        workspace = QWidget()
        outer = QVBoxLayout(workspace)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        outer.addWidget(self._build_topbar())
        outer.addWidget(self._build_review_ribbon())

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(SPACE_LG, SPACE_MD, SPACE_LG, SPACE_SM)
        content_layout.setSpacing(SPACE_XS)

        # Header shows *where you are*, and says it once. Three views used to
        # print their own heading underneath this one in a different style, so a
        # page read "Null Analysis" and then "Null Value Analysis".
        # Transient status goes to the status bar.
        self.view_eyebrow = QLabel("")
        self.view_eyebrow.setObjectName("Eyebrow")
        self.view_eyebrow.setVisible(False)
        content_layout.addWidget(self.view_eyebrow)
        self.view_title = QLabel("")
        self.view_title.setObjectName("ViewTitle")
        content_layout.addWidget(self.view_title)
        self.view_subtitle = QLabel("")
        self.view_subtitle.setObjectName("ViewSubtitle")
        self.view_subtitle.setWordWrap(True)
        content_layout.addWidget(self.view_subtitle)
        content_layout.addSpacing(SPACE_SM)

        # One page per view; page 0 is the empty state.
        self.stack = QStackedWidget()
        self.stack.setSizePolicy(QSizePolicy.Policy.Expanding,
                                 QSizePolicy.Policy.Expanding)
        content_layout.addWidget(self.stack, 1)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(False)
        content_layout.addWidget(self.progress_bar)

        outer.addWidget(content, 1)
        return workspace

    def _build_topbar(self) -> QFrame:
        bar = QFrame()
        bar.setObjectName("TopBar")
        bar.setFixedHeight(TOPBAR_HEIGHT)
        row = QHBoxLayout(bar)
        row.setContentsMargins(SPACE_LG, 0, SPACE_MD, 0)
        row.setSpacing(SPACE_SM)

        # Two labels rather than one rich-text line. The breadcrumb is chrome
        # rather than a result, but markup is still markup, and a bold file name
        # is the only thing it would buy.
        self.crumb_file = QLabel("No data")
        self.crumb_file.setObjectName("DatasetName")
        row.addWidget(self.crumb_file)
        self.crumb_rest = QLabel("")
        self.crumb_rest.setObjectName("Breadcrumb")
        row.addWidget(self.crumb_rest)
        row.addStretch()

        # A real filter over the sidebar, not a search-shaped decoration.
        # Fifteen views under six captions is past the point where scanning
        # beats typing.
        self.nav_search = QLineEdit()
        self.nav_search.setObjectName("TopSearch")
        # Short enough to survive the 140px minimum without eliding; the
        # shortcut is in the tooltip, where it does not compete for width.
        self.nav_search.setPlaceholderText("Filter views")
        self.nav_search.setToolTip("Narrow the sidebar to matching views (Ctrl+K)")
        self.nav_search.setMinimumWidth(140)
        self.nav_search.setMaximumWidth(220)
        # The placeholder role is set globally for the workspace, where it
        # measures 5.7:1; on the chrome input it needs the chrome's own muted.
        chrome_palette = self.nav_search.palette()
        chrome_palette.setColor(QPalette.ColorRole.PlaceholderText,
                                QColor(COLOR_CHROME_MUTED))
        self.nav_search.setPalette(chrome_palette)
        self.nav_search.textChanged.connect(self._filter_nav)
        self.nav_search.returnPressed.connect(self._open_first_nav_match)
        row.addWidget(self.nav_search)

        self.btn_review = QPushButton("Guided review")
        self.btn_review.setObjectName("ChromeButton")
        self.btn_review.setToolTip(
            "Walk the checks in the order a reviewer needs them, and keep "
            "track of what you have already looked at")
        self.btn_review.clicked.connect(self.toggle_guided_review)
        row.addWidget(self.btn_review)

        self.btn_load = QPushButton("Switch data")
        self.btn_load.setObjectName("ChromeButton")
        self.btn_load.setToolTip("Open another file (Ctrl+O)")
        self.btn_load.clicked.connect(self.browse_file)
        row.addWidget(self.btn_load)
        return bar

    def _build_review_ribbon(self) -> QFrame:
        """
        The guided-review ribbon, hidden unless that mode is on.

        Which is the point of it being a mode: a progress bar nobody asked for
        is a claim that they are behind on something. Each stage carries its own
        reviewed/total count, so the state is a number rather than only a
        colour.
        """
        ribbon = QFrame()
        ribbon.setObjectName("ReviewRibbon")
        ribbon.setVisible(False)
        row = QHBoxLayout(ribbon)
        row.setContentsMargins(SPACE_LG, SPACE_SM, SPACE_MD, SPACE_SM)
        row.setSpacing(SPACE_LG)

        heading = QVBoxLayout()
        heading.setSpacing(0)
        title = QLabel("Guided review")
        title.setObjectName("RibbonTitle")
        heading.addWidget(title)
        self.review_hint = QLabel("")
        self.review_hint.setObjectName("RibbonSub")
        heading.addWidget(self.review_hint)
        row.addLayout(heading)

        self.review_stage_labels = {}
        for stage_name, _views in REVIEW_STAGES:
            label = QLabel(stage_name)
            label.setObjectName("StageName")
            self.review_stage_labels[stage_name] = label
            row.addWidget(label)
        row.addStretch()

        self.review_progress = QLabel("")
        self.review_progress.setObjectName("RibbonTitle")
        row.addWidget(self.review_progress)

        self.btn_review_step = QPushButton("Mark reviewed")
        self.btn_review_step.setObjectName("ChromePrimary")
        self.btn_review_step.setToolTip(
            "Record that you have looked at this view, and move to the next")
        self.btn_review_step.clicked.connect(self.mark_step_reviewed)
        row.addWidget(self.btn_review_step)

        exit_button = QPushButton("Exit review")
        exit_button.setObjectName("ChromeButton")
        exit_button.clicked.connect(self.toggle_guided_review)
        row.addWidget(exit_button)

        self.review_ribbon = ribbon
        return ribbon

    # --- Sidebar filtering ------------------------------------------------

    def _filter_nav(self, text: str) -> None:
        """
        Narrow the sidebar to matching views, hiding captions left with none.

        Never changes the selection: the page on screen is not what was being
        searched for, and closing it under the user would lose their place.
        """
        needle = text.strip().lower()
        captions = []
        caption, shown = None, 0
        for row in range(self.nav_list.count()):
            item = self.nav_list.item(row)
            name = item.data(Qt.ItemDataRole.UserRole)
            if name is None:
                if caption is not None:
                    captions.append((caption, shown))
                caption, shown = item, 0
                continue
            match = not needle or needle in name.lower()
            item.setHidden(not match)
            shown += int(match)
        if caption is not None:
            captions.append((caption, shown))
        for item, count in captions:
            item.setHidden(count == 0)

    def _open_first_nav_match(self) -> None:
        """Enter in the filter box opens the first row still showing."""
        for row in range(self.nav_list.count()):
            item = self.nav_list.item(row)
            if item.isHidden() or item.data(Qt.ItemDataRole.UserRole) is None:
                continue
            self.nav_list.setCurrentRow(row)
            return

    def focus_nav_search(self) -> None:
        self.nav_search.setFocus(Qt.FocusReason.ShortcutFocusReason)
        self.nav_search.selectAll()

    # --- Guided review ----------------------------------------------------

    def _review_steps(self) -> List[str]:
        """The walk, as view names, skipping anything not registered."""
        return [name for _stage, names in REVIEW_STAGES for name in names
                if name in self.view_defs]

    def toggle_guided_review(self) -> None:
        """
        Turn the mode on or off. On entry it jumps to the first step not yet
        marked reviewed, which is what makes it a walk rather than a banner.
        """
        self._review_mode = not self._review_mode
        self.review_ribbon.setVisible(self._review_mode)
        self.btn_review.setText("Guided review" if not self._review_mode
                                else "Review active")
        if self._review_mode:
            steps = self._review_steps()
            nxt = next((s for s in steps if s not in self._reviewed), None)
            if nxt and nxt != self.current_view:
                self._select_view(nxt)
        self._update_review_ribbon()
        self._update_view_header()

    def mark_step_reviewed(self) -> None:
        """Record the current view as looked at, then move to the next one."""
        if not self.current_view:
            return
        self._reviewed.add(self.current_view)
        steps = self._review_steps()
        nxt = next((s for s in steps if s not in self._reviewed), None)
        if nxt:
            self._select_view(nxt)
        else:
            self.update_status("Guided review complete - every step is marked "
                               "reviewed.")
        self._update_review_ribbon()
        self._update_view_header()

    def _select_view(self, name: str) -> None:
        """Move the sidebar selection, which is what actually shows a view."""
        for row in range(self.nav_list.count()):
            if self.nav_list.item(row).data(Qt.ItemDataRole.UserRole) == name:
                self.nav_list.setCurrentRow(row)
                return

    def _update_review_ribbon(self) -> None:
        if not self._review_mode:
            return
        steps = self._review_steps()
        done = [s for s in steps if s in self._reviewed]
        for stage_name, names in REVIEW_STAGES:
            label = self.review_stage_labels.get(stage_name)
            if label is None:
                continue
            live = [n for n in names if n in self.view_defs]
            marked = sum(1 for n in live if n in self._reviewed)
            label.setText(f"{stage_name} {marked}/{len(live)}")
            if live and marked == len(live):
                set_state(label, "ok")
            elif self.current_view in live:
                set_state(label, "busy")
            else:
                set_state(label, None)
        share = round(100 * len(done) / len(steps)) if steps else 0
        self.review_progress.setText(
            f"{share}%  -  {len(done)} of {len(steps)} reviewed")
        remaining = len(steps) - len(done)
        self.review_hint.setText(
            "Every step reviewed" if not remaining
            else f"{remaining} step{'s' if remaining != 1 else ''} left. "
                 "Nothing here changes your data.")
        current_done = self.current_view in self._reviewed
        self.btn_review_step.setText("Reviewed" if current_done
                                     else "Mark reviewed")
        self.btn_review_step.setEnabled(not current_done
                                        and self.current_view is not None)

    def _update_view_header(self) -> None:
        """
        The eyebrow above the page title.

        In guided review it is the position in the walk; outside it, the sidebar
        group the view sits in, which is orientation rather than progress. A
        view outside the walk says so instead of being given a step number it
        does not have.
        """
        name = self.current_view
        if not name:
            self.view_eyebrow.setVisible(False)
            return
        if self._review_mode:
            steps = self._review_steps()
            if name in steps:
                stage = next((s for s, names in REVIEW_STAGES if name in names),
                             "")
                text = f"{stage}  /  Step {steps.index(name) + 1} of {len(steps)}"
            else:
                text = "Not part of the guided review"
        else:
            text = self._nav_group_of(name)
        self.view_eyebrow.setText(text)
        self.view_eyebrow.setVisible(bool(text))

    def _nav_group_of(self, name: str) -> str:
        for group_name, names in NAV_GROUPS:
            if name in names:
                return group_name
        return "MORE"

    # --- Navigation badges ------------------------------------------------

    def set_nav_badge(self, name: str, count: Optional[int] = None,
                      severity: str = "warn") -> None:
        """
        Put a finding count on a sidebar row, or clear it with count=None.

        Absent rather than zero: a grey "0" beside every clean view reads as a
        score, and there is deliberately no score in this app. Nothing sets a
        badge speculatively - it is only ever written from a check that ran.
        """
        item = self._nav_item(name)
        if item is None:
            return
        item.setData(NAV_ROLE_BADGE, None if not count else f"{count:,}")
        item.setData(NAV_ROLE_SEVERITY, severity)

    def clear_nav_badges(self) -> None:
        for row in range(self.nav_list.count()):
            item = self.nav_list.item(row)
            if item.data(Qt.ItemDataRole.UserRole):
                item.setData(NAV_ROLE_BADGE, None)

    def _nav_item(self, name: str):
        for row in range(self.nav_list.count()):
            item = self.nav_list.item(row)
            if item.data(Qt.ItemDataRole.UserRole) == name:
                return item
        return None

    def setup_menu_bar(self):
        """Setup the application menu bar."""
        menubar = self.menuBar()
        
        # File Menu
        file_menu = menubar.addMenu("&File")
        
        open_action = QAction("&Open Data File...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.browse_file)
        file_menu.addAction(open_action)

        file_menu.addSeparator()

        self.export_action = QAction("&Export Report...", self)
        self.export_action.setShortcut("Ctrl+E")
        self.export_action.setEnabled(False)   # nothing to export until data loads
        self.export_action.triggered.connect(self.export_report)
        file_menu.addAction(self.export_action)

        file_menu.addSeparator()

        find_action = QAction("&Find a View", self)
        find_action.setShortcut("Ctrl+K")
        find_action.triggered.connect(self.focus_nav_search)
        file_menu.addAction(find_action)

        review_action = QAction("&Guided Review", self)
        review_action.setShortcut("Ctrl+R")
        review_action.triggered.connect(self.toggle_guided_review)
        file_menu.addAction(review_action)

        file_menu.addSeparator()

        settings_action = QAction("&Settings...", self)
        settings_action.setShortcut("Ctrl+,")
        settings_action.triggered.connect(self.show_settings_dialog)
        file_menu.addAction(settings_action)

        file_menu.addSeparator()
        
        exit_action = QAction("E&xit", self)
        exit_action.setShortcut("Alt+F4")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Help Menu
        help_menu = menubar.addMenu("&Help")
        
        user_guide_action = QAction("&User Guide", self)
        user_guide_action.setShortcut("F1")
        user_guide_action.triggered.connect(self.show_user_guide_dialog)
        help_menu.addAction(user_guide_action)
        
        shortcuts_action = QAction("&Keyboard Shortcuts", self)
        shortcuts_action.triggered.connect(self.show_shortcuts_dialog)
        help_menu.addAction(shortcuts_action)
        
        help_menu.addSeparator()
        
        about_action = QAction("&About TicketAudit", self)
        about_action.setMenuRole(QAction.MenuRole.NoRole)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

    def show_about_dialog(self):
        """Show about dialog with integrity protection."""
        import hashlib
        import base64
        
        # Encoded author info (base64) - harder to find/modify via text search
        _a = base64.b64decode(b'QW5lZWsgSGFpdA==').decode()  # Author
        _n = base64.b64decode(b'VGlja2V0QXVkaXQ=').decode()  # App name
        _v = base64.b64decode(b'Mi4yLjE=').decode()  # Version
        _y = base64.b64decode(b'MjAyNS0yMDI2').decode()  # Year
        # Public links for the community edition.
        _site = "https://aneekhait.github.io"
        _gh = "https://github.com/AneekHait"
        _li = "https://www.linkedin.com/in/aneekhait/"
        # The interpreter actually running, rather than a literal that goes
        # stale - the hardcoded one already said 3.13 on a 3.14 install.
        _py = "%d.%d.%d" % sys.version_info[:3]

        # Integrity check - if someone modifies the values, hash won't match
        _check = hashlib.sha256(f"{_a}{_n}{_v}".encode()).hexdigest()[:8]
        _expected = "xx3554d44c"[2:10]  # Obfuscated expected hash
        
        if _check != _expected:
            QMessageBox.critical(self, "Integrity Error", 
                "Application files have been modified.\nPlease reinstall from official source.")
            return
        
        about_text = f"""
<div style="text-align: center;">
<h2 style="color: $COLOR_ACCENT; margin: 0;">{_n}</h2>
<p style="color: $COLOR_TEXT_MUTED; font-size: $FONT_SM; margin: 2px 0 8px 0;"><i>Know your queue before you report it</i></p>
<p style="margin: 5px 0; font-size: $FONT_XS;">
<b>Version:</b> {_v} &nbsp;|&nbsp; <b>Created by:</b> {_a} &nbsp;|&nbsp; <b>Release:</b> Aug 2026
</p>
<table align="center" cellspacing="8" cellpadding="0" style="margin: 10px auto 4px auto;">
<tr>
<td align="center" bgcolor="$COLOR_ACCENT_FILL" style="padding: 7px 18px; border: 1px solid $COLOR_ACCENT; border-radius: $RADIUS_MD;">
<a href="{_site}" title="{_site}" style="color: $COLOR_ACCENT; text-decoration: none; font-size: $FONT_SM;"><b>Website</b></a>
</td>
<td align="center" bgcolor="$COLOR_ACCENT_FILL" style="padding: 7px 18px; border: 1px solid $COLOR_ACCENT; border-radius: $RADIUS_MD;">
<a href="{_gh}" title="{_gh}" style="color: $COLOR_ACCENT; text-decoration: none; font-size: $FONT_SM;"><b>GitHub</b></a>
</td>
<td align="center" bgcolor="$COLOR_ACCENT_FILL" style="padding: 7px 18px; border: 1px solid $COLOR_ACCENT; border-radius: $RADIUS_MD;">
<a href="{_li}" title="{_li}" style="color: $COLOR_ACCENT; text-decoration: none; font-size: $FONT_SM;"><b>LinkedIn</b></a>
</td>
</tr>
</table>
</div>
<hr style="margin: 8px 0;">

<p style="font-size: $FONT_XS; margin: 5px 0;"><b>What is {_n}?</b><br>
An offline desktop app for data quality analysis on Excel/CSV files. Fourteen
checks over one sheet, every one of them exportable, and nothing leaves the
machine.</p>

<table style="font-size: $FONT_XS; width: 100%;">
<tr><td style="vertical-align: top; padding-right: 10px;">
<p style="margin: 5px 0;"><b>Review</b></p>
<ul style="margin: 2px 0 5px 15px; padding: 0;">
<li>Overview - every check, worst first, with what was never looked at</li>
<li>Guided Review - an optional walk through the checks in order</li>
</ul>
<p style="margin: 5px 0;"><b>Structure</b></p>
<ul style="margin: 2px 0 5px 15px; padding: 0;">
<li>Column Check - required fields, and the score behind each match</li>
<li>Null Analysis - how empty each column is, against your threshold</li>
<li>Logic Checks - eight date and cross-field rules, with the rows behind each</li>
</ul>
<p style="margin: 5px 0;"><b>Content</b></p>
<ul style="margin: 2px 0 5px 15px; padding: 0;">
<li>Language Check - non-English text, and which detection tiers ran</li>
<li>Description Quality - too short, too long, or repeated across tickets</li>
</ul>
</td><td style="vertical-align: top;">
<p style="margin: 5px 0;"><b>Insights</b></p>
<ul style="margin: 2px 0 5px 15px; padding: 0;">
<li>Monthly Inflow - ticket volume by month</li>
<li>Pivot Tables - how one column's values are distributed</li>
<li>Data Profile - type, fill and cardinality per column</li>
</ul>
<p style="margin: 5px 0;"><b>Data</b></p>
<ul style="margin: 2px 0 5px 15px; padding: 0;">
<li>Duplicate Check - records sharing an ID</li>
<li>Raw Data - the sheet as loaded</li>
<li>Extract - write a smaller file: chosen months, chosen columns</li>
</ul>
<p style="margin: 5px 0;"><b>Export</b></p>
<ul style="margin: 2px 0 5px 15px; padding: 0;">
<li>Excel Report - six tabs with native PivotTables and charts</li>
</ul>
</td></tr>
</table>

<hr style="margin: 8px 0;">

<p style="font-size: $FONT_XS; margin: 5px 0;"><b>Version History</b></p>
<table style="font-size: $FONT_XS; width: 100%;">
<tr style="background: $COLOR_SURFACE_RAISED;"><td style="padding: 4px;"><b>v2.2.1</b> (Aug 2026)</td><td style="padding: 4px;">Cross-platform support (macOS, Linux), shell launcher, screenshots in README</td></tr>
<tr style="background: $COLOR_SURFACE_RAISED;"><td style="padding: 4px;"><b>v2.2.0</b> (Aug 2026)</td><td style="padding: 4px;">Hybrid shell with sidebar navigation, the Overview register, Logic Checks as a finding-first workbench, guided review, and a three-face type system</td></tr>
<tr style="background: $COLOR_SURFACE_RAISED;"><td style="padding: 4px;"><b>v2.1.0</b> (Aug 2026)</td><td style="padding: 4px;">Extract view, six-tab Excel report with native pivots, correctable column mapping and date order, cross-field logic checks</td></tr>
<tr style="background: $COLOR_SURFACE_RAISED;"><td style="padding: 4px;"><b>v2.0.0</b> (Jan 2026)</td><td style="padding: 4px;">PySide6 GUI, new branding, Help system, QtCharts, performance optimizations</td></tr>
<tr><td style="padding: 4px;"><b>v1.0.0</b> (2025)</td><td style="padding: 4px;">Initial release with CustomTkinter, basic sanity checks, LLM chat</td></tr>
</table>

<hr style="margin: 8px 0;">
<p style="color: $COLOR_TEXT_MUTED; font-size: $FONT_XS; text-align: center; margin: 5px 0;">
<b>Tech:</b> Python {_py} • PySide6 • Pandas • Lingua<br>
Created by {_a} • Community Edition • Free to use and share.
</p>
"""
        # Create custom dialog for better sizing control
        dialog = QDialog(self)
        dialog.setWindowTitle(f"About {_n}")
        dialog.setModal(True)
        
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # Everything except the OK button scrolls. Without this the dialog sized
        # itself to its content and came out 725px tall on a 672px screen, which
        # put OK below the bottom edge - the window could only be dismissed with
        # Escape, and nothing on screen said so.
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(SPACE_SM)
        scroll.setWidget(body)
        
        # The brand lockup: the mark, the wordmark and the tagline, still
        # hash-checked, because this remains the only thing verifying the asset.
        #
        # The *light* lockup, because this dialog is a workspace surface - the
        # dark variant sets its wordmark in COLOR_CHROME_TEXT, which on white is
        # invisible. Rendered by scripts/render_brand.py rather than taken from
        # mockups/brand/renders/, whose two lockup PNGs have their text rendered
        # as tofu boxes by something that had no fonts installed. Re-running
        # that script means updating the digest below.
        _lp = os.path.join(os.path.dirname(__file__), "assets",
                           "ticketaudit-lockup.png")
        _lh = "x1e8174f92be93e7f7c4a1ce494120e984dd3adcded5158ff64a3267a9baebffd"[1:]
        if os.path.exists(_lp):
            with open(_lp, 'rb') as _f:
                if hashlib.sha256(_f.read()).hexdigest() == _lh:
                    mark = QLabel()
                    mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    mark.setPixmap(QPixmap(_lp).scaled(
                        260, 92, Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation))
                    body_layout.addWidget(mark)

        content = QLabel(html(about_text))
        content.setWordWrap(True)
        content.setTextFormat(Qt.TextFormat.RichText)
        content.setAlignment(Qt.AlignmentFlag.AlignTop)
        # Without this the website/GitHub/LinkedIn anchors render but do nothing
        # when clicked. make_text_selectable keeps them clickable while making
        # the surrounding text copyable, the same pairing the User Guide uses.
        content.setOpenExternalLinks(True)
        make_text_selectable(content)
        body_layout.addWidget(content)
        body_layout.addStretch()
        layout.addWidget(scroll, 1)
        
        # OK button
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        ok_btn = QPushButton("OK")
        ok_btn.setFixedWidth(80)
        ok_btn.clicked.connect(dialog.accept)
        btn_layout.addWidget(ok_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        # Sized to the content, then held inside the screen it will open on.
        # adjustSize alone is what produced a window taller than the display;
        # the cap is the fix, and the scroll area above is what makes the cap
        # non-destructive.
        dialog.resize(ABOUT_WIDTH, ABOUT_HEIGHT)
        available = (dialog.screen() or QApplication.primaryScreen())\
            .availableGeometry()
        dialog.resize(
            min(dialog.width(), available.width() - ABOUT_SCREEN_MARGIN),
            min(dialog.height(), available.height() - ABOUT_SCREEN_MARGIN))

        dialog.exec()

    def show_user_guide_dialog(self):
        """Show the User Guide dialog."""
        guide_text = """
<h2 style="color: $COLOR_ACCENT; margin: 0 0 10px 0;">User Guide</h2>

<h3 style="color: $COLOR_TEXT;">Getting Started</h3>
<ol>
<li><b>Load your data:</b> Click <i>"Browse Data File"</i> or press <b>Ctrl+O</b> to open an Excel (.xlsx/.xls) or delimited text (.csv/.tsv/.txt) file</li>
<li><b>Explore checks:</b> Pick a check from the sidebar to examine your data</li>
<li><b>Export results:</b> Click <i>"Export Report"</i> or press <b>Ctrl+E</b> to save findings</li>
</ol>

<h3 style="color: $COLOR_TEXT;">Analysis Views Explained</h3>
<table style="font-size: $FONT_XS; width: 100%;">
<tr><td style="vertical-align: top; width: 50%; padding-right: 15px;">
<p><b>Column Check</b><br>Verifies if required columns exist. Select a category to see expected vs found columns.</p>
<p><b>Null Analysis</b><br>Shows missing values per column. Red = above threshold (default 10%).</p>
<p><b>Logic Checks</b><br>Validates date sequences (e.g., Created Date &lt; Resolved Date) and
future dates, then cross-checks fields against each other: closed tickets with no resolution date,
open tickets that already have one, open tickets with no priority, and state values that look like
typos or non-standard entries. A rule that cannot run (missing column) is reported as
&quot;not checked&quot; rather than as a pass.</p>
<p><b>Language Check</b><br>Detects non-English text using AI language detection.</p>
</td><td style="vertical-align: top;">
<p><b>Monthly Inflow</b><br>Charts record creation trends by month. Great for spotting data patterns.</p>
<p><b>Pivot Tables</b><br>Quick value distribution analysis. Select a column to see counts.</p>
<p><b>Data Profile</b><br>Comprehensive statistics for each column (type, unique values, min/max).</p>
<p><b>Duplicate Check</b><br>Finds duplicate records by ID column.</p>
<p><b>Extract</b><br>Writes a smaller file out of this one - choose a month range and the
columns you need. For exports too large for Excel to filter without hanging. The size
updates as you choose, so you can see whether the result will open comfortably before
writing it. CSV is near-instant; .xlsx is slower but opens on a double-click.</p>
</td></tr>
</table>

<h3 style="color: $COLOR_TEXT;">Export Options</h3>
<ul>
<li><b>Excel Report:</b> Multi-sheet workbook with all analysis results</li>
</ul>

<h3 style="color: $COLOR_TEXT;">Settings</h3>
<ul>
<li><b>Null Threshold:</b> Adjust the percentage above which null counts are flagged</li>
</ul>
"""
        self._show_help_dialog("User Guide", html(guide_text), 580, 520)

    def show_shortcuts_dialog(self):
        """Show keyboard shortcuts dialog."""
        shortcuts_text = """
<h2 style="color: $COLOR_ACCENT; margin: 0 0 15px 0;">Keyboard Shortcuts</h2>

<table style="font-size: $FONT_SM; width: 100%;">
<tr style="background: $COLOR_SURFACE_RAISED;"><th style="padding: 8px; text-align: left;">Shortcut</th><th style="padding: 8px; text-align: left;">Action</th></tr>
<tr><td style="padding: 6px;"><b>Ctrl + O</b></td><td style="padding: 6px;">Open Excel/CSV file</td></tr>
<tr style="background: $COLOR_SURFACE_SUNKEN;"><td style="padding: 6px;"><b>Ctrl + E</b></td><td style="padding: 6px;">Export Report</td></tr>
<tr><td style="padding: 6px;"><b>Ctrl + K</b></td><td style="padding: 6px;">Filter the sidebar to find a view</td></tr>
<tr style="background: $COLOR_SURFACE_SUNKEN;"><td style="padding: 6px;"><b>Ctrl + R</b></td><td style="padding: 6px;">Start or leave the guided review</td></tr>
<tr><td style="padding: 6px;"><b>Ctrl + ,</b></td><td style="padding: 6px;">Open Settings</td></tr>
<tr style="background: $COLOR_SURFACE_SUNKEN;"><td style="padding: 6px;"><b>F1</b></td><td style="padding: 6px;">Open User Guide</td></tr>
<tr><td style="padding: 6px;"><b>Alt + F4</b></td><td style="padding: 6px;">Exit Application</td></tr>
</table>

<p style="color: $COLOR_TEXT_MUTED; margin-top: 15px; font-size: $FONT_XS;">
<b>Tip:</b> Use <b>Tab</b> and <b>Shift+Tab</b> to navigate between UI elements.<br>
<b>Tip:</b> <b>Ctrl + C</b> copies any table or finding; right-click for
"Copy for report", which adds the file, sheet and row counts.
</p>
"""
        self._show_help_dialog("Keyboard Shortcuts", html(shortcuts_text), 460, 400)

    def _show_help_dialog(self, title: str, html_content: str, width: int, height: int):
        """Helper to show a help dialog with scrollable content."""
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setModal(True)
        
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # Scrollable content
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        content = QLabel(html_content)
        content.setWordWrap(True)
        content.setTextFormat(Qt.TextFormat.RichText)
        content.setAlignment(Qt.AlignmentFlag.AlignTop)
        content.setOpenExternalLinks(True)
        # Selectable, with the links still clickable - the guide holds model
        # names and config keys that are meant to be copied out.
        make_text_selectable(content)
        scroll.setWidget(content)
        
        layout.addWidget(scroll)
        
        # OK button
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        ok_btn = QPushButton("OK")
        ok_btn.setFixedWidth(80)
        ok_btn.clicked.connect(dialog.accept)
        btn_layout.addWidget(ok_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        # Size based on screen
        screen = self.screen().availableGeometry()
        final_width = min(width, int(screen.width() * 0.6))
        final_height = min(height, int(screen.height() * 0.8))
        dialog.setFixedSize(final_width, final_height)
        
        dialog.exec()

    def setup_views(self):
        """Build the view registry, the sidebar navigation and the stacked pages."""
        self.tab_widgets = {}
        self.view_indices = {}
        self.current_view = None
        self._settings_dialog = None

        # Authoritative registry: name -> creator. Adding an entry here is what
        # makes a view exist; NAV_GROUPS only decides where it appears.
        self.tab_defs = [
            ("Overview", self.create_overview_tab),
            ("Column Check", self.create_column_check_tab),
            ("Null Analysis", self.create_null_analysis_tab),
            ("Logic Checks", self.create_logic_checks_tab),
            ("Language Check", self.create_language_check_tab),
            ("Description Quality", self.create_description_quality_tab),
            ("Monthly Inflow", self.create_monthly_inflow_tab),
            ("Pivot Tables", self.create_pivot_tables_tab),
            ("Data Profile", self.create_data_profile_tab),
            ("Duplicate Check", self.create_duplicate_check_tab),
            ("Raw Data", self.create_raw_data_tab),
            ("Extract", self.create_extract_tab),
            ("Report", self.create_report_tab),
            ("Settings", self.create_settings_tab)
        ]
        # Name -> creator lookup, replacing a linear scan on every switch
        self.view_defs = dict(self.tab_defs)

        # Stack page 0 is the empty state, shown until a file is loaded
        self.stack.addWidget(self._build_empty_state())

        # Views are placed by NAV_GROUPS. Anything registered in tab_defs but
        # not placed there still appears, under a trailing MORE group, so
        # "append to tab_defs" remains sufficient to add a view.
        placed = {n for _, names in NAV_GROUPS for n in names} | NON_NAV_VIEWS
        unplaced = [n for n, _ in self.tab_defs if n not in placed]
        groups = list(NAV_GROUPS) + ([("MORE", unplaced)] if unplaced else [])

        # Group captions and rows are both painted by NavDelegate, which is
        # what lets a row carry an index and a badge without either becoming
        # part of its text. The caption is still flagged unselectable here, so
        # keyboard navigation skips it.
        self._first_view_row = None
        position = 0
        for group_index, (group_name, names) in enumerate(groups):
            header = QListWidgetItem(group_name)
            header.setFlags(Qt.ItemFlag.NoItemFlags)
            # Extra height above every group except the first, to separate them
            header.setSizeHint(QSize(
                0, NAV_GROUP_HEIGHT if group_index else NAV_GROUP_FIRST_HEIGHT))
            self.nav_list.addItem(header)

            for name in names:
                if name not in self.view_defs:
                    continue
                page = QWidget()
                self.tab_widgets[name] = page
                self.view_indices[name] = self.stack.addWidget(page)

                position += 1
                item = QListWidgetItem(name)
                item.setData(Qt.ItemDataRole.UserRole, name)
                item.setData(NAV_ROLE_INDEX, f"{position:02d}")
                item.setSizeHint(QSize(0, SIDEBAR_ROW_HEIGHT))
                self.nav_list.addItem(item)
                if self._first_view_row is None:
                    self._first_view_row = self.nav_list.row(item)

        # Nothing is selected until data arrives; the empty state is showing.
        self._set_content_enabled(False)

    def _build_empty_state(self) -> QWidget:
        """Guidance panel shown before any file is loaded (stack page 0)."""
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.addStretch()

        row = QHBoxLayout()
        row.addStretch()
        card, card_layout = make_card(margin=SPACE_LG * 2)
        card.setMaximumWidth(520)

        title = QLabel("No data loaded")
        title.setObjectName("SectionHeader")
        card_layout.addWidget(title)

        steps = QLabel(
            "1.  Choose an Excel or CSV file\n"
            "2.  Pick a sheet, if the workbook has more than one\n"
            "3.  Select a check from the sidebar\n\n"
            "Supports .xlsx, .xls, .csv, .tsv and .txt."
        )
        steps.setWordWrap(True)
        set_state(steps, "muted")
        card_layout.addWidget(steps)

        browse = QPushButton("Browse for a file")
        browse.setObjectName("PrimaryButton")
        browse.setIcon(icon("folder", COLOR_TEXT_STRONG))
        browse.clicked.connect(self.browse_file)
        card_layout.addWidget(browse)

        row.addWidget(card)
        row.addStretch()
        outer.addLayout(row)
        outer.addStretch()
        return page

    def _set_content_enabled(self, enabled: bool) -> None:
        """
        Gate the views on having data.

        Deliberately does not disable self.stack: that would grey out the empty
        state's own Browse button and leave the user with no way forward.
        """
        self.nav_list.setEnabled(enabled)
        for page in self.tab_widgets.values():
            page.setEnabled(enabled)
        can_export = enabled and self.df is not None
        self.btn_export.setEnabled(can_export)
        # The menu action was previously always enabled, so Ctrl+E worked with
        # no data loaded.
        self.export_action.setEnabled(can_export)

    def set_dataset_badge(self) -> None:
        """
        Say what is loaded, in the three places it is needed.

        The sidebar block is where you look before starting; the top-bar
        breadcrumb is where you look while reading a table on a maximised
        window; the status bar is the one surface for transient state and keeps
        the same facts beside it. All three come from here so they cannot drift.
        """
        if self.df is None:
            self.dataset_badge.setText("No data loaded")
            self.dataset_meta.setText("Open a file to begin")
            self.crumb_file.setText("No data")
            self.crumb_rest.setText("")
            return
        name = os.path.basename(self.filepath) if self.filepath else "data"
        sheet = self.combo_sheets.currentText()
        has_sheet = bool(sheet) and not self.combo_sheets.isHidden()
        shape = f"{len(self.df):,} rows  ·  {len(self.df.columns)} fields"

        self._set_elided(self.file_label, name)
        self.dataset_meta.setText(f"{sheet}  ·  {shape}" if has_sheet else shape)
        self.crumb_file.setText(name)
        self._update_breadcrumb()
        where = f"{name} › {sheet}" if has_sheet else name
        self.dataset_badge.setText(
            f"{where}  ·  {len(self.df):,} rows × {len(self.df.columns)} cols"
        )

    @staticmethod
    def _set_elided(label, text: str) -> None:
        """
        Put `text` on a one-line label, shortened to fit, full value as tooltip.

        Elided against the label's current width, so dragging the splitter
        leaves it stale until the next write. That is deliberate rather than an
        event filter: set_dataset_badge is the only writer, a sheet change
        re-runs it, and the untruncated name is on the tooltip, the breadcrumb
        and the status bar regardless.
        """
        width = label.width()
        label.setToolTip(text)
        if width <= 0:
            label.setText(text)
            return
        label.setText(label.fontMetrics().elidedText(
            text, Qt.TextElideMode.ElideMiddle, width))

    def _update_breadcrumb(self) -> None:
        """The trail after the file name: sheet, then the view being read."""
        parts = []
        sheet = self.combo_sheets.currentText()
        if sheet and not self.combo_sheets.isHidden():
            parts.append(sheet)
        if self.current_view:
            parts.append(self.current_view)
        self.crumb_rest.setText("  /  ".join([""] + parts) if parts else "")

    def table_copy_context(self, view) -> List[str]:
        """
        Provenance header prepended when a table is copied for a report.

        Called by the table copy helper via view.window(), so every table gets
        it without individual wiring. A pasted table of counts is unusable in a
        report without the file and sheet it came from - and a percentage is
        unreadable without its denominator, which is why the row counts are
        here rather than left implicit in the column header.
        """
        title = getattr(view, "_copy_title", None) or self.current_view or "Results"
        lines = [f"TicketAudit — {title}"]

        if self.filepath:
            lines.append(f"Dataset: {os.path.basename(self.filepath)}")
        sheet = self.combo_sheets.currentText()
        if sheet and not self.combo_sheets.isHidden():
            lines.append(f"Sheet: {sheet}")
        if self.df is not None:
            lines.append(f"Rows in sheet: {len(self.df):,}")

        for label, value in getattr(view, "_copy_extra", {}).items():
            lines.append(f"{label}: {value}")

        lines.append(f"Generated: {datetime.datetime.now():%Y-%m-%d %H:%M}")
        return lines

    # --- Worker Thread Methods ---
    
    def run_in_background(self, func, callback=None, err_callback=None,
                          report_progress=False, *args, **kwargs):
        """
        Run a function in the background thread pool.

        report_progress hands the task a `progress` callable. WorkerSignals has
        always had a progress signal wired to update_status, but nothing could
        ever emit it because the task was never given the emitter. What the task
        receives is `signal.emit`, which belongs to a QObject rather than a
        widget, and Qt delivers it to the main thread as a queued event - so the
        worker still never touches a widget.
        """
        log.debug("starting background task %s", func.__name__)
        worker = Worker(func, *args, **kwargs)
        if report_progress:
            worker.kwargs["progress"] = worker.signals.progress.emit
        
        # Keep reference
        self._active_workers.append(worker)
        
        # Cleanup when finished
        def cleanup():
            if worker in self._active_workers:
                self._active_workers.remove(worker)
        
        worker.signals.finished.connect(cleanup)
        
        if callback:
            worker.signals.result.connect(callback)
        if err_callback:
            worker.signals.error.connect(err_callback)
        else:
            def default_error_handler(e):
                log.warning("background task failed: %s", e)
                self.show_error(f"Error: {e}")
            worker.signals.error.connect(default_error_handler)
            
        worker.signals.progress.connect(self.update_status)
        self.threadpool.start(worker)

    def _refresh_async(self, key, task, populate, busy_msg=None, err_callback=None):
        """
        Run an analysis off the UI thread and hand the result to `populate`.

        `task` is called with the analyzer as its only argument. That is not
        cosmetic: a closure reading `self.analyzer` resolves it when the worker
        *runs*, so a sheet switch mid-flight silently changes which data is
        analysed. Binding it here pins the task to the analyzer it was queued
        for, and the sequence guard then discards its result.

        Three guards that a bare run_in_background does not provide:

        * **Sequence** - `_data_seq` is bumped on every sheet load. A result
          computed for an earlier sheet is dropped instead of being painted
          over the current one. Without this, a slow refresh that finishes
          after the user has switched sheets wins, and the view shows the
          previous sheet's numbers under the current sheet's name.
        * **Lock** - `SanityAnalyzer._cache` is lazily populated and shared, so
          two concurrent refreshes can duplicate work on it.
        * **In-flight** - a view already computing is not queued again, so
          repeatedly revisiting it cannot pile up workers.
        """
        if self.df is None or self.analyzer is None:
            return
        if key in self._refresh_inflight:
            return

        seq = self._data_seq
        analyzer = self.analyzer
        # Bind the lock too, not just the analyzer: reading self._analyzer_lock
        # inside the worker would pick up the *new* sheet's lock and defeat the
        # point of replacing it.
        lock = self._analyzer_lock
        self._refresh_inflight.add(key)
        if busy_msg:
            self.update_status(busy_msg)

        def guarded_task():
            with lock:
                return task(analyzer)

        def on_result(result):
            self._refresh_inflight.discard(key)
            if seq != self._data_seq:
                log.debug("discarding stale %s result (seq %s != %s)",
                          key, seq, self._data_seq)
                return
            populate(result)
            # A populate can create labels (findings cards, summaries), and
            # those land after show_view has already walked the page.
            if self.current_view in self.tab_widgets:
                make_text_selectable(self.tab_widgets[self.current_view])

        def on_error(err):
            self._refresh_inflight.discard(key)
            if seq != self._data_seq:
                return
            if err_callback:
                err_callback(err)
            else:
                log.warning("%s refresh failed: %s", key, err)
                self.show_error(f"Failed to refresh {key}: {err}")

        self.run_in_background(guarded_task, on_result, on_error)

    def update_status(self, message):
        """Single surface for transient status. Also the worker progress slot."""
        self.status_bar.showMessage(message, 5000)

    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)

    # --- File Operations ---

    def browse_file(self):
        start_dir = self.config.get("last_folder", "")
        if not os.path.isdir(start_dir):
            start_dir = ""
        filename, _ = QFileDialog.getOpenFileName(
            self, "Open Data File", start_dir,
            "Data Files (*.xlsx *.xls *.csv *.tsv *.txt);;"
            "Excel Files (*.xlsx *.xls);;"
            "CSV / Text Files (*.csv *.tsv *.txt);;"
            "All Files (*)"
        )
        if filename:
            self.config["last_folder"] = os.path.dirname(filename)
            self.filepath = filename
            self.file_label.setText(os.path.basename(filename))
            self.load_file_async(filename)

    def load_file_async(self, filename):
        self.update_status("Loading file...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate
        self.btn_load.setEnabled(False)
        self.combo_sheets.setEnabled(False)
        
        def load_task():
            sheets = loader.list_sheets(filename)
            log.debug("%s: sheets=%s", os.path.basename(filename), sheets)
            return sheets

        def on_loaded(sheet_names):
            self.progress_bar.setVisible(False)
            self.btn_load.setEnabled(True)
            # Block signals while repopulating: addItems() selects index 0 and emits
            # currentTextChanged, which would start a second concurrent load_sheet
            # worker racing the explicit call below for self.df / self.analyzer.
            self.combo_sheets.blockSignals(True)
            self.combo_sheets.clear()
            self.combo_sheets.addItems(sheet_names)
            self.combo_sheets.blockSignals(False)

            # Delimited text has no sheets - hide the picker entirely rather
            # than showing an empty dropdown.
            has_sheets = bool(sheet_names)
            self.combo_sheets.setVisible(has_sheets)
            self.combo_sheets.setEnabled(has_sheets)
            self.load_sheet(sheet_names[0] if has_sheets else None)
        
        def on_error(err):
            log.warning("failed to list sheets in %s: %s", filename, err)
            self.progress_bar.setVisible(False)
            self.btn_load.setEnabled(True)
            self.show_error(f"Failed to load file: {err}")

        self.run_in_background(load_task, on_loaded, on_error)

    def load_sheet(self, sheet_name=None):
        """
        Load an Excel sheet, or the single table of a delimited-text file.

        sheet_name is None for CSV/TSV (no sheets). An empty string means the
        sheet combo was cleared and is not a real load request.
        """
        log.debug("load_sheet(%r) from %s", sheet_name, self.filepath)
        if not self.filepath or sheet_name == "":
            return

        display = sheet_name or os.path.basename(self.filepath)
        self.update_status(f"Loading {display}...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self._set_content_enabled(False)

        # The saved mapping and date order must be applied here too, or loading
        # a sheet silently reverts to the automatic guesses the user corrected.
        overrides = dict(self.config.get("column_overrides") or {})
        dayfirst = self.config.get("date_dayfirst")

        def load_task():
            df = loader.read_table(self.filepath, sheet_name)
            return df, SanityAnalyzer(df, dayfirst=dayfirst,
                                      column_overrides=overrides)

        def on_loaded(result):
            self.df, self.analyzer = result
            # Invalidate anything still computing against the previous sheet,
            # so its result is discarded rather than painted over this one.
            # The fresh lock lets this sheet's refreshes start immediately
            # instead of queueing behind the old sheet's unfinished work.
            self._data_seq += 1
            self._refresh_inflight.clear()
            self._analyzer_lock = threading.Lock()
            # Findings belong to the sheet they were computed on; carrying them
            # over would put the previous sheet's languages in this one's report.
            self.language_results.clear()
            # A hand-made column selection belongs to the sheet it was made on;
            # the next one starts from the remembered set again.
            self._extract_ticks = None
            self._extract_month_ticks = None
            self._extract_fill = None
            self._extract_cancelled_info = None
            self._extract_last_path = None
            # Having run the register on one sheet says nothing about the next,
            # and re-running it automatically here would put back the whole cost
            # of a file open that gating it removed.
            self._overview_started = False
            # Same for the report: a document written for the previous sheet is
            # not a document about this one, so it is asked for again.
            self._report_started = False
            self.progress_bar.setVisible(False)
            self._set_content_enabled(True)
            self.update_status(f"Loaded {len(self.df):,} rows.")
            # A badge counts findings in the sheet it was computed on; carrying
            # one over would put the previous sheet's numbers beside this one.
            self.clear_nav_badges()
            self.set_dataset_badge()

            # First successful load selects a view; later loads refresh in place.
            # setCurrentRow on an unchanged row does not re-emit, hence the else.
            if self.current_view is None and self._first_view_row is not None:
                self.nav_list.setCurrentRow(self._first_view_row)
            else:
                self.refresh_current_tab()

        def on_error(err):
            # Without this the default handler only shows a dialog, leaving the
            # progress bar spinning and the content disabled until restart.
            log.warning("failed to load %s: %s", display, err)
            self.progress_bar.setVisible(False)
            # Re-enables only if earlier data is still loaded, so a first-load
            # failure correctly leaves the empty state showing.
            self._set_content_enabled(self.df is not None)
            self.update_status("Load failed.")
            self.show_error(f"Failed to load {display}: {err}")

        self.run_in_background(load_task, on_loaded, on_error)

    # --- View switching (lazy construction) ---

    def on_nav_changed(self, current, previous):
        """Sidebar selection changed. Group headers carry no view name."""
        name = current.data(Qt.ItemDataRole.UserRole) if current is not None else None
        if name:
            self.show_view(name)

    def show_view(self, name: str) -> None:
        """
        Display a view, building it on first use.

        The build-then-refresh order matters: every refresh_* method assumes the
        widgets its create_*_tab made already exist.

        The guard is on `tab_widgets`, not `view_defs`. A NON_NAV_VIEWS entry
        such as Settings is registered in `view_defs` but is placed in no
        NAV_GROUPS group, so `setup_views` never builds it a stack page - it is
        reached as a dialog instead. Guarding on `view_defs` let such a name
        past here and then raised KeyError on the `tab_widgets` lookup below,
        but only once a file was open, because the empty-state return above
        fires first.
        """
        if name not in self.tab_widgets:
            return

        self.current_view = name
        self.view_title.setText(name)
        self.view_subtitle.setText(VIEW_SUBTITLES.get(name, ""))
        self.view_subtitle.setVisible(bool(VIEW_SUBTITLES.get(name)))
        self._update_view_header()
        self._update_breadcrumb()
        self._update_review_ribbon()

        if self.df is None:
            self.stack.setCurrentIndex(0)   # empty state
            return

        page = self.tab_widgets[name]
        if page.layout() is None:
            self.view_defs[name](page)

        self.stack.setCurrentIndex(self.view_indices[name])
        self.refresh_tab(name)
        # After the refresh, so labels a populate step creates are covered too.
        # Idempotent per label, and one wiring point rather than 70.
        make_text_selectable(page)

    def refresh_current_tab(self):
        """Kept under its original name - existing call sites rely on it."""
        if self.current_view:
            self.show_view(self.current_view)

    # --- Window state persistence ---

    def _restore_window_state(self):
        """
        Restore size, position, sidebar width and last view from config.

        Stored as plain ints rather than a saveGeometry() blob so config.json
        stays hand-editable.
        """
        try:
            geom = self.config.get("window_geometry") or {}
            if all(k in geom for k in ("x", "y", "w", "h")):
                rect = QRect(int(geom["x"]), int(geom["y"]),
                             int(geom["w"]), int(geom["h"]))
                # A saved position on a monitor that is no longer attached would
                # open the window off-screen with no way to reach it.
                if any(s.availableGeometry().intersects(rect)
                       for s in QApplication.screens()):
                    self.setGeometry(rect)
                if geom.get("maximized"):
                    self.showMaximized()

            width = int(self.config.get("sidebar_width", SIDEBAR_WIDTH))
            width = max(200, min(width, SIDEBAR_MAX_WIDTH))
            total = max(self.width(), width + 400)
            self.splitter.setSizes([width, total - width])

            # Last view, restored by selecting its nav row rather than calling
            # show_view, so the sidebar highlight, the page header and
            # current_view cannot disagree. With no file open that lands on the
            # empty state, which is correct - what this buys is the *first*
            # load: load_sheet only jumps to the first row while current_view
            # is None, and refreshes in place once it is set, so the session
            # resumes on the view it ended on.
            #
            # Guarded on tab_widgets, so a view that has been renamed or
            # removed since the config was written - or a NON_NAV_VIEWS name
            # like Settings, which has no page - falls through to the default
            # instead of leaving the window on a blank selection.
            saved_view = self.config.get("last_view") or ""
            if saved_view in self.tab_widgets:
                for row in range(self.nav_list.count()):
                    item = self.nav_list.item(row)
                    if item.data(Qt.ItemDataRole.UserRole) == saved_view:
                        self.nav_list.setCurrentRow(row)
                        break
        except Exception as exc:
            log.warning("could not restore window state: %s", exc)

    def closeEvent(self, event):
        """Persist window state. A config write must never block exit."""
        try:
            rect = self.normalGeometry() if self.isMaximized() else self.geometry()
            self.config["window_geometry"] = {
                "x": rect.x(), "y": rect.y(),
                "w": rect.width(), "h": rect.height(),
                "maximized": self.isMaximized(),
            }
            sizes = self.splitter.sizes()
            if sizes and sizes[0] > 0:
                self.config["sidebar_width"] = sizes[0]
            if self.current_view:
                self.config["last_view"] = self.current_view
            self.config_mgr.config = self.config
            self.config_mgr.save()
        except Exception as exc:
            log.warning("could not save window state: %s", exc)
        super().closeEvent(event)

    def show_settings_dialog(self):
        """
        Settings as a modeless dialog rather than a peer of the analysis views.

        The dialog is cached on the instance and never destroyed.
        """
        if self._settings_dialog is None:
            dialog = QDialog(self)
            dialog.setWindowTitle("Settings")
            dialog.setMinimumWidth(620)
            layout = QVBoxLayout(dialog)
            layout.setContentsMargins(SPACE_MD, SPACE_MD, SPACE_MD, SPACE_MD)

            body = QWidget()
            self.create_settings_tab(body)   # unchanged creator
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.Shape.NoFrame)
            scroll.setWidget(body)
            layout.addWidget(scroll)

            close_row = QHBoxLayout()
            close_row.addStretch()
            close_btn = QPushButton("Close")
            close_btn.clicked.connect(dialog.hide)
            close_row.addWidget(close_btn)
            layout.addLayout(close_row)

            # Model paths and load errors live here, and both are things a user
            # needs to paste somewhere.
            make_text_selectable(dialog)
            self._settings_dialog = dialog

        self._settings_dialog.show()
        self._settings_dialog.raise_()
        self._settings_dialog.activateWindow()

    def refresh_tab(self, tab_name):
        """Dispatch refresh to specific tab methods."""
        if self.df is None or self.analyzer is None:
            return

        refresh_map = {
            "Overview": self.refresh_overview,
            "Column Check": self.refresh_column_check,
            "Null Analysis": self.refresh_null_analysis,
            "Logic Checks": self.refresh_logic_checks,
            "Language Check": self.refresh_language_check,
            "Description Quality": self.refresh_description_quality,
            "Monthly Inflow": self.refresh_monthly_inflow,
            "Pivot Tables": self.refresh_pivot_tables,
            "Data Profile": self.refresh_data_profile,
            "Duplicate Check": self.refresh_duplicate_check,
            "Raw Data": self.refresh_raw_data,
            "Extract": self.refresh_extract,
            "Report": self.refresh_report
        }
        
        if tab_name in refresh_map:
            refresh_map[tab_name]()

    # --- Tab Implementations ---

    # 0. Overview - the register, and what was never looked at
    #
    # Which view an Area belongs to. The register is only useful if a row leads
    # somewhere, and Area is the column reporter.py already sorts by, so this is
    # the whole drill-through map. Date-format findings are filed under Structure
    # by the reporter but are corrected in Logic Checks, which is why Check is
    # consulted before Area.
    OVERVIEW_AREA_VIEWS = {
        "Structure": "Column Check",
        "Completeness": "Null Analysis",
        "Logic": "Logic Checks",
        "Uniqueness": "Duplicate Check",
        "Text quality": "Description Quality",
        "Language": "Language Check",
    }
    OVERVIEW_TABS = ["All", "Failing", "Review", "Not checked", "Passed"]
    OVERVIEW_COLUMNS = ["Result", "Severity", "Check", "Area", "Rows Affected",
                        "Detail"]

    def create_overview_tab(self, parent):
        """
        The one view that answers "is this file usable?".

        Deliberately not scrollable: the register table takes the leftover
        height, which is the whole point of putting it here rather than in a
        card sized to its content. Nothing in this view pins a height, so the
        rule that a pinning view must scroll does not apply.
        """
        layout = make_view(parent)

        tiles = QHBoxLayout()
        tiles.setSpacing(SPACE_SM)
        self.overview_kpis = {}
        for label in ("Failing", "Needs review", "Not checked", "Passing",
                      "Rows flagged"):
            frame, value = make_kpi(tiles, label)
            self.overview_kpis[label] = (frame, value)
        layout.addLayout(tiles)

        body = QHBoxLayout()
        body.setSpacing(SPACE_SM)

        register, register_layout = make_card(margin=SPACE_SM)
        bar = QHBoxLayout()
        bar.setSpacing(SPACE_XS)
        title = QLabel("FINDINGS REGISTER")
        title.setObjectName("CardTitle")
        bar.addWidget(title)
        bar.addSpacing(SPACE_MD)
        self.overview_tabs = make_seg_tabs(bar, self.OVERVIEW_TABS,
                                           self._show_overview_tab)
        bar.addStretch()
        self.btn_overview_open = QPushButton("Open this check")
        self.btn_overview_open.setObjectName("GhostButton")
        self.btn_overview_open.setToolTip(
            "Go to the view that produced the selected row")
        self.btn_overview_open.setEnabled(False)
        self.btn_overview_open.clicked.connect(self._open_selected_finding)
        bar.addWidget(self.btn_overview_open)
        register_layout.addLayout(bar)

        self.overview_note = QLabel("")
        self.overview_note.setWordWrap(True)
        set_state(self.overview_note, "muted")
        register_layout.addWidget(self.overview_note)

        # The register runs every check in the app - ~1.2s at 100k rows and
        # ~3.7s at 300k, almost all of it check_description_quality - so opening
        # a file does not start it. Until the user asks, this card is what the
        # panel holds; it states what has not happened rather than showing
        # zeros, which read as a clean file.
        self.overview_prompt, prompt_layout = make_card(margin=SPACE_MD)
        prompt_title = QLabel("Nothing has been checked yet")
        prompt_title.setObjectName("CardTitle")
        prompt_layout.addWidget(prompt_title)
        self.overview_prompt_body = QLabel()
        self.overview_prompt_body.setWordWrap(True)
        prompt_layout.addWidget(self.overview_prompt_body)
        prompt_actions = QHBoxLayout()
        prompt_actions.setSpacing(SPACE_SM)
        self.btn_overview_run = QPushButton("Run all checks")
        self.btn_overview_run.setObjectName("PrimaryButton")
        self.btn_overview_run.setToolTip(
            "Run every check on this sheet and build the findings register")
        self.btn_overview_run.clicked.connect(self._start_overview)
        prompt_actions.addWidget(self.btn_overview_run)
        prompt_actions.addStretch()
        prompt_layout.addLayout(prompt_actions)
        self.overview_prompt_hint = QLabel(
            "Every other view works without this. The register is what puts all "
            "of their findings in one place, worst first.")
        self.overview_prompt_hint.setObjectName("Caption")
        self.overview_prompt_hint.setWordWrap(True)
        prompt_layout.addWidget(self.overview_prompt_hint)
        register_layout.addWidget(self.overview_prompt)

        self.overview_table = QTableView()
        configure_table(self.overview_table, title="Findings register",
                        stretch_last=True)
        self.overview_table.doubleClicked.connect(
            lambda _index: self._open_selected_finding())
        register_layout.addWidget(self.overview_table, 1)

        # Collects the leftover height while the table is hidden. Without it a
        # QVBoxLayout hands that space to whichever visible items will take it -
        # a QLabel's vertical policy is Preferred, not Fixed - so the card title
        # drifted to the middle of the panel and the prompt sat at the bottom
        # with 115px of white above it.
        #
        # A spacer whose stretch is toggled, not a filler widget: a plain
        # QWidget is matched by the `QMainWindow, QWidget` QSS rule and painted
        # in the page colour, which put a grey block under the prompt. What is
        # wanted here is space, and a spacer cannot be painted at all. The
        # factor has to be toggled rather than left at 1, or it would compete
        # with the table's own stretch and halve the register once it shows.
        self._overview_register_layout = register_layout
        register_layout.addStretch(0)
        self._overview_filler_index = register_layout.count() - 1
        body.addWidget(register, 3)

        rail = QVBoxLayout()
        rail.setSpacing(SPACE_SM)

        gaps, gaps_layout = make_card(rail, margin=SPACE_MD)
        gaps_title = QLabel("NOT CHECKED")
        gaps_title.setObjectName("CardTitle")
        gaps_layout.addWidget(gaps_title)
        # Named rather than counted: "3 not checked" tells a reader they have a
        # gap without telling them which one, and the fix differs per check.
        self.overview_gaps = QLabel("Working...")
        self.overview_gaps.setWordWrap(True)
        gaps_layout.addWidget(self.overview_gaps)

        facts, facts_layout = make_card(rail, margin=SPACE_MD)
        facts_title = QLabel("THIS FILE")
        facts_title.setObjectName("CardTitle")
        facts_layout.addWidget(facts_title)
        self.overview_facts = QLabel("")
        self.overview_facts.setWordWrap(True)
        facts_layout.addWidget(self.overview_facts)

        # The area breakdown takes the rail's leftover height instead of a
        # stretch, which is what stops the column below "this file" being void.
        # It is also the fastest way to see which part of the file is the
        # problem, which is the question the register's ordering cannot answer.
        areas, areas_layout = make_card(margin=SPACE_MD)
        areas_title = QLabel("BY AREA")
        areas_title.setObjectName("CardTitle")
        areas_layout.addWidget(areas_title)
        # A table of headers with no rows reads as a rendering fault, and it
        # even draws its own horizontal scrollbar. While nothing has run the
        # card says so instead - the same choice Duplicate Check makes.
        self.overview_area_note = QLabel(
            "One row per area once the checks have run.")
        self.overview_area_note.setWordWrap(True)
        set_state(self.overview_area_note, "muted")
        areas_layout.addWidget(self.overview_area_note)
        self.overview_area_table = QTableView()
        configure_table(self.overview_area_table, title="Findings by area",
                        stretch_last=True)
        areas_layout.addWidget(self.overview_area_table, 1)
        areas_layout.addStretch()
        rail.addWidget(areas, 1)

        rail_host = QWidget()
        rail_host.setLayout(rail)
        # A minimum as well as a maximum: with only a stretch factor the rail
        # collapsed to about 120px on an 1100px window and every line in it
        # wrapped to one or two words.
        rail_host.setMinimumWidth(272)
        rail_host.setMaximumWidth(340)
        body.addWidget(rail_host, 1)

        layout.addLayout(body, 1)
        self._overview_frame = None
        self._overview_tab = "All"
        # Paints the prompt rather than leaving the tiles blank on first build.
        # refresh_overview runs straight after this, but a view that is created
        # and not yet refreshed is briefly on screen.
        self._show_overview_idle()

    def refresh_overview(self):
        """
        Build the register off the UI thread, once the user has asked for it.

        _build_findings_df runs every check in the app - ~1.2s at 100k rows and
        ~3.7s at 300k - so this can never be synchronous, and it must say it is
        working rather than showing zeros that look like a clean file.

        It also does not start on its own. Landing on this view is what opening
        a file does, so an automatic run made every file open pay for the whole
        register whether or not anyone was going to read it, which is what made
        loading feel slow. `_overview_started` is reset per sheet in load_sheet:
        having run the checks on one sheet says nothing about the next.
        """
        if not self._overview_started:
            self._show_overview_idle()
            return

        # Snapshotted before dispatch, for the reason _refresh_async binds the
        # analyzer: a closure over self.df resolves when the worker *runs*, so a
        # sheet switch mid-flight would build the register from other data.
        frame = self.df
        filepath = self.filepath
        lang_results = list(self.language_results.values())
        threshold = float(self.config.get("null_threshold", 20))
        desc_min = int(self.config.get("desc_min_length", 20))
        desc_max = int(self.config.get("desc_max_length", 5000))

        self._set_overview_running(True)
        for _label, (tile, value) in self.overview_kpis.items():
            value.setText("...")
            set_state(value, "busy")
            set_severity(tile, None)
        self.overview_gaps.setText("Working...")
        set_state(self.overview_gaps, "busy")

        def task(analyzer):
            reporter = ReportGenerator(
                frame, analyzer, filepath, null_threshold=threshold,
                desc_min_length=desc_min, desc_max_length=desc_max,
                lang_results=lang_results)
            register = reporter._build_findings_df()
            flags = analyzer.get_row_flags()
            masks = list(flags.get('flags', {}).values())
            flagged = int(np.logical_or.reduce([m.to_numpy() for m in masks]).sum()) \
                if masks else 0
            return {'register': register, 'flagged': flagged,
                    'rows': len(frame), 'unchecked_rules': flags.get('skipped', [])}

        self._refresh_async("Overview", task, self._populate_overview,
                            busy_msg="Collecting every check...",
                            err_callback=self._overview_failed)

    def _start_overview(self) -> None:
        """The button under the prompt. Runs the register for this sheet."""
        if self.df is None or self.analyzer is None:
            return
        self._overview_started = True
        self.refresh_overview()

    def _overview_failed(self, err) -> None:
        """
        A failed register returns to the prompt rather than to empty tiles.

        Without this the view keeps the busy "..." for ever and the button that
        would retry is hidden behind it, so the only way out is a sheet reload.
        """
        log.warning("could not build the findings register: %s", err)
        self._overview_started = False
        self._show_overview_idle()
        self.overview_prompt_body.setText(
            f"The checks could not be completed on this sheet: {err}\n\n"
            f"Every other view still works. Press the button to try again.")
        set_state(self.overview_prompt_body, "error")
        self.btn_overview_run.setText("Try again")
        self.update_status("The findings register could not be built.")

    def _set_overview_running(self, running: bool) -> None:
        """
        Swap the panel between the prompt and the register.

        The table is hidden rather than left empty for the reason Duplicate
        Check hides its splitter: a window of blank rows is a worse answer than
        one sentence saying what has not happened. The tab strip goes with it -
        filtering a register that does not exist yet is a control that cannot do
        anything.
        """
        self.overview_prompt.setVisible(not running)
        self._overview_register_layout.setStretch(
            self._overview_filler_index, 0 if running else 1)
        self.overview_table.setVisible(running)
        self.overview_note.setVisible(running)
        self.btn_overview_open.setVisible(running)
        self.overview_area_table.setVisible(running)
        self.overview_area_note.setVisible(not running)
        for button in self.overview_tabs.values():
            button.setVisible(running)

    def _show_overview_idle(self) -> None:
        """
        The view before anyone has asked for the checks.

        Every tile reads as absent, never as zero: a row of zeros here would
        say this file has no failures, which is a claim about data nothing has
        looked at. The facts card is still filled in, because rows, fields and
        the column map are already known from loading and cost nothing.
        """
        self._overview_frame = None
        self._set_overview_running(False)
        self.btn_overview_run.setText("Run all checks")
        set_state(self.overview_prompt_body, "muted")

        rows = len(self.df) if self.df is not None else 0
        self.overview_prompt_body.setText(
            f"This sheet has {rows:,} rows. Running the register works through "
            f"every check in the app and can take a few seconds on a file this "
            f"size, so it waits until you ask.\n\n"
            f"It collects what all the other views found into one table, worst "
            f"first, and names the checks that could not run at all.")

        for _label, (tile, value) in self.overview_kpis.items():
            # An em dash, not 0. The tile has no number to show, and a zero on
            # "Failing" is the single most misleading thing this view could say.
            value.setText("—")
            set_state(value, "muted")
            set_severity(tile, None)

        self.overview_gaps.setText(
            "Not run yet. Nothing on this sheet has been checked, which is not "
            "the same as nothing being wrong with it.")
        set_state(self.overview_gaps, "muted")
        if self.analyzer is not None:
            self._populate_overview_facts(rows)
        # A badge counts failing checks, so there can be none while no check has
        # run. Leaving the previous sheet's counts here was the bug that
        # clear_nav_badges in load_sheet exists to prevent.
        self.clear_nav_badges()

    def _populate_overview(self, result):
        register = result['register']
        self._overview_frame = register
        self._set_overview_running(True)

        counts = {name: 0 for name in RESULT_STATES}
        if not register.empty:
            counts.update(register['Result'].value_counts().to_dict())

        tiles = {
            "Failing": (counts.get("Fail", 0), "error"),
            "Needs review": (counts.get("Review", 0), "warn"),
            "Not checked": (counts.get("Not checked", 0), "warn"),
            "Passing": (counts.get("Pass", 0), "ok"),
            "Rows flagged": (result['flagged'], "error"),
        }
        for label, (count, severity) in tiles.items():
            frame, value = self.overview_kpis[label]
            value.setText(f"{count:,}")
            # A zero is the good outcome for three of these, so it is left in
            # the body colour rather than tinted green - a green 0 beside a red
            # 0 in the next tile reads as two different kinds of nothing.
            set_state(value, severity if count else None)
            set_severity(frame, severity if count and severity != "ok" else None)

        self._show_overview_tab(self._overview_tab)
        self._populate_overview_areas(register)
        self._populate_overview_gaps(register, result['unchecked_rules'])
        self._populate_overview_facts(result['rows'])
        self._apply_overview_badges(register)

        set_copy_context(
            self.overview_table,
            **{"Checks in register": f"{len(register):,}",
               "Failing": f"{counts.get('Fail', 0):,}",
               "Not checked": f"{counts.get('Not checked', 0):,}"})
        # Replaces the busy message, which otherwise sits in the status bar
        # claiming to still be collecting long after the table has filled.
        self.update_status(
            f"{len(register):,} checks: {counts.get('Fail', 0):,} failing, "
            f"{counts.get('Not checked', 0):,} not checked.")

    def _show_overview_tab(self, name: str) -> None:
        """Filter the register in place. Never rebuilds it - it costs seconds."""
        self._overview_tab = name
        register = self._overview_frame
        if register is None:
            return
        wanted = {"Failing": "Fail", "Review": "Review",
                  "Not checked": "Not checked", "Passed": "Pass"}.get(name)
        shown = register if wanted is None else \
            register[register['Result'] == wanted]
        columns = [c for c in self.OVERVIEW_COLUMNS if c in shown.columns]
        set_table_model(self.overview_table,
                        FindingsModel(shown[columns].reset_index(drop=True)))
        selection = self.overview_table.selectionModel()
        if selection is not None:
            selection.selectionChanged.connect(self._on_finding_selected)
        self.btn_overview_open.setEnabled(False)

        if shown.empty:
            self.overview_note.setText(
                f"No check has the outcome \u201c{name}\u201d on this sheet.")
        elif wanted is None:
            self.overview_note.setText(
                f"{len(shown):,} checks, worst first: failures, then anything "
                "that was never looked at, then judgements, then passes. "
                "Double-click a row to open the view behind it.")
        else:
            self.overview_note.setText(
                f"{len(shown):,} of {len(register):,} checks. Double-click a "
                "row to open the view behind it.")

    def _on_finding_selected(self, *_args) -> None:
        self.btn_overview_open.setEnabled(
            bool(self.overview_table.selectionModel()
                 and self.overview_table.selectionModel().selectedRows()))

    def _open_selected_finding(self) -> None:
        """Jump to the view that produced the selected row."""
        selection = self.overview_table.selectionModel()
        rows = selection.selectedRows() if selection else []
        if not rows:
            return
        model = self.overview_table.model()
        frame = model._df
        row = frame.iloc[rows[0].row()]
        self._select_view(self._view_for_finding(row.get("Area", ""),
                                                str(row.get("Check", ""))))

    def _view_for_finding(self, area: str, check: str) -> str:
        if check.startswith("Date format of"):
            return "Logic Checks"
        return self.OVERVIEW_AREA_VIEWS.get(area, "Column Check")

    def _populate_overview_areas(self, register) -> None:
        """
        One row per area: how many of its checks failed, were skipped, passed.

        Areas come from the register rather than a list here, so an area the
        reporter adds shows up without a second place to edit. "Not checked" is
        its own column and never folded into a total, for the same reason it is
        its own outcome everywhere else.
        """
        if register.empty:
            set_table_model(self.overview_area_table,
                            PandasModel(pd.DataFrame(
                                columns=["Area", "Fail", "Not checked", "Pass"])))
            return
        pivot = (register.pivot_table(index="Area", columns="Result",
                                      values="Check", aggfunc="count")
                 .fillna(0).astype(int))
        for column in ("Fail", "Not checked", "Review", "Pass"):
            if column not in pivot.columns:
                pivot[column] = 0
        table = pivot[["Fail", "Not checked", "Review", "Pass"]].reset_index()
        table = table.sort_values(["Fail", "Not checked", "Review"],
                                  ascending=False)
        set_table_model(self.overview_area_table,
                        FindingsModel(table.reset_index(drop=True),
                                      colour_column=None))

    def _populate_overview_gaps(self, register, unchecked_rules) -> None:
        """
        The checks that never ran, named, with what to do about each.

        A report with no language section reads as "nothing found" when the
        truth is "never looked", and the same trap applies on screen: a register
        filtered to failures hides its own blind spots. So they get their own
        panel rather than only a row in the table.
        """
        if register.empty:
            self.overview_gaps.setText("No checks have run yet.")
            set_state(self.overview_gaps, "muted")
            return
        gaps = register[register['Result'] == "Not checked"]
        if gaps.empty:
            self.overview_gaps.setText(
                "Every check ran on this sheet. Nothing was skipped for want of "
                "a column.")
            set_state(self.overview_gaps, "ok")
            return
        lines = [f"{row['Check']}  -  {row['Detail']}"
                 for _, row in gaps.head(8).iterrows()]
        if len(gaps) > 8:
            lines.append(f"...and {len(gaps) - 8} more, listed in the register.")
        lines.append("")
        lines.append("A check that could not run is not a check that passed.")
        self.overview_gaps.setText("\n".join(lines))
        set_state(self.overview_gaps, "warn")

    def _populate_overview_facts(self, rows: int) -> None:
        cache = self.analyzer._cache if self.analyzer else {}
        mapping = cache.get('column_map', {}) or {}
        lines = [
            f"{rows:,} rows  x  {len(self.df.columns)} fields",
            f"Ticket ID: {self.analyzer.id_column or 'not identified'}",
            f"Created: {self.analyzer.created_column or 'not identified'}",
            f"Closed: {self.analyzer.closed_column or 'not identified'}",
            f"Fields matched to a requirement: "
            f"{sum(1 for v in mapping.values() if v)} of {len(mapping)}",
        ]
        if not self.language_results:
            lines.append("")
            lines.append("Language: not checked on any column yet.")
        self.overview_facts.setText("\n".join(lines))
        set_state(self.overview_facts, "muted")

    def _apply_overview_badges(self, register) -> None:
        """
        Put each view's failing-check count on its sidebar row.

        Failing checks only, and one meaning per badge: a pill wide enough for
        two digits cannot also carry "and three were skipped", and a badge that
        sometimes counts rows and sometimes counts checks would be a number
        nobody could read.
        """
        self.clear_nav_badges()
        if register.empty:
            return
        failing = register[register['Result'] == "Fail"]
        tally: Dict[str, int] = {}
        for _, row in failing.iterrows():
            view = self._view_for_finding(row.get("Area", ""),
                                         str(row.get("Check", "")))
            tally[view] = tally.get(view, 0) + 1
        for view, count in tally.items():
            self.set_nav_badge(view, count, "error")

    # 1. Column Check
    def create_column_check_tab(self, parent):
        """
        The verdict first, the mapping second, and the fields nothing reads
        third - that last table is what used to be 450px of empty page.

        Scrolled, because the mapping table's height is pinned to its nine rows:
        a QVBoxLayout with less room than its size hints need takes the space
        from whatever can shrink, and laid the labels below out on top of it at
        1000x700. That was a real failure of TestViewsDoNotOverlap, not a
        hypothetical.
        """
        layout = make_view(parent, scroll=True)

        tiles = QHBoxLayout()
        tiles.setSpacing(SPACE_SM)
        self.col_kpis = {}
        for label in ("Matched", "Missing", "Set by hand", "Fields unused"):
            frame, value = make_kpi(tiles, label)
            self.col_kpis[label] = (frame, value)
        tiles.addStretch()
        self.btn_col_map = QPushButton("Change mapping…")
        self.btn_col_map.setObjectName("PrimaryButton")
        self.btn_col_map.setToolTip(
            "Pick the column each requirement should read, or declare one "
            "absent")
        self.btn_col_map.clicked.connect(self.show_column_mapping_dialog)
        tiles.addWidget(self.btn_col_map, 0, Qt.AlignmentFlag.AlignVCenter)
        layout.addLayout(tiles)

        self.col_verdict = QLabel("")
        self.col_verdict.setWordWrap(True)
        layout.addWidget(self.col_verdict)

        # The score is the only thing that makes a wrong match visible, so what
        # it means stays on the page rather than moving into a tooltip.
        intro = QLabel(
            "Columns are matched by name, so a match can be wrong — and "
            "every other view reads whatever was matched here. The score is how "
            "confident the match is: 100 an exact name, 50 a word found inside "
            "a longer one."
        )
        intro.setObjectName("Caption")
        intro.setWordWrap(True)
        layout.addWidget(intro)

        self.col_check_table = QTableView()
        configure_table(self.col_check_table, stretch_last=True,
                        title="Column Check")
        layout.addWidget(self.col_check_table)

        # What no check reads. A required column that matched the wrong field is
        # invisible from the table above alone - the giveaway is the right field
        # sitting here, unused.
        unused_bar = QHBoxLayout()
        unused_bar.setSpacing(SPACE_SM)
        unused_title = QLabel("FIELDS NO CHECK READS")
        unused_title.setObjectName("CardTitle")
        unused_bar.addWidget(unused_title)
        self.col_unused_note = QLabel("")
        self.col_unused_note.setWordWrap(True)
        set_state(self.col_unused_note, "muted")
        unused_bar.addWidget(self.col_unused_note, 1)
        layout.addLayout(unused_bar)

        self.col_unused_table = QTableView()
        configure_table(self.col_unused_table, stretch_last=True,
                        title="Fields no check reads")
        layout.addWidget(self.col_unused_table, 1)

        self.col_override_label = QLabel()
        self.col_override_label.setWordWrap(True)
        layout.addWidget(self.col_override_label)

    def _populate_column_summary(self, choices) -> None:
        matched = [r for r, i in choices.items() if i['chosen'] is not None]
        missing = [r for r, i in choices.items() if i['chosen'] is None]
        by_hand = [r for r, i in choices.items() if i['overridden']]
        used = {i['chosen'] for i in choices.values() if i['chosen']}
        unused = [c for c in self.df.columns if c not in used]

        for label, count, severity in (("Matched", len(matched), "ok"),
                                       ("Missing", len(missing), "error"),
                                       ("Set by hand", len(by_hand), "info"),
                                       ("Fields unused", len(unused), None)):
            frame, value = self.col_kpis[label]
            value.setText(f"{count:,}")
            set_state(value, severity if count and severity != "info" else None)
            set_severity(frame, severity if count and severity == "error"
                         else None)

        if missing:
            self.col_verdict.setText(
                f"{len(missing)} of {len(choices)} required fields could not be "
                f"matched: {', '.join(missing)}. Every check that reads one of "
                f"them will report itself as not checked.")
            set_state(self.col_verdict, "error")
        else:
            self.col_verdict.setText(
                f"All {len(choices)} required fields matched a column in this "
                f"file. Check the scores below before trusting the matches.")
            set_state(self.col_verdict, "ok")

    def _populate_unused_columns(self) -> None:
        """
        The fields in the file that no requirement matched.

        Not a defect - most exports carry columns this tool has no rule for.
        It is here because a *wrong* match is invisible from the mapping table
        alone: the tell is the column that should have matched sitting in this
        list instead.
        """
        choices = self.analyzer.column_choices()
        used = {info['chosen'] for info in choices.values() if info['chosen']}
        nulls = self.analyzer.check_nulls(
            float(self.config.get("null_threshold", 20)))
        rows = len(self.df)

        unused = [c for c in self.df.columns if c not in used]
        model = QStandardItemModel(len(unused), 3)
        model.setHorizontalHeaderLabels(["Field", "Type", "Filled (%)"])
        for i, column in enumerate(unused):
            model.setItem(i, 0, QStandardItem(str(column)))
            model.setItem(i, 1, QStandardItem(
                str(self.df[column].dtype.categories.dtype
                    if hasattr(self.df[column].dtype, "categories")
                    else self.df[column].dtype)))
            share = 100.0 - float(nulls.get(column, {}).get('percentage', 0.0))
            filled = QStandardItem()
            filled.setData(round(share, 1), Qt.ItemDataRole.DisplayRole)
            if rows and share == 0:
                filled.setForeground(QBrush(QColor(COLOR_WARNING)))
            model.setItem(i, 2, filled)
        set_table_model(self.col_unused_table, model)

        self.col_unused_note.setText(
            "Every field in this file is read by a check."
            if not unused else
            f"{len(unused)} of {len(self.df.columns)} fields match no "
            f"requirement. That is normal - but if one of them is the field a "
            f"check should have used, set it above.")
        set_copy_context(self.col_unused_table, **{
            "Fields in sheet": f"{len(self.df.columns)}",
            "Fields read by a check": f"{len(used)}",
        })

    def refresh_column_check(self):
        choices = self.analyzer.column_choices()
        self._populate_column_summary(choices)
        # After the mapping table, so the pinned height is measured against a
        # populated model: the unused table below it takes the leftover space,
        # which is what stops this view being nine rows and half a page of
        # white.
        self._populate_unused_columns()

        headers = ["Required Column", "Status", "Matched Column", "Score",
                   "Other candidates"]
        model = QStandardItemModel(len(choices), len(headers))
        model.setHorizontalHeaderLabels(headers)

        for i, (req_col, info) in enumerate(choices.items()):
            model.setItem(i, 0, QStandardItem(req_col))

            found = info['chosen'] is not None
            if info['overridden']:
                status, colour = "Set by you", COLOR_ACCENT
            elif found:
                status, colour = "Found", COLOR_SUCCESS
            else:
                status, colour = "Missing", COLOR_ERROR
            status_item = QStandardItem(status)
            status_item.setForeground(QBrush(QColor(colour)))
            model.setItem(i, 1, status_item)

            model.setItem(i, 2, QStandardItem(info['chosen'] or ""))

            score_item = QStandardItem()
            score_item.setData(info['score'], Qt.ItemDataRole.DisplayRole)
            # A low-scoring match is the one worth a second look.
            if found and not info['overridden'] and info['score'] < 70:
                score_item.setForeground(QBrush(QColor(COLOR_WARNING)))
            model.setItem(i, 3, score_item)

            # Naming the runners-up with their scores is what makes an arbitrary
            # tie visible: impact and urgency both score 130 for Priority.
            alternatives = ", ".join(f"{col} ({score})"
                                     for col, score in info['alternatives'][:4])
            model.setItem(i, 4, QStandardItem(alternatives))

        set_table_model(self.col_check_table, model)
        # Nine requirements is the whole list, so it should be visible at once
        # rather than scrolling inside a panel while the table below it sits
        # half empty.
        fit_table_height(self.col_check_table, max_rows=12)
        set_copy_context(self.col_check_table,
                         **{"Mappings changed by hand":
                            ", ".join(sorted(self.analyzer.overridden)) or "none"})

        notes = self.analyzer.override_notes
        self.col_override_label.setText("  ".join(notes))
        set_state(self.col_override_label,
                  "warn" if any("ignored" in n for n in notes)
                  else "ok" if notes else "muted")

    def show_column_mapping_dialog(self):
        """
        Let the user correct a wrong automatic match.

        Rebuilds the analyzer on accept, for the same reason apply_date_format
        does: the mapping decides the ID column, the date columns and every
        check that reads them, so a partially updated cache would leave views
        describing different columns of the same file.
        """
        if self.df is None or self.analyzer is None:
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Change column mapping")
        dialog.setMinimumWidth(620)
        layout = QVBoxLayout(dialog)

        blurb = QLabel(
            "Pick the column each category should read. “Not present” marks a "
            "category as genuinely absent, so its checks report as not checked "
            "rather than silently reading the wrong field."
        )
        blurb.setWordWrap(True)
        set_state(blurb, "muted")
        layout.addWidget(blurb)

        combos = {}
        choices = self.analyzer.column_choices()
        for required, info in choices.items():
            row = QHBoxLayout()
            label = QLabel(required)
            label.setMinimumWidth(160)
            row.addWidget(label)

            combo = QComboBox()
            combo.addItem("Automatic", None)
            combo.addItem("Not present", COLUMN_ABSENT)
            scores = dict(info['alternatives'])
            if info['chosen']:
                scores[info['chosen']] = info['score']
            for column in self.df.columns:
                score = scores.get(column)
                suffix = f"  (score {score})" if score else ""
                combo.addItem(f"{column}{suffix}", column)

            existing = self.config.get("column_overrides", {}).get(required)
            index = combo.findData(existing) if existing else 0
            combo.setCurrentIndex(max(index, 0))
            row.addWidget(combo, stretch=1)
            combos[required] = combo
            layout.addLayout(row)

        buttons = QHBoxLayout()
        buttons.addStretch()
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(dialog.reject)
        buttons.addWidget(btn_cancel)
        btn_apply = QPushButton("Apply")
        btn_apply.setObjectName("PrimaryButton")
        btn_apply.clicked.connect(dialog.accept)
        buttons.addWidget(btn_apply)
        layout.addLayout(buttons)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        overrides = {required: combo.currentData()
                     for required, combo in combos.items()
                     if combo.currentData() is not None}
        self.apply_column_overrides(overrides)

    def apply_column_overrides(self, overrides: Dict[str, str]) -> None:
        """Persist the mapping and recompute everything that depends on it."""
        if overrides == (self.config.get("column_overrides") or {}):
            return

        self.config["column_overrides"] = overrides
        self.config_mgr.config = self.config
        self.config_mgr.save()

        self.analyzer = SanityAnalyzer(
            self.df, dayfirst=self.analyzer.dayfirst if self.analyzer else None,
            column_overrides=overrides)
        self._data_seq += 1
        self._refresh_inflight.clear()
        self._analyzer_lock = threading.Lock()
        # Language findings name a column; a remapping can change which.
        self.language_results.clear()

        self.update_status(
            f"Column mapping updated ({len(overrides)} set by hand). "
            f"All checks recalculated." if overrides
            else "Column mapping reset to automatic. All checks recalculated.")
        self.refresh_tab(self.current_view)

    # 2. Null Analysis
    def create_null_analysis_tab(self, parent):
        layout = make_view(parent)

        tiles = QHBoxLayout()
        tiles.setSpacing(SPACE_SM)
        self.null_kpis = {}
        for label in ("Over threshold", "Entirely empty", "Emptiest field",
                      "Empty cells"):
            frame, value = make_kpi(tiles, label)
            self.null_kpis[label] = (frame, value)
        tiles.addStretch()
        layout.addLayout(tiles)

        # The conclusion, before the evidence. This view used to be a bare
        # table: sixteen rows of numbers and no statement of what they meant.
        self.null_verdict = QLabel("")
        self.null_verdict.setWordWrap(True)
        layout.addWidget(self.null_verdict)

        self.null_table = QTableView()
        configure_table(self.null_table, title="Null Analysis",
                        stretch_last=True)
        layout.addWidget(self.null_table, 1)

    def refresh_null_analysis(self):
        threshold = float(self.config.get("null_threshold", 20))
        results = self.analyzer.check_nulls(threshold)
        rows = len(self.df)

        # "Status" is a word, not only a tint. A red percentage is invisible to
        # a reader who cannot see red, and it is also invisible in a paste.
        model = QStandardItemModel(len(results), 4)
        model.setHorizontalHeaderLabels(
            ["Column", "Null Count", "Percentage (%)", "Status"])
        for i, (col, data) in enumerate(results.items()):
            model.setItem(i, 0, QStandardItem(col))

            count_item = QStandardItem()
            count_item.setData(data['count'], Qt.ItemDataRole.DisplayRole)
            model.setItem(i, 1, count_item)

            pct_item = QStandardItem()
            pct_item.setData(data['percentage'], Qt.ItemDataRole.DisplayRole)
            model.setItem(i, 2, pct_item)

            over = data['above_threshold']
            blank = bool(rows) and data['count'] >= rows
            status = QStandardItem("Entirely empty" if blank
                                   else "Over threshold" if over else "OK")
            if over:
                tint = QBrush(QColor(COLOR_ERROR))
                pct_item.setForeground(tint)
                status.setForeground(tint)
            else:
                status.setForeground(QBrush(QColor(COLOR_TEXT_MUTED)))
            model.setItem(i, 3, status)

        set_table_model(self.null_table, model)

        breaching = [c for c, d in results.items() if d['above_threshold']]
        empty = [c for c, d in results.items()
                 if rows and d['count'] >= rows]
        worst_name, worst = max(results.items(),
                                key=lambda kv: kv[1]['percentage'],
                                default=(None, {'percentage': 0.0}))
        cells = rows * len(results)
        blanks = sum(d['count'] for d in results.values())

        # Only the two counts take a colour. The two percentages are context,
        # not findings: an emptiest field of 0.8% tinted amber is a false alarm,
        # and it is the threshold above that decides what counts as a problem.
        for label, text, severity in (
                ("Over threshold", f"{len(breaching):,}",
                 "error" if breaching else None),
                ("Entirely empty", f"{len(empty):,}",
                 "error" if empty else None),
                ("Emptiest field", f"{worst['percentage']:.1f}%", None),
                ("Empty cells",
                 f"{(blanks / cells * 100 if cells else 0):.1f}%", None)):
            frame, value = self.null_kpis[label]
            value.setText(text)
            set_state(value, severity)
            set_severity(frame, severity)
        # The name goes in a tooltip rather than the label: a field name is as
        # long as it likes, and the tile is 150px.
        self.null_kpis["Emptiest field"][1].setToolTip(worst_name or "")

        if breaching:
            listed = ", ".join(breaching[:6])
            more = (f" and {len(breaching) - 6} more."
                    if len(breaching) > 6 else ".")
            self.null_verdict.setText(
                f"{len(breaching)} of {len(results)} fields are more than "
                f"{threshold:g}% empty: {listed}{more}")
            set_state(self.null_verdict, "error")
        else:
            self.null_verdict.setText(
                f"No field in this sheet is more than {threshold:g}% empty. "
                f"The threshold lives in Settings.")
            set_state(self.null_verdict, "ok")

        set_copy_context(self.null_table, **{
            "Threshold": f"{threshold:g}% null",
            "Rows analysed": f"{rows:,}",
            "Fields over threshold": f"{len(breaching)} of {len(results)}",
        })

    # 3. Logic Checks - a finding-first workbench
    #
    # The eight rules, in the order a reviewer wants them, with everything the
    # inspector needs that is *not* a number: the flag whose mask holds the
    # affected rows, the fields the rule reads, and the three pieces of prose.
    #
    # Declaring them here rather than inline in a populate is what makes this a
    # workbench instead of a stack of bands. The table, the inspector and the
    # records dialog all read one record per rule, so they cannot disagree about
    # what a rule found - which is exactly what five parallel if-ladders used to
    # risk. `count_key` names the field in check_date_logic /
    # check_cross_field_logic that holds the headline, and `flag` names the mask
    # in get_row_flags that holds the same rows; the two are summed from the
    # same array by the analyzer, which is why they can be trusted together.
    LOGIC_RULES = [
        {
            'key': 'closed_before_created',
            'title': "Resolution precedes creation",
            'note': "A negative lifecycle duration is impossible",
            'severity': "Critical",
            'confidence': "Direct",
            'scope': ('created_column', 'closed_column'),
            'why': "A record that closes before it opens gives a negative "
                   "duration. Any mean time to resolution computed over this "
                   "file is pulled toward zero, or below it, by these rows.",
            'definition': "The resolution timestamp is strictly earlier than "
                          "the creation timestamp. Rows where either date is "
                          "unparsed are not counted, because a missing date is "
                          "a completeness problem, not a sequence one.",
            'advice': "Almost always a field-order misread or a bulk import "
                      "with swapped columns. Check the Date format control "
                      "above before treating it as a data-entry fault.",
        },
        {
            'key': 'future_created',
            'title': "Creation date in the future",
            'note': "Recorded as opening after the file was read",
            'severity': "Medium",
            'confidence': "Direct",
            'scope': ('created_column',),
            'why': "A future creation date puts the record outside every "
                   "reporting period it should appear in, so it quietly "
                   "disappears from monthly volumes.",
            'definition': "The creation timestamp is later than the moment "
                          "this file was analysed. That instant is pinned once "
                          "per load, so the count does not drift while you "
                          "read it.",
            'advice': "Usually a typed year. Compare against the row's "
                      "resolution date, which is normally right.",
        },
        {
            'key': 'future_closed',
            'title': "Resolution date in the future",
            'note': "Recorded as resolved after the file was read",
            'severity': "Medium",
            'confidence': "Direct",
            'scope': ('closed_column',),
            'why': "A future resolution date makes a record look closed in a "
                   "period that has not happened, and inflates closure rates "
                   "for the current month.",
            'definition': "The resolution timestamp is later than the moment "
                          "this file was analysed.",
            'advice': "Check whether the field is being used to hold a target "
                      "or due date rather than an actual resolution.",
        },
        {
            'key': 'open_with_date',
            'title': "Open record has a resolution date",
            'note': "State contradicts the populated resolution field",
            'severity': "High",
            'confidence': "Direct",
            'scope': ('state_column', 'closed_column'),
            'why': "These records can be counted as both open and resolved, "
                   "inflating backlog and closure measures at the same time.",
            'definition': "The state reads as active while the resolution "
                          "field holds a value. States that read as neither "
                          "open nor closed are excluded rather than forced "
                          "into either group.",
            'advice': "Decide which field is authoritative before reporting "
                      "either number. This changes the narrative, never the "
                      "source data.",
        },
        {
            'key': 'closed_no_date',
            'title': "Closed record has no resolution date",
            'note': "Resolution timestamp never recorded",
            'severity': "High",
            'confidence': "Direct",
            'scope': ('state_column', 'closed_column'),
            'why': "Resolution time cannot be computed for these records, so "
                   "every duration statistic silently describes a subset of "
                   "the file rather than all of it.",
            'definition': "The state reads as closed or resolved while the "
                          "resolution field is empty.",
            'advice': "State the subset size alongside any resolution-time "
                      "figure taken from this file.",
        },
        {
            'key': 'open_missing_priority',
            'title': "Open record has no priority",
            'note': "Cannot be triaged or measured against an SLA",
            'severity': "Medium",
            'confidence': "Direct",
            'scope': ('state_column', 'priority_column'),
            'why': "An open record with no priority is invisible to any "
                   "queue that sorts by urgency, and has no SLA to breach.",
            'definition': "The state reads as active and the priority field "
                          "is empty. Every flavour of blank counts - null, "
                          "empty string and whitespace.",
            'advice': "Worth raising with the queue owner rather than fixing "
                      "in the export: these rows are live work.",
        },
    ]
    LOGIC_TABS = ["Findings", "Passed", "Not run"]
    LOGIC_COLUMNS = ["Severity", "Rule", "Scope", "Records", "Confidence"]
    # Rows shown in the panel's evidence table. 200 is enough to scroll through
    # and judge a pattern; the full set - which can be six figures - goes to its
    # own window, where the row count is stated rather than implied.
    LOGIC_EVIDENCE_ROWS = 200

    def create_logic_checks_tab(self, parent):
        """
        Findings first, evidence second, methodology in the inspector.

        The five full-width tinted bands this replaced spent 1600px of window
        on one sentence each and gave a reader no way to see the 1,560 rows they
        were talking about. Here the rules are a table, the selected one is
        explained beside it, and its rows are one click away.

        The view itself does not scroll: nothing in the left column pins a
        height, so the findings table absorbs the deficit. The inspector body
        does scroll, because the evidence sample is pinned to its content.
        """
        layout = make_view(parent)

        # The date-order control stays at the top and stays conditional. It is
        # not a finding - it is the input every rule below it depends on, and
        # reading 05/03 the wrong way round turns a 3-day resolution into 92
        # with nothing reporting a problem.
        self.lc_format_card, fmt_l = make_card(layout, margin=SPACE_MD)
        fmt_title = QLabel("DATE FIELD ORDER")
        fmt_title.setObjectName("CardTitle")
        fmt_l.addWidget(fmt_title)
        self.lc_format_label = QLabel()
        self.lc_format_label.setWordWrap(True)
        fmt_l.addWidget(self.lc_format_label)

        fmt_buttons = QHBoxLayout()
        fmt_buttons.addWidget(QLabel("Read these dates as:"))
        self.lc_format_combo = QComboBox()
        for label, value in (("Infer from the data", None),
                             ("DD/MM/YYYY (day first)", True),
                             ("MM/DD/YYYY (month first)", False)):
            self.lc_format_combo.addItem(label, value)
        fmt_buttons.addWidget(self.lc_format_combo)
        btn_apply_format = QPushButton("Apply")
        btn_apply_format.setObjectName("PrimaryButton")
        btn_apply_format.clicked.connect(self.apply_date_format)
        fmt_buttons.addWidget(btn_apply_format)
        fmt_buttons.addStretch()
        fmt_l.addLayout(fmt_buttons)
        self.lc_format_card.setVisible(False)

        tiles = QHBoxLayout()
        tiles.setSpacing(SPACE_SM)
        self.lc_kpis = {}
        for label in ("Rows with an issue", "Rules with findings",
                      "Rules passed", "Rules not run"):
            frame, value = make_kpi(tiles, label)
            self.lc_kpis[label] = (frame, value)
        tiles.addStretch()
        layout.addLayout(tiles)

        # Which columns the rules actually read. It used to be a full-width
        # tinted card of its own; it is three facts and belongs on one line.
        self.lc_fields_label = QLabel("")
        self.lc_fields_label.setObjectName("Caption")
        self.lc_fields_label.setWordWrap(True)
        layout.addWidget(self.lc_fields_label)

        body = QHBoxLayout()
        body.setSpacing(SPACE_SM)

        # The left column scrolls, the inspector does not. Wrapping the whole
        # view instead would put the inspector - which is pinned head-to-foot
        # with its actions at the bottom - onto the same scrollbar, so its Open
        # records button would slide off the screen.
        panel_scroll = QScrollArea()
        panel_scroll.setWidgetResizable(True)
        panel_scroll.setFrameShape(QFrame.Shape.NoFrame)

        panel, panel_layout = make_card(margin=SPACE_SM)
        panel_scroll.setWidget(panel)
        bar = QHBoxLayout()
        bar.setSpacing(SPACE_XS)
        self.lc_tabs = make_seg_tabs(bar, self.LOGIC_TABS, self._show_logic_tab)
        bar.addStretch()
        panel_layout.addLayout(bar)

        self.lc_tab_note = QLabel("")
        self.lc_tab_note.setWordWrap(True)
        set_state(self.lc_tab_note, "muted")
        panel_layout.addWidget(self.lc_tab_note)

        # Pinned to its content, capped so a long list scrolls inside itself
        # rather than pushing the evidence off the page. Eight rules is the
        # whole rule set, so in practice it never scrolls.
        self.lc_findings_table = QTableView()
        configure_table(self.lc_findings_table, title="Logic rules",
                        stretch_last=True)
        panel_layout.addWidget(self.lc_findings_table)

        # The rows behind the selected rule, at full width. They started in the
        # inspector, where the mockup put them - but the inspector is 360px and
        # a ticket ID, a state and a timestamp do not fit in it, while the panel
        # beside it had 600px of white space under a one-row table. Evidence is
        # the thing that deserves the width.
        evidence_bar = QHBoxLayout()
        evidence_bar.setSpacing(SPACE_SM)
        self.lc_evidence_title = QLabel("EVIDENCE")
        self.lc_evidence_title.setObjectName("CardTitle")
        evidence_bar.addWidget(self.lc_evidence_title)
        self.lc_evidence_note = QLabel("")
        self.lc_evidence_note.setWordWrap(True)
        set_state(self.lc_evidence_note, "muted")
        evidence_bar.addWidget(self.lc_evidence_note, 1)
        panel_layout.addLayout(evidence_bar)

        self.lc_evidence_table = QTableView()
        configure_table(self.lc_evidence_table, sortable=True, stretch_last=True,
                        title="Logic Checks - evidence")
        panel_layout.addWidget(self.lc_evidence_table, 1)
        body.addWidget(panel_scroll, 1)
        body.addWidget(self._build_logic_inspector())

        layout.addLayout(body, 1)
        self._logic_records = []
        self._logic_tab = "Findings"
        self._logic_selected = None

    def _build_logic_inspector(self) -> QFrame:
        """
        The evidence panel: what the selected rule found, and the rows it found.

        Persistent rather than a dialog, because triage is a loop - read a
        finding, look at three of its rows, move to the next. A modal would put
        an open-and-close between every step.
        """
        panel = QFrame()
        panel.setObjectName("Inspector")
        panel.setMinimumWidth(INSPECTOR_MIN_WIDTH)
        panel.setMaximumWidth(INSPECTOR_WIDTH)
        column = QVBoxLayout(panel)
        column.setContentsMargins(0, 0, 0, 0)
        column.setSpacing(0)

        head = QFrame()
        head.setObjectName("InspectorHead")
        head_layout = QVBoxLayout(head)
        head_layout.setContentsMargins(SPACE_MD, SPACE_MD, SPACE_MD, SPACE_MD)
        head_layout.setSpacing(SPACE_XS)
        self.lc_ins_kicker = QLabel("")
        self.lc_ins_kicker.setObjectName("Eyebrow")
        self.lc_ins_kicker.setWordWrap(True)
        head_layout.addWidget(self.lc_ins_kicker)
        self.lc_ins_title = QLabel("")
        self.lc_ins_title.setObjectName("InspectorTitle")
        self.lc_ins_title.setWordWrap(True)
        head_layout.addWidget(self.lc_ins_title)
        self.lc_ins_summary = QLabel("")
        self.lc_ins_summary.setWordWrap(True)
        set_state(self.lc_ins_summary, "muted")
        head_layout.addWidget(self.lc_ins_summary)
        column.addWidget(head)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        inner = QWidget()
        inner_layout = QVBoxLayout(inner)
        inner_layout.setContentsMargins(SPACE_MD, SPACE_MD, SPACE_MD, SPACE_MD)
        inner_layout.setSpacing(SPACE_SM)

        stats = QHBoxLayout()
        stats.setSpacing(SPACE_SM)
        self.lc_ins_stats = {}
        for label in ("affected records", "of this sheet"):
            frame, value = make_kpi(stats, label)
            self.lc_ins_stats[label] = (frame, value)
        inner_layout.addLayout(stats)

        self.lc_ins_sections = {}
        for name in ("Why this matters", "Rule definition"):
            header = QLabel(name.upper())
            header.setObjectName("CardTitle")
            inner_layout.addWidget(header)
            body = QLabel("")
            body.setWordWrap(True)
            inner_layout.addWidget(body)
            self.lc_ins_sections[name] = header
            self.lc_ins_sections[name + " body"] = body

        note = QFrame()
        note.setObjectName("NoteBlock")
        note_layout = QVBoxLayout(note)
        note_layout.setContentsMargins(SPACE_MD, SPACE_SM, SPACE_MD, SPACE_SM)
        note_layout.setSpacing(2)
        note_title = QLabel("WHAT TO DO")
        note_title.setObjectName("CardTitle")
        note_layout.addWidget(note_title)
        self.lc_ins_advice = QLabel("")
        self.lc_ins_advice.setWordWrap(True)
        note_layout.addWidget(self.lc_ins_advice)
        inner_layout.addWidget(note)

        inner_layout.addStretch()
        scroll.setWidget(inner)
        column.addWidget(scroll, 1)

        foot = QFrame()
        foot.setObjectName("InspectorFoot")
        foot_layout = QHBoxLayout(foot)
        foot_layout.setContentsMargins(SPACE_MD, SPACE_SM, SPACE_MD, SPACE_SM)
        foot_layout.setSpacing(SPACE_SM)
        self.btn_lc_copy = QPushButton("Copy finding")
        self.btn_lc_copy.setIcon(icon("copy", COLOR_TEXT_MUTED))
        self.btn_lc_copy.setToolTip(
            "The finding, its numbers and the file it came from, as text")
        self.btn_lc_copy.clicked.connect(self._copy_logic_finding)
        foot_layout.addWidget(self.btn_lc_copy)
        self.btn_lc_rows = QPushButton("Open records")
        self.btn_lc_rows.setObjectName("PrimaryButton")
        self.btn_lc_rows.clicked.connect(self._open_logic_rows)
        foot_layout.addWidget(self.btn_lc_rows, 1)
        column.addWidget(foot)

        self.lc_inspector = panel
        return panel

    # --- Logic Checks: collect ------------------------------------------

    def refresh_logic_checks(self):
        """
        Build one record per rule, then paint.

        Still synchronous: every call below is cached on the analyzer and the
        whole thing measures inside the same budget the old version did. The
        evidence samples are the only new work, and each is one head() over a
        boolean mask.
        """
        self._populate_date_format()
        self._logic_records = self._collect_logic_records()
        self._populate_logic_fields()
        self._populate_logic_kpis()
        self._show_logic_tab(self._logic_tab)

    def _collect_logic_records(self) -> List[Dict[str, Any]]:
        """
        One record per rule: its outcome, its count and its rows.

        Counts come from get_row_flags's masks, which is the same array
        check_date_logic and check_cross_field_logic sum - so a rule's headline
        and its rows can never describe different sets. A rule absent from
        'flags' and named in 'skipped' is *not checked*, never zero: that
        distinction is the whole reason those two keys exist.
        """
        analyzer = self.analyzer
        total = len(self.df)
        row_flags = analyzer.get_row_flags()
        masks = row_flags.get('flags', {})
        skipped = set(row_flags.get('skipped', []))
        xfield = analyzer.check_cross_field_logic()
        vocabulary = analyzer.check_state_values()

        columns = {
            'created_column': analyzer.created_column,
            'closed_column': analyzer.closed_column,
            'state_column': xfield.get('state_col'),
            'priority_column': xfield.get('priority_col'),
        }

        records: List[Dict[str, Any]] = []
        for rule in self.LOGIC_RULES:
            fields = [columns.get(name) for name in rule['scope']]
            named = [f for f in fields if f]
            record = dict(rule)
            record['fields'] = named
            record['scope_text'] = " + ".join(named) if named else "no column"
            if rule['key'] in skipped or not named:
                record.update(outcome="Not checked", count=None, mask=None,
                              reason=self._logic_skip_reason(rule, columns))
            else:
                mask = masks.get(rule['key'])
                count = int(mask.sum()) if mask is not None else 0
                record.update(outcome="Fail" if count else "Pass",
                              count=count, mask=mask, reason=None)
            records.append(record)

        records.append(self._logic_unclassified_record(xfield, vocabulary))
        records.append(self._logic_vocabulary_record(vocabulary))
        return records

    def _logic_skip_reason(self, rule, columns) -> str:
        """Why a rule could not run, and which column would let it."""
        needed = {
            'created_column': "a Created date column",
            'closed_column': "a Closed or Resolved date column",
            'state_column': "a State or Status column",
            'priority_column': "a Priority column",
        }
        missing = [needed[name] for name in rule['scope']
                   if not columns.get(name)]
        if not missing:
            return ("The columns exist but hold no readable values, so the "
                    "rule had nothing to compare.")
        return ("No column in this file could be matched to "
                + " or ".join(missing)
                + ". Set it by hand in Column Check and this rule will run.")

    def _logic_unclassified_record(self, xfield, vocabulary) -> Dict[str, Any]:
        """
        States that read as neither open nor closed.

        A judgement, not a defect: the vocabulary is somebody else's, and a
        value this tool cannot place is a limit of the tool. So the outcome is
        Review, and the rows are excluded from the open/closed totals rather
        than forced into a bucket.
        """
        count = xfield.get('unclassified_count', 0) if xfield.get('has_state') \
            else None
        state_column = xfield.get('state_col')
        record = {
            'key': 'unclassified_state',
            'title': "State reads as neither open nor closed",
            'note': "Excluded from the open and closed totals, not assumed",
            'severity': "Low",
            'confidence': "Judgement",
            'fields': [state_column] if state_column else [],
            'scope_text': state_column or "no column",
            'why': "Every rule above that mentions a state depends on this "
                   "classification. Rows it cannot place are left out of both "
                   "totals, so those totals describe fewer rows than the file "
                   "holds.",
            'definition': "A value is read as closed if it contains a "
                          "terminal keyword, and as open if it contains an "
                          "active one - terminal is tested first, so "
                          "\u201cResolved Pending Closure\u201d counts as "
                          "closed. Anything matching neither is counted here.",
            'advice': "Look at the values in the Findings row below and "
                      "decide which side each belongs on. This is a "
                      "vocabulary question about your service desk, not a "
                      "defect in the export.",
            'count': count,
            'mask': None,
            'outcome': ("Not checked" if count is None
                        else "Review" if count else "Pass"),
            'reason': (None if count is not None else
                       "No column in this file could be matched to a State or "
                       "Status column. Set it by hand in Column Check and this "
                       "rule will run."),
        }
        # The count comes from check_cross_field_logic and the *values* from
        # check_state_values, which is the only place a per-value classification
        # exists. Without them this rule can report a number nobody can trace
        # back to a value, which is the thing it exists to avoid.
        record['unclassified_values'] = self._unclassified_state_values(vocabulary)
        if count and record['unclassified_values']:
            record['note'] = "Values: " + ", ".join(
                entry['value'] for entry in record['unclassified_values'][:6])
        return record

    @staticmethod
    def _unclassified_state_values(vocabulary) -> List[Dict[str, Any]]:
        entries = ((vocabulary.get('standard_values') or [])
                   + (vocabulary.get('suspect_values') or []))
        unknown = [e for e in entries if e.get('classification') == 'unknown']
        return sorted(unknown, key=lambda e: -e.get('count', 0))

    def _logic_vocabulary_record(self, vocabulary) -> Dict[str, Any]:
        """
        State values that look like typos or one-offs.

        Rarity alone is the weaker signal, so a value one edit from a *more
        common* value is reported whatever its share - which is what catches
        'Closd' at 1.7% of a 60-row file where the frequency rule stays quiet,
        correctly, since 1 in 20 is normal at that size.
        """
        has_state = vocabulary.get('has_state')
        suspect = vocabulary.get('suspect_values') or []
        count = vocabulary.get('total_suspect_rows', 0) if has_state else None
        note = "Rare values, and values one edit from a commoner one"
        if suspect:
            note = "Values: " + ", ".join(
                f"\u201c{e['value']}\u201d" for e in suspect[:4])
        return {
            'key': 'state_vocabulary',
            'title': "Suspect state value",
            'note': note,
            'severity': "Medium",
            'confidence': "Judgement",
            'fields': [vocabulary.get('state_col')] if has_state else [],
            'scope_text': vocabulary.get('state_col') or "no column",
            'why': "A typo makes a state its own category, so a queue splits "
                   "in two and neither half looks wrong on its own.",
            'definition': "Two independent tests: a value under both the "
                          "share and count thresholds, or a value within one "
                          "edit of a more common one. The pairwise pass is "
                          "skipped above a few hundred distinct values, so a "
                          "misdetected free-text column cannot trigger an "
                          "O(n\u00b2) scan.",
            'advice': "Merge at source if it is a typo. If it is a real "
                      "state that is simply rare, nothing needs doing - this "
                      "is a judgement and it can be wrong.",
            'count': count,
            'mask': None,
            'outcome': ("Not checked" if count is None
                        else "Review" if count else "Pass"),
            'reason': (None if count is not None else
                       "No column in this file could be matched to a State or "
                       "Status column. Set it by hand in Column Check and this "
                       "rule will run."),
            'vocabulary': vocabulary,
        }

    # --- Logic Checks: paint --------------------------------------------

    def _populate_logic_fields(self) -> None:
        analyzer = self.analyzer
        xfield = analyzer.check_cross_field_logic()
        pairs = [("Created", analyzer.created_column),
                 ("Resolved", analyzer.closed_column),
                 ("State", xfield.get('state_col')),
                 ("Priority", xfield.get('priority_col'))]
        # The column name is quoted because the role and the column are often
        # the same word: "Created: Created" reads as a mistake, Created:
        # \u201cCreated\u201d does not.
        parts = [f"{label}: \u201c{column}\u201d" if column
                 else f"{label}: not matched" for label, column in pairs]
        self.lc_fields_label.setText(
            f"{len(self.df):,} records.   Fields read \u2014 "
            + "     ".join(parts))

    def _populate_logic_kpis(self) -> None:
        records = self._logic_records
        failing = [r for r in records if r['outcome'] == "Fail"]
        review = [r for r in records if r['outcome'] == "Review"]
        passed = [r for r in records if r['outcome'] == "Pass"]
        notrun = [r for r in records if r['outcome'] == "Not checked"]
        rows = sum(r['count'] or 0 for r in failing)

        for label, count, severity in (
                ("Rows with an issue", rows, "error"),
                ("Rules with findings", len(failing) + len(review), "error"),
                ("Rules passed", len(passed), "ok"),
                ("Rules not run", len(notrun), "warn")):
            frame, value = self.lc_kpis[label]
            value.setText(f"{count:,}")
            set_state(value, severity if count else None)
            set_severity(frame, severity if count and severity != "ok" else None)

        for name, group in (("Findings", failing + review),
                            ("Passed", passed), ("Not run", notrun)):
            self.lc_tabs[name].setText(f"{name} {len(group)}")

    def _logic_group(self, name: str) -> List[Dict[str, Any]]:
        if name == "Findings":
            wanted = ("Fail", "Review")
        elif name == "Passed":
            wanted = ("Pass",)
        else:
            wanted = ("Not checked",)
        return [r for r in self._logic_records if r['outcome'] in wanted]

    def _show_logic_tab(self, name: str) -> None:
        self._logic_tab = name
        group = self._logic_group(name)
        frame = pd.DataFrame([{
            'Severity': (record['severity'] if record['outcome'] in ("Fail", "Review")
                         else record['outcome']),
            'Rule': record['title'],
            'Scope': record['scope_text'],
            'Records': ("Not checked" if record['count'] is None
                        else f"{record['count']:,}"),
            'Confidence': record['confidence'],
        } for record in group], columns=self.LOGIC_COLUMNS)
        set_table_model(self.lc_findings_table,
                        FindingsModel(frame, colour_column="Severity",
                                      colours=SEVERITY_STATES))
        # Pinned to its content: a one-row rule list left 150px of white below
        # itself, which is the emptiness this rebuild exists to remove. The left
        # column is a scroll area for exactly this reason.
        fit_table_height(self.lc_findings_table, max_rows=8)
        selection = self.lc_findings_table.selectionModel()
        if selection is not None:
            selection.selectionChanged.connect(self._on_logic_row_selected)

        others = {
            "Findings": [("passed", "Passed"), ("could not run", "Not run")],
            "Passed": [("found something", "Findings"),
                       ("could not run", "Not run")],
            "Not run": [("found something", "Findings"), ("passed", "Passed")],
        }[name]
        # Says what is in the other tabs, so a one-row Findings list cannot read
        # as "one rule exists". The mockup carried the same line under its
        # findings list, for the same reason.
        tail = ", ".join(f"{len(self._logic_group(tab))} {verb}"
                         for verb, tab in others)
        notes = {
            "Findings": "Rules that found something. Select one to see why it "
                        "matters and the rows behind it.  " + tail + ".",
            "Passed": "Rules that ran and found nothing. Kept where you can "
                      "check them, not deleted.  " + tail + ".",
            "Not run": "Rules that could not run at all. A rule that could "
                       "not run is not a rule that passed.  " + tail + ".",
        }
        empty = {
            "Findings": "No logic rule found anything on this sheet.",
            "Passed": "No rule ran cleanly on this sheet.",
            "Not run": "Every rule ran. Nothing was skipped for want of a "
                       "column.",
        }
        self.lc_tab_note.setText(notes[name] if group else empty[name])

        set_copy_context(self.lc_findings_table, **{
            "Tab": name,
            "Rules shown": f"{len(group)} of {len(self._logic_records)}",
            "Rows in sheet": f"{len(self.df):,}",
        })

        if group:
            self.lc_findings_table.selectRow(0)
            self._select_logic_record(group[0])
        else:
            self._select_logic_record(None)

    def _on_logic_row_selected(self, *_args) -> None:
        selection = self.lc_findings_table.selectionModel()
        rows = selection.selectedRows() if selection else []
        group = self._logic_group(self._logic_tab)
        if rows and rows[0].row() < len(group):
            self._select_logic_record(group[rows[0].row()])

    def _select_logic_record(self, record) -> None:
        """Fill the inspector for one rule, or empty it when there is none."""
        self._logic_selected = record
        if record is None:
            self.lc_ins_kicker.setText("NOTHING SELECTED")
            self.lc_ins_title.setText("No rule to inspect")
            self.lc_ins_summary.setText(
                "This tab is empty, so there is nothing to explain. The other "
                "tabs hold the rules that did run.")
            for label in self.lc_ins_stats:
                frame, value = self.lc_ins_stats[label]
                value.setText("\u2014")
                set_state(value, None)
                set_severity(frame, None)
            self.lc_ins_sections["Why this matters body"].setText("")
            self.lc_ins_sections["Rule definition body"].setText("")
            self.lc_ins_advice.setText("")
            self._set_logic_evidence(pd.DataFrame(),
                                     "Select a rule to see its rows.")
            self.btn_lc_rows.setEnabled(False)
            self.btn_lc_rows.setText("Open records")
            self.btn_lc_copy.setEnabled(False)
            return

        total = len(self.df) or 1
        count = record['count']
        outcome = record['outcome']
        state = {"Fail": "error", "Review": "warn", "Not checked": "warn",
                 "Pass": "ok"}[outcome]

        self.lc_ins_kicker.setText(
            f"{record['severity'] if outcome in ('Fail', 'Review') else outcome}"
            f"  \u00b7  {record['confidence'].lower()} evidence"
            if outcome in ("Fail", "Review")
            else f"{outcome}  \u00b7  {record['confidence'].lower()} evidence")
        self.lc_ins_title.setText(record['title'])

        if count is None:
            self.lc_ins_summary.setText(record['reason'] or
                                        "This rule could not run on this file.")
        elif count:
            share = count / total * 100
            self.lc_ins_summary.setText(
                f"{count:,} of {len(self.df):,} records "
                f"({share:.2f}%) match this rule. {record['note']}")
        else:
            self.lc_ins_summary.setText(
                f"No record in this sheet matches this rule, across all "
                f"{len(self.df):,} rows.")
        set_state(self.lc_ins_summary, state if outcome != "Pass" else "muted")

        frame, value = self.lc_ins_stats["affected records"]
        value.setText("\u2014" if count is None else f"{count:,}")
        set_state(value, state if count else None)
        set_severity(frame, state if count and outcome != "Pass" else None)
        frame, value = self.lc_ins_stats["of this sheet"]
        value.setText("\u2014" if count is None else f"{count / total * 100:.2f}%")
        set_state(value, state if count else None)
        set_severity(frame, state if count and outcome != "Pass" else None)

        self.lc_ins_sections["Why this matters body"].setText(record['why'])
        self.lc_ins_sections["Rule definition body"].setText(record['definition'])
        self.lc_ins_advice.setText(
            record['reason'] if count is None else record['advice'])

        self._populate_logic_evidence(record)
        self.btn_lc_copy.setEnabled(True)

    def _populate_logic_evidence(self, record) -> None:
        """
        A few real rows, or an honest statement that there are none.

        The vocabulary rules carry no row mask - they are about *values*, so
        their evidence is the value list rather than a sample of rows, and
        showing three arbitrary rows instead would suggest a per-row defect
        the check never claimed.
        """
        mask = record.get('mask')
        if record['key'] == 'state_vocabulary':
            vocabulary = record.get('vocabulary') or {}
            suspect = vocabulary.get('suspect_values') or []
            frame = pd.DataFrame([{
                'Value': entry['value'],
                'Count': entry['count'],
                'Share %': entry['pct'],
                'Why': entry.get('reason') or "check",
            } for entry in suspect])
            self._set_logic_evidence(
                frame,
                "Every state value, and how it reads, is in the Data Profile "
                "and the exported report." if frame.empty
                else "The values, not a sample of rows: this rule is about the "
                     "vocabulary, not individual records.",
                "SUSPECT VALUES")
            self.btn_lc_rows.setEnabled(False)
            self.btn_lc_rows.setText("No row set")
            self.btn_lc_rows.setToolTip(
                "This rule reports values rather than rows, so there is no row "
                "set to open")
            return
        if record['key'] == 'unclassified_state':
            frame = pd.DataFrame([{
                'Value': entry.get('value', ''),
                'Rows': entry.get('count', 0),
            } for entry in (record.get('unclassified_values') or [])])
            self._set_logic_evidence(
                frame,
                "Nothing was left unclassified." if frame.empty
                else "The values this tool could not place. Each is counted in "
                     "neither the open nor the closed total.",
                "UNRECOGNISED VALUES")
            self.btn_lc_rows.setEnabled(False)
            self.btn_lc_rows.setText("No row set")
            self.btn_lc_rows.setToolTip(
                "This rule reports values rather than rows, so there is no row "
                "set to open")
            return

        if mask is None or not record['count']:
            self._set_logic_evidence(
                pd.DataFrame(),
                "This rule could not run, so there are no rows to show."
                if record['count'] is None
                else "No row in this sheet matched, so there is nothing to "
                     "show.",
                "ROWS THIS RULE MATCHED")
            self.btn_lc_rows.setEnabled(False)
            self.btn_lc_rows.setText("Open records")
            self.btn_lc_rows.setToolTip("")
            return

        columns = self._logic_evidence_columns(record)
        sample = self.df.loc[mask, columns].head(self.LOGIC_EVIDENCE_ROWS)
        note = (f"All {record['count']:,} matching records."
                if record['count'] <= self.LOGIC_EVIDENCE_ROWS
                else f"First {len(sample):,} of {record['count']:,} matching "
                     f"records, in file order.")
        self._set_logic_evidence(sample.reset_index(drop=True), note,
                                 "ROWS THIS RULE MATCHED")
        self.btn_lc_rows.setEnabled(True)
        self.btn_lc_rows.setText(f"Open {record['count']:,} records")
        self.btn_lc_rows.setToolTip(
            "Every row this rule matched, in its own copyable window")

    def _logic_evidence_columns(self, record) -> List[str]:
        """The ID column, then the fields the rule reads. Nothing else fits."""
        wanted = []
        if self.analyzer.id_column:
            wanted.append(self.analyzer.id_column)
        wanted.extend(record['fields'])
        seen, columns = set(), []
        for name in wanted:
            if name and name in self.df.columns and name not in seen:
                seen.add(name)
                columns.append(name)
        return columns or list(self.df.columns[:3])

    def _set_logic_evidence(self, frame, note: str, title: str = "EVIDENCE") -> None:
        set_table_model(self.lc_evidence_table, PandasModel(frame))
        self.lc_evidence_table.setVisible(not frame.empty)
        self.lc_evidence_title.setText(title)
        self.lc_evidence_note.setText(note)
        set_copy_context(self.lc_evidence_table, **{
            "Rule": self.lc_ins_title.text(),
            "Rows in sheet": f"{len(self.df):,}",
        })

    def _copy_logic_finding(self) -> None:
        """
        The finding as text, with the provenance a ticket needs.

        The tables were copyable while this was not, so the one sentence
        somebody actually wants to paste had to be retyped.
        """
        record = self._logic_selected
        if record is None:
            return
        lines = [f"TicketAudit - Logic Checks: {record['title']}"]
        if self.filepath:
            lines.append(f"Dataset: {os.path.basename(self.filepath)}")
        sheet = self.combo_sheets.currentText()
        if sheet and not self.combo_sheets.isHidden():
            lines.append(f"Sheet: {sheet}")
        lines += [
            f"Fields read: {record['scope_text']}",
            f"Outcome: {record['outcome']}"
            + ("" if record['outcome'] not in ("Fail", "Review")
               else f" ({record['severity']}, {record['confidence'].lower()} "
                    f"evidence)"),
            f"Rows in sheet: {len(self.df):,}",
            "",
            self.lc_ins_summary.text(),
            "",
            f"Why it matters: {record['why']}",
            f"Rule: {record['definition']}",
            f"What to do: {record['reason'] or record['advice']}",
            "",
            f"Generated: {datetime.datetime.now():%Y-%m-%d %H:%M}",
        ]
        QApplication.clipboard().setText("\n".join(lines))
        self.update_status("Finding copied.")

    def _open_logic_rows(self) -> None:
        """
        Every row a rule matched, in a copyable table.

        A dialog rather than a jump to Raw Data: the rows are only meaningful
        beside the rule that selected them, and a filter left behind on another
        view is a trap for whoever opens it next.
        """
        record = self._logic_selected
        if record is None or record.get('mask') is None:
            return
        columns = self._logic_evidence_columns(record)
        rows = self.df.loc[record['mask'], columns]

        dialog = QDialog(self)
        dialog.setWindowTitle(f"{record['title']} - {len(rows):,} records")
        dialog.resize(760, 520)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(SPACE_MD, SPACE_MD, SPACE_MD, SPACE_MD)
        layout.setSpacing(SPACE_SM)

        caption = QLabel(
            f"{len(rows):,} of {len(self.df):,} records in "
            f"{os.path.basename(self.filepath) if self.filepath else 'this sheet'}"
            f", matched on {record['scope_text']}.")
        caption.setWordWrap(True)
        set_state(caption, "muted")
        layout.addWidget(caption)

        table = QTableView()
        configure_table(table, title=f"Logic Checks - {record['title']}",
                        row_numbers=True)
        set_table_model(table, PandasModel(rows))
        set_copy_context(table, **{
            "Rule": record['title'],
            "Matched": f"{len(rows):,} of {len(self.df):,} rows",
        })
        layout.addWidget(table, 1)

        close_row = QHBoxLayout()
        close_row.addStretch()
        close_button = QPushButton("Close")
        close_button.clicked.connect(dialog.accept)
        close_row.addWidget(close_button)
        layout.addLayout(close_row)

        make_text_selectable(dialog)
        dialog.exec()

    def apply_date_format(self):
        """
        Re-read the dates under the chosen field order.

        Rebuilds the analyzer rather than patching its cache: the parsed dates
        feed the column map, the row flags, monthly inflow and the description
        scores, so a half-updated cache would leave views disagreeing about the
        same file. _data_seq is bumped so anything still computing against the
        old parse is discarded, exactly as a sheet switch does.
        """
        if self.df is None:
            return
        dayfirst = self.lc_format_combo.currentData()
        if self.analyzer is not None and self.analyzer.dayfirst == dayfirst:
            return

        # Persisted, because the next sheet or file from the same source will
        # have the same ambiguity and asking again every time is noise.
        self.config["date_dayfirst"] = dayfirst
        self.config_mgr.config = self.config
        self.config_mgr.save()

        self.analyzer = SanityAnalyzer(
            self.df, dayfirst=dayfirst,
            column_overrides=dict(self.config.get("column_overrides") or {}))
        self._data_seq += 1
        self._refresh_inflight.clear()
        self._analyzer_lock = threading.Lock()
        # The findings are about the old reading of the dates.
        self.language_results.clear()

        label = self.lc_format_combo.currentText()
        self.update_status(f"Dates re-read as {label}. All checks recalculated.")
        self.refresh_tab(self.current_view)

    def _populate_date_format(self):
        """Show the field-order finding, and hide the card when there is none."""
        report = self.analyzer.check_date_formats()
        needs_attention = [c for c in report.get('columns', [])
                           if c.get('needs_confirmation') or c.get('conflicting')]

        # Keep the combo in step with the analyzer without re-triggering apply.
        index = self.lc_format_combo.findData(self.analyzer.dayfirst)
        if index >= 0 and index != self.lc_format_combo.currentIndex():
            self.lc_format_combo.blockSignals(True)
            self.lc_format_combo.setCurrentIndex(index)
            self.lc_format_combo.blockSignals(False)

        if not needs_attention:
            self.lc_format_card.setVisible(False)
            return

        lines = []
        severity = "warn"
        for info in needs_attention:
            column = info['column']
            if info.get('conflicting'):
                severity = "error"
                lines.append(
                    f"{column} mixes both field orders: "
                    f"{info['day_first_rows']:,} rows can only be DD/MM and "
                    f"{info['month_first_rows']:,} can only be MM/DD. No single "
                    f"reading is correct — some dates here are wrong whichever "
                    f"is chosen, so this needs fixing at source.")
            else:
                order = ("DD/MM" if self.analyzer._effective_dayfirst(info)
                         else "MM/DD")
                lines.append(
                    f"{column}: {info['ambiguous']:,} of {info['total']:,} "
                    f"values read either way (e.g. "
                    f"{', '.join(info['examples'])}). Currently read as {order}; "
                    f"the other reading changes "
                    f"{info['rows_that_would_change']:,} rows and every "
                    f"resolution time and trend derived from them.")

        self.lc_format_label.setText("\n\n".join(lines))
        set_severity(self.lc_format_card, severity)
        self.lc_format_card.setVisible(True)

    # 4. Language Check
    def create_language_check_tab(self, parent):
        # Scrolled, like Logic Checks, and for a reason worth stating: this view
        # stacks three tables whose heights are pinned to their content with
        # setFixedHeight. A QVBoxLayout given less room than its size hints
        # needs can only take space from the items that are able to shrink - so
        # the section headers were laid out *inside* the unshrinkable detail
        # table, and the tables drew over each other. Scrolling turns that
        # deficit into a scrollbar instead of overlapping widgets.
        outer = QVBoxLayout(parent)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        inner = QWidget()
        outer.addWidget(scroll, 1)
        scroll.setWidget(inner)

        layout = QVBoxLayout(inner)
        layout.setSpacing(SPACE_SM)
        layout.setContentsMargins(SPACE_XS, SPACE_XS, SPACE_XS, SPACE_XS)

        # Controls bar
        controls_card, controls_layout = make_card(horizontal=True, margin=SPACE_SM)
        controls_layout.addWidget(QLabel("Column:"))
        self.lang_col_combo = QComboBox()
        self.lang_col_combo.setMinimumWidth(200)
        controls_layout.addWidget(self.lang_col_combo)

        controls_layout.addWidget(QLabel("Mode:"))
        self.lang_mode_combo = QComboBox()
        # Human labels, with the internal mode id as item data - the raw
        # 'lingua_only'/'rules_only' identifiers gave no clue what they skip.
        for label, mode_id in (
            ("All methods (recommended)", "all"),
            ("Lingua only (ML model)", "lingua_only"),
            ("Rules only (no ML model)", "rules_only"),
        ):
            self.lang_mode_combo.addItem(label, mode_id)
        self.lang_mode_combo.setToolTip(
            "All methods: non-English scripts, then keyword patterns, then the Lingua model.\n"
            "Lingua only: the ML model alone - skips script and keyword detection.\n"
            "Rules only: script and keyword detection alone - no ML model, much faster."
        )
        controls_layout.addWidget(self.lang_mode_combo)

        btn_run = QPushButton("Run Check")
        btn_run.setObjectName("PrimaryButton")
        btn_run.clicked.connect(self.run_language_check)
        controls_layout.addWidget(btn_run)
        controls_layout.addStretch()
        layout.addWidget(controls_card)

        # Tells the reader that running a check here is what puts language into
        # the export. Detection costs ~1.2ms per distinct text, so the report
        # reuses what was run rather than recomputing - which means a column
        # nobody checked has no language in the report at all.
        self.lang_export_hint = QLabel(
            "Results from each column you check here are carried into the "
            "exported report, including a per-ticket language column. Columns "
            "you do not check are reported as “not checked” rather "
            "than assumed to be English.")
        self.lang_export_hint.setWordWrap(True)
        set_state(self.lang_export_hint, "muted")
        layout.addWidget(self.lang_export_hint)

        # Summary card (hidden until a check completes)
        self.lang_summary_card, lang_sum_layout = make_card(margin=SPACE_SM)
        self.lang_summary_card.setVisible(False)
        self.lang_summary_label = QLabel()
        self.lang_summary_label.setWordWrap(True)
        lang_sum_layout.addWidget(self.lang_summary_label)
        layout.addWidget(self.lang_summary_card)

        # Language breakdown table (hidden when all-English)
        self.lang_breakdown_header = QLabel("Language Breakdown")
        self.lang_breakdown_header.setObjectName("SectionHeader")
        self.lang_breakdown_header.setVisible(False)
        layout.addWidget(self.lang_breakdown_header)

        self.lang_breakdown_table = QTableView()
        configure_table(self.lang_breakdown_table, sortable=True, stretch_last=False,
                        title="Language Check — Language Breakdown")
        self.lang_breakdown_table.setVisible(False)
        layout.addWidget(self.lang_breakdown_table)

        # Samples table (hidden when all-English)
        self.lang_samples_header = QLabel("Non-English Samples (top 10 by confidence)")
        self.lang_samples_header.setObjectName("SectionHeader")
        self.lang_samples_header.setVisible(False)
        layout.addWidget(self.lang_samples_header)

        self.lang_samples_table = QTableView()
        configure_table(self.lang_samples_table, sortable=False, stretch_last=True,
                        title="Language Check — Non-English Samples")
        self.lang_samples_table.setVisible(False)
        layout.addWidget(self.lang_samples_table)

        # Methodology, and therefore last. It is shown on every completed check,
        # including an all-English one, so the reader can tell a clean result
        # from a narrowly-scoped one - but it is *how* the answer was reached,
        # and it used to sit between the answer and the evidence for it.
        self.lang_detail_header = QLabel("Detection Details")
        self.lang_detail_header.setObjectName("SectionHeader")
        self.lang_detail_header.setVisible(False)
        layout.addWidget(self.lang_detail_header)

        self.lang_detail_table = QTableView()
        configure_table(self.lang_detail_table, sortable=False, stretch_last=True,
                        title="Language Check — Detection Details")
        self.lang_detail_table.setVisible(False)
        layout.addWidget(self.lang_detail_table)

        # Placeholder shown before first run
        self.lang_placeholder = QLabel("Select a column and click Run Check...")
        set_state(self.lang_placeholder, "muted")
        self.lang_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lang_placeholder, stretch=1)

    def refresh_language_check(self):
        if self.df is not None:
            self.lang_col_combo.clear()
            self.lang_col_combo.addItems(list(self.df.columns))

            # Auto-select description/summary column
            for i in range(self.lang_col_combo.count()):
                text = self.lang_col_combo.itemText(i).lower()
                if "description" in text or "summary" in text:
                    self.lang_col_combo.setCurrentIndex(i)
                    break

    def run_language_check(self):
        col = self.lang_col_combo.currentText()
        if not col or self.df is None or self.analyzer is None:
            return

        self.lang_placeholder.setVisible(False)
        self.lang_summary_card.setVisible(True)
        self.lang_summary_label.setText("Analyzing… please wait.")
        set_severity(self.lang_summary_card, None)
        self.lang_detail_header.setVisible(False)
        self.lang_detail_table.setVisible(False)
        self.lang_breakdown_header.setVisible(False)
        self.lang_breakdown_table.setVisible(False)
        self.lang_samples_header.setVisible(False)
        self.lang_samples_table.setVisible(False)

        mode = self.lang_mode_combo.currentData()

        def run_check(analyzer):
            # analyzer.df, not self.df: the analyzer passed in is pinned to the
            # sheet this run was queued for, so a mid-flight sheet switch cannot
            # redirect the check onto different data.
            return self.lang_checker.check_column(analyzer.df, col, mode=mode)

        def on_result(result):
            # Render directly from the structured dict — no emoji-prefix parsing
            if 'error' in result:
                self.lang_summary_label.setText(result['error'])
                set_severity(self.lang_summary_card, "error")
                return

            if 'warning' in result:
                self.lang_summary_label.setText(result['warning'])
                set_severity(self.lang_summary_card, "warn")
                # The details are the explanation for the warning, so they are
                # more useful here than on a successful run, not less.
                self._populate_language_detail(result)
                self._set_language_copy_context(result)
                # Remembered even though it failed: the report must be able to
                # say "not checked" rather than silently omitting the column.
                self._remember_language_result(result)
                return

            non_english = result['non_english_count']
            pct = result['percentage']
            total = result['total_analyzed']
            original = result['total_original']

            self._populate_language_detail(result)
            self._set_language_copy_context(result)
            self._remember_language_result(result)

            # Name the methods that actually ran, not the ones that are
            # installed - reporting availability made 'rules only' runs claim
            # the ML model had been used.
            scope = result.get('mode_label', result.get('mode', 'unknown'))

            if non_english > 0:
                self.lang_summary_label.setText(
                    f"{non_english:,} of {total:,} analyzed entries ({pct:.1f}%) "
                    f"contain non-English text.\nDetection method: {scope}."
                )
                set_severity(self.lang_summary_card, "error")

                # Language breakdown
                langs = result['languages']
                bm = QStandardItemModel(len(langs), 3)
                bm.setHorizontalHeaderLabels(["Language", "Rows", "% of Analyzed"])
                for i, (lang, count) in enumerate(sorted(langs.items(), key=lambda x: -x[1])):
                    bm.setItem(i, 0, QStandardItem(lang))
                    ci = QStandardItem()
                    ci.setData(count, Qt.ItemDataRole.DisplayRole)
                    bm.setItem(i, 1, ci)
                    pi = QStandardItem()
                    pi.setData(round(count / total * 100, 2) if total else 0.0,
                               Qt.ItemDataRole.DisplayRole)
                    bm.setItem(i, 2, pi)
                self.lang_breakdown_header.setVisible(True)
                self.lang_breakdown_table.setVisible(True)
                set_table_model(self.lang_breakdown_table, bm)
                # Fit to content: usually only a handful of languages, and a
                # fixed height would leave a large empty panel below one row.
                # Capped at 6, past which the table scrolls itself.
                fit_table_height(self.lang_breakdown_table, max_rows=6)

                # Samples. "Found by" is its own column: a character hit is
                # definitive, a keyword hit is a fixed 0.90 heuristic, and only
                # Lingua carries a real probability - the old single
                # "Confidence" column printed all three as if comparable.
                samples = result['samples']
                sm = QStandardItemModel(len(samples), 4)
                sm.setHorizontalHeaderLabels(
                    ["Language", "Found By", "Confidence", "Text"])
                for i, (idx, text, lang, conf, method) in enumerate(samples):
                    sm.setItem(i, 0, QStandardItem(lang))
                    sm.setItem(i, 1, QStandardItem(_LANG_METHOD_LABELS.get(method, method)))
                    sm.setItem(i, 2, QStandardItem(_lang_confidence_text(method, conf)))
                    sm.setItem(i, 3, QStandardItem(text[:120]))
                self.lang_samples_header.setVisible(True)
                self.lang_samples_table.setVisible(True)
                set_table_model(self.lang_samples_table, sm)
            else:
                self.lang_summary_label.setText(
                    f"All {total:,} analyzed entries appear to be in English.\n"
                    f"Detection method: {scope}."
                )
                set_severity(self.lang_summary_card, "ok")

        def on_error(err):
            self.lang_summary_label.setText(f"Error: {err}")
            set_severity(self.lang_summary_card, "error")

        # Key includes the parameters: re-clicking the same column while it runs
        # is a duplicate and is dropped, but picking a different column is a new
        # request and must not be.
        self._refresh_async(f"Language Check:{col}:{mode}", run_check,
                            on_result, err_callback=on_error)

    def _remember_language_result(self, result):
        """
        Keep the finding so Export Report and Generate Report can include it.
        Keyed by column, so checking a second column adds to the report rather
        than replacing what was already found in the first.
        """
        column = result.get('column')
        if column:
            self.language_results[column] = result

    def _set_language_copy_context(self, result):
        """
        Record which column and settings produced this result, so a table or the
        summary sentence copied into a report carries its own denominator and
        scope rather than arriving as a bare list of counts.
        """
        fields = {
            "Column checked": result.get('column'),
            "Detection method": result.get('mode_label'),
            "Rows analyzed": f"{result.get('total_analyzed', 0):,}",
            "Rows skipped (too short)": f"{result.get('skipped_short', 0):,}",
        }
        if 'non_english_count' in result:
            fields["Non-English rows"] = (
                f"{result['non_english_count']:,} ({result.get('percentage', 0)}% of analyzed)"
            )
        # The summary label is included: it is the one sentence most likely to be
        # pasted into a ticket, and "9.7%" means nothing without its column.
        for widget in (self.lang_detail_table, self.lang_breakdown_table,
                       self.lang_samples_table, self.lang_summary_label):
            set_copy_context(widget, **fields)

    def _populate_language_detail(self, result):
        """
        Show what the run actually did: which tiers ran, how many rows were in
        scope, and which tier found each detection.

        Shown for clean results too. "No non-English text found" means something
        quite different depending on whether all three tiers ran or only one, and
        the summary line alone cannot carry that.
        """
        if 'mode' not in result:      # error results carry no run info
            return

        rules_used = result.get('rules_used', False)
        lingua_used = result.get('using_lingua', False)
        lingua_available = result.get('lingua_available', False)
        methods = result.get('methods', {})
        min_len = result.get('min_length', 0)

        if lingua_used:
            lingua_state, lingua_sev = "Used", "ok"
        elif not lingua_available:
            # Only fatal when it was the sole tier requested; in 'rules_only'
            # a missing model is irrelevant to the result.
            lingua_state = "Not installed"
            lingua_sev = "error" if not rules_used else "muted"
        else:
            lingua_state, lingua_sev = "Skipped by mode", "muted"

        rules_state = ("Used", "ok") if rules_used else ("Skipped by mode", "muted")

        rows = [
            ("Detection method", result.get('mode_label', result['mode']), None),
            ("Non-English script matching", rules_state[0], rules_state[1]),
            ("Keyword pattern matching", rules_state[0], rules_state[1]),
            ("Lingua ML model", lingua_state, lingua_sev),
            ("Rows with text in column", f"{result.get('total_original', 0):,}", None),
            ("Rows analyzed", f"{result.get('total_analyzed', 0):,}", None),
            (f"Rows skipped (under {min_len} characters)",
             f"{result.get('skipped_short', 0):,}", None),
        ]

        # Only report tiers that were allowed to run - "Keyword pattern: 0" in
        # Lingua-only mode would read as "tried and found nothing".
        enabled = []
        if rules_used:
            enabled += [LanguageChecker.METHOD_CHARACTERS, LanguageChecker.METHOD_KEYWORDS]
        if lingua_used:
            enabled.append(LanguageChecker.METHOD_LINGUA)
        for method in enabled:
            rows.append((f"Detected by {_LANG_METHOD_LABELS[method]}",
                         f"{methods.get(method, 0):,}", None))

        model = QStandardItemModel(len(rows), 2)
        model.setHorizontalHeaderLabels(["Setting", "Value"])
        tints = {"ok": COLOR_SUCCESS, "error": COLOR_ERROR, "muted": COLOR_TEXT_MUTED}
        for i, (label, value, sev) in enumerate(rows):
            model.setItem(i, 0, QStandardItem(label))
            item = QStandardItem(value)
            if sev in tints:
                item.setForeground(QBrush(QColor(tints[sev])))
            model.setItem(i, 1, item)

        self.lang_detail_header.setVisible(True)
        self.lang_detail_table.setVisible(True)
        set_table_model(self.lang_detail_table, model)
        fit_table_height(self.lang_detail_table)

    # 5. Description Quality
    def create_description_quality_tab(self, parent):
        layout = make_view(parent, scroll=True)

        # make_card rather than a bare QFrame with the Card name: padding has
        # to come from setContentsMargins, because QSS padding does not move a
        # laid-out QFrame's children - so this "card" had none.
        # horizontal=True because these are a row of controls; make_card
        # defaults to a column, which stacked them down half the page.
        controls, h_layout = make_card(horizontal=True, margin=SPACE_SM)

        h_layout.addWidget(QLabel("Min length:"))
        self.desc_min_spin = QSpinBox()
        self.desc_min_spin.setRange(1, 10000)
        self.desc_min_spin.setValue(int(self.config.get("desc_min_length", 20)))
        self.desc_min_spin.setToolTip("Descriptions shorter than this are flagged as too brief to analyze")
        h_layout.addWidget(self.desc_min_spin)

        h_layout.addWidget(QLabel("Max length:"))
        self.desc_max_spin = QSpinBox()
        self.desc_max_spin.setRange(100, 100000)
        self.desc_max_spin.setValue(int(self.config.get("desc_max_length", 5000)))
        self.desc_max_spin.setToolTip("Descriptions longer than this are flagged as unreviewed dumps")
        h_layout.addWidget(self.desc_max_spin)

        btn_run = QPushButton("Re-analyze")
        btn_run.setObjectName("PrimaryButton")
        btn_run.clicked.connect(self.run_description_quality)
        h_layout.addWidget(btn_run)

        h_layout.addStretch()
        layout.addWidget(controls)

        # KPI tiles — populated after each run, cleared on the "analyzing"
        # transient so they never show stale numbers from a previous column set.
        dq_tiles = QHBoxLayout()
        dq_tiles.setSpacing(SPACE_SM)
        self._dq_kpis = {}
        for key, caption in (("cols", "Columns analyzed"),
                              ("score", "Avg quality score"),
                              ("issues", "Fields with issues"),
                              ("groups", "Recurring groups")):
            _frame, lbl = make_kpi(dq_tiles, caption)
            self._dq_kpis[key] = lbl
        dq_tiles.addStretch()
        layout.addLayout(dq_tiles)

        self.desc_status = QLabel("")
        self.desc_status.setWordWrap(True)
        layout.addWidget(self.desc_status)

        splitter = QSplitter(Qt.Orientation.Vertical)

        # Per-column length metrics
        metrics_panel = QWidget()
        metrics_layout = QVBoxLayout(metrics_panel)
        metrics_layout.setContentsMargins(0, 0, 0, 0)
        metrics_caption = QLabel("Field metrics")
        metrics_caption.setObjectName("CardTitle")
        metrics_layout.addWidget(metrics_caption)
        self.desc_metrics_table = QTableView()
        configure_table(self.desc_metrics_table,
                        title="Description Quality — Field Metrics")
        metrics_layout.addWidget(self.desc_metrics_table)
        splitter.addWidget(metrics_panel)

        # Recurring issue groups
        groups_panel = QWidget()
        groups_layout = QVBoxLayout(groups_panel)
        groups_layout.setContentsMargins(0, 0, 0, 0)
        groups_caption = QLabel(
            "Recurring issues — 'Repeated' is identical once ticket IDs, dates "
            "and numbers are ignored; 'Reworded' groups overlapping wording")
        groups_caption.setObjectName("CardTitle")
        groups_caption.setWordWrap(True)
        groups_layout.addWidget(groups_caption)
        self.desc_groups_table = QTableView()
        configure_table(self.desc_groups_table, stretch_last=True,
                        title="Description Quality — Recurring Issues")
        groups_layout.addWidget(self.desc_groups_table)
        splitter.addWidget(groups_panel)

        splitter.setSizes([200, 400])
        layout.addWidget(splitter)

    def run_description_quality(self):
        """Persist the thresholds, then re-run the analysis."""
        self.config["desc_min_length"] = self.desc_min_spin.value()
        self.config["desc_max_length"] = self.desc_max_spin.value()
        self.config_mgr.config = self.config
        self.config_mgr.save()
        self.refresh_description_quality()

    def refresh_description_quality(self):
        if self.df is None or self.analyzer is None:
            return

        min_len = self.desc_min_spin.value()
        max_len = self.desc_max_spin.value()
        self.desc_status.setText("Analyzing descriptions...")
        # Clear any previous run's state, or the last result's red/green would
        # bleed into this transient message.
        set_state(self.desc_status, "busy")
        for lbl in self._dq_kpis.values():
            lbl.setText("—")

        def task(analyzer):
            return analyzer.check_description_quality(min_len, max_len)

        def on_error(err):
            self.desc_status.setText(f"Error: {err}")
            set_state(self.desc_status, "error")

        self._refresh_async("Description Quality", task,
                            self._populate_description_quality,
                            err_callback=on_error)

    def _populate_description_quality(self, result):
        """Fill the metrics and recurring-issue tables from the analysis result."""
        red = QBrush(QColor(COLOR_ERROR))

        if result.get('error'):
            self.desc_status.setText(f"{result['error']}")
            set_state(self.desc_status, "warn")
            self.desc_metrics_table.setModel(QStandardItemModel())
            self.desc_groups_table.setModel(QStandardItemModel())
            return

        # --- Field metrics ---
        columns = result['columns']
        headers = ["Field", "Column", "Quality Score", "Empty", "Too Short",
                   "Too Long", "Avg Length", "Median", "Longest"]
        model = QStandardItemModel(len(columns), len(headers))
        model.setHorizontalHeaderLabels(headers)

        for i, info in enumerate(columns):
            model.setItem(i, 0, QStandardItem(info['category']))
            model.setItem(i, 1, QStandardItem(info['column']))

            score = info.get('overall_score', 0)
            score_item = QStandardItem(f"{score}/100")
            score_item.setData(score, Qt.ItemDataRole.UserRole)
            if score >= 80:
                score_item.setForeground(QBrush(QColor(COLOR_SUCCESS)))
            elif score >= 50:
                score_item.setForeground(QBrush(QColor(COLOR_WARNING)))
            else:
                score_item.setForeground(red)
            model.setItem(i, 2, score_item)

            for col_idx, (count, pct) in enumerate([
                (info['empty'], info['empty_pct']),
                (info['too_short'], info['too_short_pct']),
                (info['too_long'], info['too_long_pct']),
            ], start=3):
                item = QStandardItem(f"{count:,} ({pct}%)")
                if count:
                    item.setForeground(red)
                model.setItem(i, col_idx, item)

            for col_idx, value in enumerate(
                [info['avg_length'], info['median_length'], info['max_length']], start=6
            ):
                item = QStandardItem()
                item.setData(value, Qt.ItemDataRole.DisplayRole)
                model.setItem(i, col_idx, item)

        set_table_model(self.desc_metrics_table, model)

        # --- Recurring issues ---
        rows = []
        for rep in result['repetition']:
            for group in rep['top_groups']:
                rows.append((rep['column'], group))

        group_headers = ["Column", "Type", "Tickets", "% of Rows",
                         "Sample Ticket IDs", "Description"]
        group_model = QStandardItemModel(len(rows), len(group_headers))
        group_model.setHorizontalHeaderLabels(group_headers)

        for i, (col, group) in enumerate(rows):
            group_model.setItem(i, 0, QStandardItem(col))

            kind_item = QStandardItem(
                "Repeated" if group['kind'] == 'exact' else "Reworded"
            )
            if group['kind'] == 'exact':
                kind_item.setForeground(red)
            else:
                kind_item.setForeground(QBrush(QColor(COLOR_WARNING)))
            group_model.setItem(i, 1, kind_item)

            count_item = QStandardItem()
            count_item.setData(group['count'], Qt.ItemDataRole.DisplayRole)
            group_model.setItem(i, 2, count_item)

            pct_item = QStandardItem()
            pct_item.setData(group['percentage'], Qt.ItemDataRole.DisplayRole)
            group_model.setItem(i, 3, pct_item)

            group_model.setItem(i, 4, QStandardItem(", ".join(group['sample_ids'])))
            group_model.setItem(i, 5, QStandardItem(group['text']))

        set_table_model(self.desc_groups_table, group_model)

        # --- KPI tiles ---
        scores = [c.get('overall_score', 0) for c in result.get('columns', [])]
        avg_score = round(sum(scores) / len(scores)) if scores else 0
        fields_with_issues = sum(
            1 for c in result.get('columns', [])
            if c.get('empty', 0) or c.get('too_short', 0) or c.get('too_long', 0)
        )
        total_groups = sum(len(rep.get('top_groups', [])) for rep in result.get('repetition', []))
        self._dq_kpis["cols"].setText(str(len(result.get('columns', []))))
        self._dq_kpis["score"].setText(f"{avg_score}/100")
        self._dq_kpis["issues"].setText(str(fields_with_issues))
        set_state(self._dq_kpis["issues"], "error" if fields_with_issues else "ok")
        self._dq_kpis["groups"].setText(str(total_groups))
        set_state(self._dq_kpis["groups"], "error" if total_groups else "ok")

        # --- Status line ---
        if result['issues']:
            summary = " • ".join(result['issues'])
            self.desc_status.setText(f"{summary}")
            set_state(self.desc_status, "error")
        else:
            self.desc_status.setText("No description quality issues found.")
            set_state(self.desc_status, "ok")

        if any(rep.get('similar_truncated') for rep in result['repetition']):
            self.desc_status.setText(
                self.desc_status.text() +
                f"\nNote: reworded-issue matching scanned the first "
                f"{SanityAnalyzer._SIMILAR_MAX_ROWS:,} populated rows only."
            )

    # 6. Monthly Inflow (QtCharts)
    def create_monthly_inflow_tab(self, parent):
        layout = make_view(parent, scroll=True)
        layout.setSpacing(SPACE_SM)

        # KPI tiles — answer "how much data, over what span, with what peak?"
        # before the reader looks at the chart.
        tiles = QHBoxLayout()
        tiles.setSpacing(SPACE_SM)
        self._monthly_kpis = {}
        for key, caption in (("total", "Total tickets"),
                              ("months", "Months covered"),
                              ("peak", "Peak month"),
                              ("avg", "Avg / month")):
            _frame, lbl = make_kpi(tiles, caption)
            self._monthly_kpis[key] = lbl
        tiles.addStretch()
        layout.addLayout(tiles)

        # Splitter so the breakdown can be grown - a fixed stretch cut the table
        # off at ~4 rows, and a year of data is 12.
        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.setChildrenCollapsible(False)

        self.chart_view = QChartView()
        self.chart_view.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.chart_view.setBackgroundBrush(QBrush(QColor(COLOR_SURFACE_BASE)))
        # QChartView's size hint is tiny, so without a minimum the splitter
        # collapses the chart to a sliver on first show.
        self.chart_view.setMinimumHeight(240)
        splitter.addWidget(self.chart_view)

        # Breakdown table (compact panel below the chart)
        details_card, details_layout = make_card(margin=SPACE_SM)

        hdr_row = QHBoxLayout()
        hdr_lbl = QLabel("Monthly Breakdown")
        hdr_lbl.setObjectName("SectionHeader")
        hdr_row.addWidget(hdr_lbl)
        hdr_row.addStretch()
        self.monthly_total_label = QLabel("")
        set_state(self.monthly_total_label, "muted")
        hdr_row.addWidget(self.monthly_total_label)
        details_layout.addLayout(hdr_row)

        # The chart's total is not the row count whenever a created date is
        # missing, and the gap is invisible from a bar chart. Saying it here is
        # the difference between "5,000 tickets" and "the 4,960 we can date".
        self.monthly_coverage = QLabel("")
        self.monthly_coverage.setObjectName("Caption")
        self.monthly_coverage.setWordWrap(True)
        details_layout.addWidget(self.monthly_coverage)

        self.monthly_table = QTableView()
        configure_table(self.monthly_table, sortable=True, stretch_last=False, title="Monthly Breakdown")
        details_layout.addWidget(self.monthly_table)

        splitter.addWidget(details_card)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        # setSizes seeds the initial split; stretch factors only govern resizing
        splitter.setSizes([460, 300])
        layout.addWidget(splitter)

    def refresh_monthly_inflow(self):
        result = self.analyzer.get_monthly_inflow()
        if 'error' in result:
            self.monthly_total_label.setText(result['error'])
            set_state(self.monthly_total_label, "error")
            chart = QChart()
            theme_chart(chart)
            self.chart_view.setChart(chart)
            set_table_model(self.monthly_table, QStandardItemModel())
            for lbl in self._monthly_kpis.values():
                lbl.setText("—")
            return

        if not result.get('data'):
            self.monthly_total_label.setText("No monthly data available")
            set_state(self.monthly_total_label, "muted")
            for lbl in self._monthly_kpis.values():
                lbl.setText("—")
            return

        categories = list(result['data'].keys())
        values = list(result['data'].values())
        total = sum(values) or 1

        undated = len(self.df) - sum(values)
        self.monthly_coverage.setText(
            f"Every row has a date in {result.get('date_col') or 'the created column'}."
            if undated <= 0 else
            f"{sum(values):,} of {len(self.df):,} rows carry a date in "
            f"{result.get('date_col') or 'the created column'}; {undated:,} have "
            f"none and appear in no month below.")
        set_state(self.monthly_coverage, "warn" if undated > 0 else "muted")

        # KPI tiles
        peak_cat = categories[values.index(max(values))] if values else "—"
        avg_val = round(total / len(categories)) if categories else 0
        self._monthly_kpis["total"].setText(f"{sum(values):,}")
        self._monthly_kpis["months"].setText(str(len(categories)))
        self._monthly_kpis["peak"].setText(peak_cat)
        self._monthly_kpis["avg"].setText(f"{avg_val:,}")

        # Bar chart — use design-token colours, not raw hex literals
        set0 = QBarSet("Inflow")
        set0.setColor(QColor(COLOR_ACCENT))
        set0.setBorderColor(QColor(COLOR_ACCENT))
        set0.setLabelColor(QColor(COLOR_TEXT_MUTED))
        for v in values:
            set0.append(v)

        series = QBarSeries()
        series.append(set0)
        series.setLabelsVisible(True)
        series.setLabelsPosition(QAbstractBarSeries.LabelsPosition.LabelsOutsideEnd)
        # Bars filled the whole category slot, so twelve months read as one
        # solid block. A gap is what makes them countable.
        series.setBarWidth(0.62)

        chart = QChart()
        chart.addSeries(series)
        chart.setAnimationOptions(QChart.AnimationOption.SeriesAnimations)

        axisX = QBarCategoryAxis()
        axisX.append(categories)
        chart.addAxis(axisX, Qt.AlignmentFlag.AlignBottom)
        series.attachAxis(axisX)

        axisY = QValueAxis()
        chart.addAxis(axisY, Qt.AlignmentFlag.AlignLeft)
        series.attachAxis(axisY)

        theme_chart(chart, category_axis=axisX, value_axis=axisY,
                    top_value=max(values) if values else 0)
        self.chart_view.setChart(chart)

        # Breakdown table (replaces the old ASCII text block)
        self.monthly_total_label.setText(f"Total: {total:,}")
        set_state(self.monthly_total_label, "muted")

        model = QStandardItemModel(len(categories), 3)
        model.setHorizontalHeaderLabels(["Month", "Count", "Share %"])
        for i, (cat, val) in enumerate(zip(categories, values)):
            pct = val / total * 100
            model.setItem(i, 0, QStandardItem(cat))
            cnt_item = QStandardItem()
            cnt_item.setData(int(val), Qt.ItemDataRole.DisplayRole)
            model.setItem(i, 1, cnt_item)
            pct_item = QStandardItem()
            pct_item.setData(round(pct, 1), Qt.ItemDataRole.DisplayRole)
            model.setItem(i, 2, pct_item)

        set_table_model(self.monthly_table, model)

    # 7. Pivot Tables
    def create_pivot_tables_tab(self, parent):
        layout_outer = make_view(parent)
        layout = QHBoxLayout()
        layout_outer.addLayout(layout)
        
        # List of pivots
        self.pivot_list = QListWidget()
        self.pivot_list.setFixedWidth(200)
        self.pivot_list.currentRowChanged.connect(self.show_selected_pivot)
        layout.addWidget(self.pivot_list)
        
        # Pivot detail table
        self.pivot_table = QTableView()
        configure_table(self.pivot_table, stretch_last=True, title="Pivot")
        layout.addWidget(self.pivot_table)

    def refresh_pivot_tables(self):
        # Names only. The counts for the selected column are computed on demand
        # in show_selected_pivot and cached by the analyzer, so opening this view
        # no longer pays for 24 pivots it will not display.
        columns = self.analyzer.get_pivot_columns()

        # Repopulating the list re-emits currentRowChanged; block it so the
        # pivot is computed once, by the explicit setCurrentRow below.
        self.pivot_list.blockSignals(True)
        self.pivot_list.clear()
        self.pivot_list.addItems(columns)
        self.pivot_list.blockSignals(False)

        if self.pivot_list.count() > 0:
            self.pivot_list.setCurrentRow(0)

    def show_selected_pivot(self, row):
        if row < 0: return
        col = self.pivot_list.item(row).text()
        data = self.analyzer.get_pivot(col)

        if not data or 'error' in data:
            return

        model = QStandardItemModel(len(data['values']), 3)
        model.setHorizontalHeaderLabels([col, "Count", "Percentage"])
        
        for i, item in enumerate(data['values']):
            model.setItem(i, 0, QStandardItem(str(item['value'])))
            
            cnt = QStandardItem()
            cnt.setData(item['count'], Qt.ItemDataRole.DisplayRole)
            model.setItem(i, 1, cnt)
            
            pct = QStandardItem()
            pct.setData(item['percentage'], Qt.ItemDataRole.DisplayRole)
            model.setItem(i, 2, pct)
            
        set_table_model(self.pivot_table, model)

    # 8. Data Profile
    def create_data_profile_tab(self, parent):
        layout = make_view(parent, scroll=True)
        layout.setSpacing(SPACE_SM)

        # Summary tiles, in the same language as every other view's headline.
        # This row used to be four hand-built caption/value pairs with their own
        # spacing, which is why Data Profile read as a different app.
        tiles = QHBoxLayout()
        tiles.setSpacing(SPACE_SM)
        self._profile_stats = {}
        for key, caption in (("rows", "Rows"), ("cols", "Fields"),
                             ("mem", "Memory (MB)"),
                             ("ws", "Fields with padding")):
            frame, value = make_kpi(tiles, caption)
            self._profile_stats[key] = value
            self._profile_stats[key + "_tile"] = frame
        tiles.addStretch()
        layout.addLayout(tiles)

        self.profile_verdict = QLabel("")
        self.profile_verdict.setWordWrap(True)
        layout.addWidget(self.profile_verdict)

        self.profile_range = QLabel("")
        self.profile_range.setObjectName("Caption")
        self.profile_range.setWordWrap(True)
        layout.addWidget(self.profile_range)

        # Column details table
        hdr = QLabel("Column Details")
        hdr.setObjectName("SectionHeader")
        layout.addWidget(hdr)

        self.profile_table = QTableView()
        configure_table(self.profile_table, sortable=True, stretch_last=True, title="Data Profile")
        layout.addWidget(self.profile_table)

    def refresh_data_profile(self):
        profile = self.analyzer.get_data_profile()

        # Summary tiles
        columns = profile['columns']
        padded = [c for c in columns if c['has_whitespace']]
        self._profile_stats["rows"].setText(f"{profile['total_rows']:,}")
        self._profile_stats["cols"].setText(str(profile['total_columns']))
        self._profile_stats["mem"].setText(f"{profile['memory_mb']:.1f}")
        self._profile_stats["ws"].setText(f"{len(padded):,}")
        set_state(self._profile_stats["ws"], "warn" if padded else None)
        set_severity(self._profile_stats["ws_tile"], "warn" if padded else None)

        # Verdict sentence — plain English before the column table.
        if padded:
            names = ", ".join(c['name'] for c in padded[:3])
            more = f" and {len(padded) - 3} more" if len(padded) > 3 else ""
            self.profile_verdict.setText(
                f"{len(padded)} field{'s' if len(padded) != 1 else ''} contain "
                f"whitespace padding: {names}{more}.")
            set_state(self.profile_verdict, "warn")
        else:
            self.profile_verdict.setText(
                f"All {profile['total_columns']} fields look clean — "
                f"no whitespace padding detected.")
            set_state(self.profile_verdict, "ok")

        span = profile.get('date_range')
        self.profile_range.setText(
            f"Dates run {span['first_date']} to {span['last_date']}."
            if span else "No date column was recognised, so there is no span "
                         "to report.")

        # Column table — has_whitespace is now a proper sortable column
        cols = profile['columns']
        model = QStandardItemModel(len(cols), 4)
        model.setHorizontalHeaderLabels(["Column", "Type", "Whitespace", "Sample"])
        for i, col in enumerate(cols):
            model.setItem(i, 0, QStandardItem(col['name']))
            model.setItem(i, 1, QStandardItem(col['dtype']))
            ws_item = QStandardItem("Yes" if col['has_whitespace'] else "")
            if col['has_whitespace']:
                ws_item.setForeground(QBrush(QColor(COLOR_WARNING)))
            model.setItem(i, 2, ws_item)
            model.setItem(i, 3, QStandardItem(col.get('sample') or ""))

        set_table_model(self.profile_table, model)

    # 9. Duplicate Check
    def create_duplicate_check_tab(self, parent):
        """
        Verdict, then the IDs, then the rows.

        The ID list was a QListWidget of "INC0001: 3 occurrences" strings - two
        columns of tabular data formatted into one, so it could not be sorted,
        could not be copied as columns, and did not carry the copy actions every
        other result in the app has.
        """
        layout = make_view(parent)

        tiles = QHBoxLayout()
        tiles.setSpacing(SPACE_SM)
        self.dup_kpis = {}
        for label in ("Surplus rows", "IDs affected", "Rows involved"):
            frame, value = make_kpi(tiles, label)
            self.dup_kpis[label] = (frame, value)
        tiles.addStretch()
        layout.addLayout(tiles)

        self.dup_info = QLabel("Checking...")
        self.dup_info.setWordWrap(True)
        set_state(self.dup_info, "muted")
        layout.addWidget(self.dup_info)

        # Two empty table frames filling the window is a worse answer than a
        # sentence: when nothing repeats there is nothing to show, and the panel
        # says so instead of framing the absence.
        self.dup_clean_card, clean_layout = make_card(layout, margin=SPACE_MD)
        self.dup_clean_label = QLabel("")
        self.dup_clean_label.setWordWrap(True)
        clean_layout.addWidget(self.dup_clean_label)
        self.dup_clean_card.setVisible(False)

        # Side by side rather than stacked: the ID list is two narrow columns
        # and the row dump is as wide as the sheet, so stacking them gave the
        # narrow one half the height and the wide one half the width it needed.
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(SPACE_XS)
        ids_title = QLabel("REPEATED IDS")
        ids_title.setObjectName("CardTitle")
        left_layout.addWidget(ids_title)
        self.dup_ids_note = QLabel("")
        self.dup_ids_note.setWordWrap(True)
        set_state(self.dup_ids_note, "muted")
        left_layout.addWidget(self.dup_ids_note)
        self.dup_ids_table = QTableView()
        configure_table(self.dup_ids_table, title="Repeated IDs",
                        stretch_last=True)
        left_layout.addWidget(self.dup_ids_table, 1)
        splitter.addWidget(left)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(SPACE_XS)
        rows_title = QLabel("THE ROWS THEMSELVES")
        rows_title.setObjectName("CardTitle")
        right_layout.addWidget(rows_title)
        self.dup_rows_note = QLabel("")
        self.dup_rows_note.setWordWrap(True)
        set_state(self.dup_rows_note, "muted")
        right_layout.addWidget(self.dup_rows_note)
        self.dup_rows_table = QTableView()
        # Row labels identify which source rows are duplicated
        configure_table(self.dup_rows_table, row_numbers=True,
                        title="Duplicate Rows")
        right_layout.addWidget(self.dup_rows_table, 1)
        splitter.addWidget(right)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        self.dup_splitter = splitter
        layout.addWidget(splitter, 1)
        # Takes the height when the splitter is hidden, so the clean-state card
        # stays at the top of the page rather than stretching down it.
        layout.addStretch(0)

    def refresh_duplicate_check(self):
        result = self.analyzer.check_duplicates()

        if result.get('error'):
            self.dup_info.setText(result['error'])
            set_state(self.dup_info, "warn")
            for label in self.dup_kpis:
                frame, value = self.dup_kpis[label]
                value.setText("\u2014")
                set_state(value, None)
                set_severity(frame, None)
            self.dup_splitter.setVisible(False)
            self._show_dup_clean(
                "No column in this file could be identified as a ticket ID, so "
                "nothing was compared for uniqueness. This is not a clean "
                "result - set the ID column in Column Check and the check will "
                "run.", "warn")
            set_table_model(self.dup_ids_table, QStandardItemModel())
            set_table_model(self.dup_rows_table, PandasModel(pd.DataFrame()))
            return

        ids = result['duplicate_ids']
        mask = self.analyzer.duplicate_row_mask()
        involved = int(mask.sum()) if mask is not None else 0

        # Surplus and involved differ on purpose and the analyzer says so: a
        # value appearing twice puts 2 rows in the mask and adds 1 to the
        # surplus. Both are shown because a reader needs the first to know how
        # many rows to delete and the second to know how many to look at.
        for label, count, severity in (("Surplus rows", result['count'], "error"),
                                       ("IDs affected", len(ids), "error"),
                                       ("Rows involved", involved, "warn")):
            frame, value = self.dup_kpis[label]
            value.setText(f"{count:,}")
            set_state(value, severity if count else None)
            set_severity(frame, severity if count else None)

        if not result['count']:
            self.dup_info.setText("")
            self.dup_splitter.setVisible(False)
            self._show_dup_clean(
                f"Every value in {result['id_col']} is unique across all "
                f"{len(self.df):,} rows, so no record appears twice.", "ok")
            set_table_model(self.dup_ids_table, QStandardItemModel())
            set_table_model(self.dup_rows_table, PandasModel(pd.DataFrame()))
            return

        self.dup_splitter.setVisible(True)
        self.dup_clean_card.setVisible(False)

        self.dup_info.setText(
            f"{result['count']:,} surplus rows share a {result['id_col']} value "
            f"with another row. {involved:,} rows are involved across "
            f"{len(ids):,} repeated IDs \u2014 deleting the surplus leaves one row "
            f"per ID.")
        set_state(self.dup_info, "error")

        shown = list(ids.items())[:50]
        model = QStandardItemModel(len(shown), 2)
        model.setHorizontalHeaderLabels([result['id_col'], "Occurrences"])
        for i, (value, count) in enumerate(shown):
            model.setItem(i, 0, QStandardItem(str(value)))
            occurrences = QStandardItem()
            occurrences.setData(int(count), Qt.ItemDataRole.DisplayRole)
            occurrences.setForeground(QBrush(QColor(COLOR_ERROR)))
            model.setItem(i, 1, occurrences)
        set_table_model(self.dup_ids_table, model)
        # The cap is stated rather than left implied: a list that stops at 50
        # with no note reads as "there are 50".
        self.dup_ids_note.setText(
            f"All {len(ids):,}." if len(ids) <= 50
            else f"First 50 of {len(ids):,}. The full list is in the exported "
                 f"report.")

        rows = self.analyzer.get_duplicate_rows()
        if rows is not None and not rows.empty:
            set_table_model(self.dup_rows_table, PandasModel(rows.head(100)))
            self.dup_rows_note.setText(
                f"All {len(rows):,}, in file order." if len(rows) <= 100
                else f"First 100 of {len(rows):,}, in file order. Row numbers "
                     f"are the ones in the sheet.")
        else:
            set_table_model(self.dup_rows_table, PandasModel(pd.DataFrame()))
            self.dup_rows_note.setText("")

        set_copy_context(self.dup_ids_table, **{
            "ID column": result['id_col'],
            "Repeated IDs": f"{len(ids):,}",
            "Surplus rows": f"{result['count']:,}",
            "Rows in sheet": f"{len(self.df):,}",
        })
        set_copy_context(self.dup_rows_table, **{
            "ID column": result['id_col'],
            "Rows involved": f"{involved:,} of {len(self.df):,}",
        })

    def _show_dup_clean(self, message: str, state: str) -> None:
        self.dup_clean_label.setText(message)
        set_state(self.dup_clean_label, state)
        set_severity(self.dup_clean_card, state)
        self.dup_clean_card.setVisible(True)
        self.dup_info.setText("")

    # 10. Raw Data - Virtual Scrolling
    def create_raw_data_tab(self, parent):
        layout = make_view(parent)
        
        self.raw_info = QLabel("Showing first 1000 rows")
        set_state(self.raw_info, "muted")
        layout.addWidget(self.raw_info)

        self.raw_table = QTableView()
        # Interactive sizing: content-based sizing on a wide raw table is slow
        # and rarely what you want for arbitrary source columns.
        configure_table(self.raw_table, row_numbers=True, resize="interactive", title="Raw Data")
        layout.addWidget(self.raw_table)

    def refresh_raw_data(self):
        if self.df is None: return

        self.raw_info.setText(f"Showing first 1,000 of {len(self.df):,} rows")

        # For performance, only show head(1000)
        set_table_model(self.raw_table, PandasModel(self.df.head(1000)))

    # 11. Extract
    def create_extract_tab(self, parent):
        """
        Pull a smaller file out of one Excel cannot work with.

        The count updates as the selection changes rather than on a button, which
        is new for this app - every other view recomputes on a click or on entry.
        It is affordable because planning an extract is ~20ms at 400k rows, and
        it is the point of the feature: you see the size *before* committing,
        which is exactly what Excel cannot tell you without filtering first.

        Scrolled, like Logic Checks and Language Check: this view pins minimum
        heights on two list widgets, and a QVBoxLayout with less room than its
        size hints need takes the space from whatever can shrink - laying the
        cards below out on top of the lists. Scrolling turns that into a
        scrollbar instead.
        """
        outer = QVBoxLayout(parent)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        inner = QWidget()
        outer.addWidget(scroll)
        scroll.setWidget(inner)

        layout = QVBoxLayout(inner)
        layout.setSpacing(SPACE_SM)
        layout.setContentsMargins(SPACE_XS, SPACE_XS, SPACE_XS, SPACE_XS)

        # The subtitle already says what this view writes, so the intro says
        # only the thing it does not: the count is live.
        intro = QLabel(
            "The row and cell counts update as you change the selection, so you "
            "can see whether Excel will cope before writing anything."
        )
        intro.setWordWrap(True)
        set_state(intro, "muted")
        layout.addWidget(intro)

        # --- Rows ---------------------------------------------------------
        rows_card, rows_layout = make_card(margin=SPACE_SM)
        rows_title = QLabel("Rows")
        rows_title.setObjectName("CardTitle")
        rows_layout.addWidget(rows_title)

        picker = QHBoxLayout()
        picker.addWidget(QLabel("Date column:"))
        self.extract_date_combo = QComboBox()
        self.extract_date_combo.setMinimumWidth(160)
        picker.addWidget(self.extract_date_combo)

        picker.addWidget(QLabel("Months:"))
        self.extract_month_mode_combo = QComboBox()
        for text, mode in (("A continuous range", EXTRACT_MONTHS_RANGE),
                           ("Pick individually", EXTRACT_MONTHS_PICK)):
            self.extract_month_mode_combo.addItem(text, mode)
        self.extract_month_mode_combo.setToolTip(
            "A range takes every month between two points.\n"
            "Picking individually takes only the months you tick - "
            "February, May and July is not a range.")
        picker.addWidget(self.extract_month_mode_combo)
        picker.addStretch()
        rows_layout.addLayout(picker)

        # -- range mode: two combos ------------------------------------------
        self.extract_range_row = QWidget()
        range_layout = QHBoxLayout(self.extract_range_row)
        range_layout.setContentsMargins(0, 0, 0, 0)
        range_layout.setSpacing(SPACE_SM)
        range_layout.addWidget(QLabel("From:"))
        self.extract_from_combo = QComboBox()
        self.extract_from_combo.setMinimumWidth(150)
        range_layout.addWidget(self.extract_from_combo)
        range_layout.addWidget(QLabel("To:"))
        self.extract_to_combo = QComboBox()
        self.extract_to_combo.setMinimumWidth(150)
        range_layout.addWidget(self.extract_to_combo)
        # Qt shows 10 items by default, so a two- or three-year export hides most
        # of its months behind a scroll and the list reads as if it were short.
        for combo in (self.extract_from_combo, self.extract_to_combo):
            combo.setMaxVisibleItems(EXTRACT_MONTHS_VISIBLE)
        range_layout.addStretch()
        rows_layout.addWidget(self.extract_range_row)

        # -- pick mode: a checkable month list -------------------------------
        # The same shape as the column list, because it is the same kind of
        # choice: an arbitrary subset, not two endpoints.
        self.extract_months_panel = QWidget()
        months_layout = QVBoxLayout(self.extract_months_panel)
        months_layout.setContentsMargins(0, 0, 0, 0)
        months_layout.setSpacing(SPACE_XS)
        presets = QHBoxLayout()
        presets.setSpacing(SPACE_SM)
        for text, handler, tip in (
                ("All", self._extract_months_all, "Tick every month"),
                ("None", self._extract_months_none, "Untick every month"),
                ("Whole months only", self._extract_months_complete,
                 "Tick only the months this file covers end to end")):
            button = QPushButton(text)
            button.setToolTip(tip)
            button.clicked.connect(handler)
            presets.addWidget(button)
        presets.addStretch()
        months_layout.addLayout(presets)

        self.extract_month_list = QListWidget()
        self.extract_month_list.setMinimumHeight(150)
        months_layout.addWidget(self.extract_month_list)
        self.extract_months_panel.setVisible(False)
        rows_layout.addWidget(self.extract_months_panel)

        self.extract_undated_check = QCheckBox("Also include rows with no date")
        rows_layout.addWidget(self.extract_undated_check)

        # A cancelled ticket was never worked, so it flatters resolution times
        # and pads volume. The label names the count and the values it matched,
        # because "exclude cancelled" is a claim about someone else's vocabulary
        # and the only way to check it is to see which values it caught.
        self.extract_cancelled_check = QCheckBox("Exclude cancelled tickets")
        self.extract_cancelled_check.setVisible(False)
        rows_layout.addWidget(self.extract_cancelled_check)

        # Shown only when the month buckets rest on a guessed date format, since
        # in that case they are simply wrong.
        self.extract_format_warning = QLabel()
        self.extract_format_warning.setWordWrap(True)
        self.extract_format_warning.setVisible(False)
        rows_layout.addWidget(self.extract_format_warning)

        # A month the file only partly covers. Its row count is real but not
        # comparable with a whole month, which is the thing people carry into a
        # trend without noticing - so it is named, priced in days, and the
        # boundary cases come with a button that drops them.
        self.extract_partial_row = QWidget()
        partial_layout = QHBoxLayout(self.extract_partial_row)
        partial_layout.setContentsMargins(0, 0, 0, 0)
        partial_layout.setSpacing(SPACE_SM)
        self.extract_partial_warning = QLabel()
        self.extract_partial_warning.setWordWrap(True)
        partial_layout.addWidget(self.extract_partial_warning, stretch=1)
        self.btn_extract_trim = QPushButton("Exclude")
        self.btn_extract_trim.setToolTip(
            "Move the range past the partial months at its ends")
        self.btn_extract_trim.clicked.connect(self._extract_trim_partial_months)
        partial_layout.addWidget(self.btn_extract_trim)
        self.extract_partial_row.setVisible(False)
        rows_layout.addWidget(self.extract_partial_row)
        # Beside the Columns card, this one is the shorter of the two, and
        # without a trailing stretch its handful of controls would spread down
        # the card to fill the difference.
        rows_layout.addStretch()

        # --- Columns ------------------------------------------------------
        cols_card, cols_layout = make_card(margin=SPACE_SM)
        self.extract_cols_title = QLabel("Columns")
        self.extract_cols_title.setObjectName("CardTitle")
        cols_layout.addWidget(self.extract_cols_title)

        # A filter, because 105 columns is a realistic width for a
        # ServiceNow export and scrolling that list to find nine of them is the
        # friction this view exists to remove.
        self.extract_col_filter = QLineEdit()
        self.extract_col_filter.setPlaceholderText(
            "Filter columns… (the presets below act on what is shown)")
        self.extract_col_filter.setClearButtonEnabled(True)
        self.extract_col_filter.textChanged.connect(self._extract_filter_columns)
        cols_layout.addWidget(self.extract_col_filter)

        presets = QHBoxLayout()
        for label, handler in (("Standard", self._extract_select_standard),
                               ("All", self._extract_select_all),
                               ("None", self._extract_select_none),
                               ("Non-empty", self._extract_select_non_empty)):
            button = QPushButton(label)
            button.clicked.connect(handler)
            presets.addWidget(button)
        presets.addStretch()
        cols_layout.addLayout(presets)

        self.extract_col_list = QListWidget()
        self.extract_col_list.setMinimumHeight(180)
        cols_layout.addWidget(self.extract_col_list, 1)

        # Side by side, not stacked. Two independent choices of similar size -
        # which rows, which columns - and stacking them spent 316px of height
        # to leave 70% of the width empty. A plain QHBoxLayout rather than a
        # splitter: at 1280x720 the content column is 993px and these two need
        # 439 and 291, so there is nothing for the user to rebalance, and a
        # splitter inside a scroll area is one more thing that can collapse.
        choices = QHBoxLayout()
        choices.setSpacing(SPACE_SM)
        # Rows is top-aligned so it ends where its content ends. A QHBoxLayout
        # gives both children the row's full height, and the Columns card is
        # ~140px taller because its list grows - which left that much empty
        # white *inside* the Rows card. Page background below a shorter card
        # reads as layout; the same gap inside its border reads as a card that
        # failed to fill.
        choices.addWidget(rows_card, 1, Qt.AlignmentFlag.AlignTop)
        choices.addWidget(cols_card, 1)
        layout.addLayout(choices, 1)

        # --- Preview ------------------------------------------------------
        # The first rows of the *actual* subset, built by the same call the
        # write uses. Without it the only way to find a wrong column or a
        # mis-read date is to wait out the write and open the file.
        preview_card, preview_layout = make_card(margin=SPACE_SM)
        self.extract_preview_title = QLabel("Preview")
        self.extract_preview_title.setObjectName("CardTitle")
        preview_layout.addWidget(self.extract_preview_title)
        self.extract_preview_table = QTableView()
        configure_table(self.extract_preview_table, sortable=False,
                        title="Extract — preview")
        preview_layout.addWidget(self.extract_preview_table)
        layout.addWidget(preview_card)

        # --- Result -------------------------------------------------------
        self.extract_result_card, result_layout = make_card(margin=SPACE_SM)
        result_title = QLabel("Result")
        result_title.setObjectName("CardTitle")
        result_layout.addWidget(result_title)

        self.extract_summary = QLabel()
        self.extract_summary.setWordWrap(True)
        result_layout.addWidget(self.extract_summary)

        actions = QHBoxLayout()
        actions.addWidget(QLabel("Format:"))
        self.extract_format_combo = QComboBox()
        self.extract_format_combo.addItem("Excel (.xlsx)", extract_core.FORMAT_XLSX)
        self.extract_format_combo.addItem("CSV (much faster)",
                                          extract_core.FORMAT_CSV)
        actions.addWidget(self.extract_format_combo)
        actions.addStretch()
        self.btn_extract_reveal = QPushButton("Open folder")
        self.btn_extract_reveal.setToolTip("Show the last extract in Explorer")
        self.btn_extract_reveal.clicked.connect(self._reveal_extract)
        self.btn_extract_reveal.setVisible(False)
        actions.addWidget(self.btn_extract_reveal)

        self.btn_extract_cancel = QPushButton("Cancel")
        self.btn_extract_cancel.setObjectName("DangerButton")
        self.btn_extract_cancel.setToolTip(
            "Stop the write. The part-written file is deleted rather than left "
            "looking complete.")
        self.btn_extract_cancel.clicked.connect(self._cancel_extract)
        self.btn_extract_cancel.setVisible(False)
        actions.addWidget(self.btn_extract_cancel)

        self.btn_extract = QPushButton("Extract…")
        self.btn_extract.setObjectName("PrimaryButton")
        self.btn_extract.clicked.connect(self.run_extract)
        actions.addWidget(self.btn_extract)
        result_layout.addLayout(actions)
        # Outside the scroll area, so the write button and the size it is about
        # to write stay on screen while the pickers scroll. They are the point
        # of the view and used to sit below the fold on any window short enough
        # to need a scrollbar at all - which is every window, given the column
        # list.
        action_host = QWidget()
        action_layout = QVBoxLayout(action_host)
        action_layout.setContentsMargins(SPACE_XS, SPACE_XS, SPACE_XS, SPACE_XS)
        action_layout.setSpacing(0)
        action_layout.addWidget(self.extract_result_card)
        outer.addWidget(action_host)

    # -- population -------------------------------------------------------

    def refresh_extract(self):
        """Repopulate the pickers for the current sheet, restoring last choices."""
        if self.df is None or self.analyzer is None:
            return

        # Blocked while repopulating: every one of these emits a change signal,
        # and the plan would be recomputed once per widget for no reason.
        for widget in (self.extract_date_combo, self.extract_from_combo,
                       self.extract_to_combo, self.extract_col_list,
                       self.extract_undated_check, self.extract_format_combo,
                       self.extract_month_list, self.extract_month_mode_combo,
                       self.extract_cancelled_check):
            widget.blockSignals(True)
        try:
            self._populate_extract_date_columns()
            self._populate_extract_months()
            self._populate_extract_columns()
            self._populate_extract_cancelled()
            saved_format = self.config.get("extract_format",
                                           extract_core.FORMAT_XLSX)
            index = self.extract_format_combo.findData(saved_format)
            if index >= 0:
                self.extract_format_combo.setCurrentIndex(index)
            saved_mode = self.config.get("extract_month_mode",
                                         EXTRACT_MONTHS_RANGE)
            index = self.extract_month_mode_combo.findData(saved_mode)
            if index >= 0:
                self.extract_month_mode_combo.setCurrentIndex(index)
        finally:
            for widget in (self.extract_date_combo, self.extract_from_combo,
                           self.extract_to_combo, self.extract_col_list,
                           self.extract_undated_check,
                           self.extract_format_combo,
                           self.extract_month_list,
                           self.extract_month_mode_combo,
                           self.extract_cancelled_check):
                widget.blockSignals(False)

        self._connect_extract_signals()
        self._update_extract_plan()

    def _connect_extract_signals(self):
        """Wire the live recount once, not on every refresh."""
        if getattr(self, "_extract_signals_wired", False):
            return
        self.extract_date_combo.currentIndexChanged.connect(
            self._on_extract_date_column_changed)
        self.extract_from_combo.currentIndexChanged.connect(
            self._update_extract_plan)
        self.extract_to_combo.currentIndexChanged.connect(
            self._update_extract_plan)
        self.extract_undated_check.stateChanged.connect(
            self._update_extract_plan)
        self.extract_cancelled_check.stateChanged.connect(
            self._update_extract_plan)
        self.extract_col_list.itemChanged.connect(self._update_extract_plan)
        self.extract_month_list.itemChanged.connect(self._update_extract_plan)
        self.extract_month_mode_combo.currentIndexChanged.connect(
            self._on_extract_month_mode_changed)
        self.extract_format_combo.currentIndexChanged.connect(
            self._update_extract_plan)
        self._extract_signals_wired = True

    def _extract_date_columns(self) -> List[str]:
        """
        Every column this view can filter by, detected pair first.

        Not just created and closed: a ServiceNow export carries resolved_at,
        sys_updated_on, due_date and more, and "the tickets updated in March" is
        as reasonable a request as "the tickets opened in March". The detected
        pair leads because it is what every other view means by a date, and the
        rest follow in file order.

        A column already typed as datetime by the reader needs no parsing. One
        still held as text is only offered when the analyzer recognised it, since
        that is the path that parses it with the user's confirmed field order -
        guessing at an arbitrary text column would put unverified dates behind a
        picker that looks just as authoritative as the rest.
        """
        analyzer = self.analyzer
        ordered = [c for c in (analyzer.created_column, analyzer.closed_column)
                   if c]
        for column in self.df.columns:
            if column in ordered:
                continue
            if pd.api.types.is_datetime64_any_dtype(self.df[column]):
                ordered.append(column)
        return ordered

    def _populate_extract_date_columns(self):
        """The date columns this file offers, plus a no-filter escape."""
        self.extract_date_combo.clear()
        for column in self._extract_date_columns():
            if self.extract_date_combo.findData(column) < 0:
                self.extract_date_combo.addItem(str(column), column)
        # An extract can legitimately be columns-only: on a wide file the column
        # reduction does more for Excel than the date window does.
        self.extract_date_combo.addItem("(no date filter)", None)

        saved = self.config.get("extract_date_column")
        index = self.extract_date_combo.findData(saved) if saved else 0
        self.extract_date_combo.setCurrentIndex(max(index, 0))

    def _cancelled_state_info(self):
        """
        What the analyzer found for cancelled states on this sheet, cached.

        `mask` is None when the file has no state column, which is different
        from an empty mask: one means nothing to exclude, the other means no way
        to tell, and the control is only offered in the first case.
        """
        cached = getattr(self, "_extract_cancelled_info", None)
        if cached is None and self.analyzer is not None:
            try:
                cached = self.analyzer.cancelled_states()
            except Exception:
                log.warning("could not identify cancelled states",
                            exc_info=True)
                cached = {"state_col": None, "values": [], "rows": 0,
                          "mask": None}
            self._extract_cancelled_info = cached
        return cached or {"state_col": None, "values": [], "rows": 0,
                          "mask": None}

    def _populate_extract_cancelled(self):
        """
        Offer the exclusion only when there is something to exclude.

        Hidden when the file has no state column, and when no value in it looks
        cancelled - a ticked box that removes nothing, or an unticked one next
        to a file it cannot apply to, both invite the reader to believe rows are
        being dropped.
        """
        info = self._cancelled_state_info()
        offer = bool(info["mask"] is not None and info["rows"])
        if offer:
            shown = ", ".join(info["values"][:4])
            if len(info["values"]) > 4:
                shown += f" and {len(info['values']) - 4} more"
            # "in this file", because the number that actually leaves the
            # extract is smaller - some of these sit outside the chosen months.
            # The Result card gives that one. Two counts with their scopes
            # stated beats one count that is wrong in one of the two places.
            self.extract_cancelled_check.setText(
                f"Exclude cancelled tickets — {info['rows']:,} in this file "
                f"({shown})")
            self.extract_cancelled_check.setToolTip(
                f"Matched on the word \"cancel\" in {info['state_col']}. "
                f"Values found: {', '.join(info['values'])}.\n"
                f"Anything else - withdrawn, rejected, abandoned - is kept, "
                f"because whether those mean cancelled is a judgement about "
                f"your vocabulary rather than something this tool can know.")
            saved = bool(self.config.get("extract_exclude_cancelled", False))
            self.extract_cancelled_check.setChecked(saved)
        self.extract_cancelled_check.setVisible(offer)
        if not offer:
            self.extract_cancelled_check.setChecked(False)

    def _extract_exclusion_mask(self):
        """The rows the current options drop, or None when nothing is dropped."""
        if not self.extract_cancelled_check.isChecked():
            return None
        return self._cancelled_state_info()["mask"]

    def _extract_dates(self):
        """
        The parsed series for the chosen column, or None.

        The analyzer's own two go through it, so they honour the confirmed
        day-first setting and its cache. Any other column is already datetime -
        _extract_date_columns only offers those - so it is used as it stands.
        """
        column = self.extract_date_combo.currentData()
        if not column or self.analyzer is None or self.df is None:
            return None
        if column == self.analyzer.created_column:
            return self.analyzer._get_created_dates()
        if column == self.analyzer.closed_column:
            return self.analyzer._get_closed_dates()
        if column in self.df.columns and pd.api.types.is_datetime64_any_dtype(
                self.df[column]):
            return self.df[column]
        return None

    def _populate_extract_months(self):
        dates = self._extract_dates()
        months = extract_core.month_counts(dates)
        # Coverage is a property of the column, not of the range, so it is
        # computed here rather than in the live count - 48ms at 400k rows is
        # affordable once per column change but not on every keystroke.
        self._extract_coverage = extract_core.month_coverage(dates)
        incomplete = {c.period: c for c in self._extract_coverage if c.partial}

        for combo in (self.extract_from_combo, self.extract_to_combo):
            combo.clear()
            for period, count in months:
                label = extract_core.month_label(period)
                partial = incomplete.get(period)
                # The mark is display only - the selection is read from item
                # data, and nothing branches on this text.
                combo.addItem(f"{label}  ({count:,})"
                              + ("  · part month" if partial else ""),
                              period)
                index = combo.count() - 1
                combo.setItemData(
                    index,
                    period if partial is None
                    else f"{period} - {partial.describe()}",
                    Qt.ItemDataRole.ToolTipRole)
        # The tick list carries the same labels, counts and part-month marks.
        # A hand-made month selection wins over "all", for the reason the column
        # ticks do: refresh_* runs on every visit, and restoring the default here
        # would undo ticks the user made on the way past another view. It is
        # dropped on a sheet load, where it no longer describes the data.
        session = getattr(self, "_extract_month_ticks", None)
        available = {period for period, _count in months}
        keep = (set(session) if session and set(session) <= available
                else available)

        self.extract_month_list.clear()
        for period, count in months:
            partial = incomplete.get(period)
            item = QListWidgetItem(
                f"{extract_core.month_label(period)}  ({count:,})"
                + ("   \u26a0 part month" if partial else ""))
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setData(Qt.ItemDataRole.UserRole, period)
            item.setCheckState(Qt.CheckState.Checked if period in keep
                               else Qt.CheckState.Unchecked)
            if partial:
                item.setToolTip(partial.describe())
            self.extract_month_list.addItem(item)

        if months:
            self.extract_from_combo.setCurrentIndex(0)
            self.extract_to_combo.setCurrentIndex(len(months) - 1)

        enabled = bool(months)
        self.extract_from_combo.setEnabled(enabled)
        self.extract_to_combo.setEnabled(enabled)
        self.extract_month_mode_combo.setEnabled(enabled)
        self._extract_apply_month_mode()
        # Hidden rather than greyed with no date filter: every row is included,
        # so a visible "also include undated" box implies rows are being dropped
        # when none are.
        self.extract_undated_check.setEnabled(enabled)
        self.extract_undated_check.setVisible(enabled)

    def _extract_month_mode(self) -> str:
        """Which month control is in force. Read from item data, never text."""
        return (self.extract_month_mode_combo.currentData()
                or EXTRACT_MONTHS_RANGE)

    def _extract_apply_month_mode(self) -> None:
        """Show the control that is in force, and hide the other one."""
        picking = self._extract_month_mode() == EXTRACT_MONTHS_PICK
        has_months = self.extract_month_list.count() > 0
        self.extract_range_row.setVisible(not picking and has_months)
        self.extract_months_panel.setVisible(picking and has_months)

    def _extract_selected_months(self):
        """The ticked months, ascending, as ISO periods."""
        return [self.extract_month_list.item(i).data(Qt.ItemDataRole.UserRole)
                for i in range(self.extract_month_list.count())
                if self.extract_month_list.item(i).checkState()
                == Qt.CheckState.Checked]

    def _extract_set_months(self, wanted) -> None:
        """
        Tick exactly `wanted`, recounting once.

        Blocked while setting, for the reason the column presets are: one signal
        per item would replan three dozen times for a three-year file.
        """
        keep = set(wanted)
        self.extract_month_list.blockSignals(True)
        try:
            for i in range(self.extract_month_list.count()):
                item = self.extract_month_list.item(i)
                item.setCheckState(
                    Qt.CheckState.Checked
                    if item.data(Qt.ItemDataRole.UserRole) in keep
                    else Qt.CheckState.Unchecked)
        finally:
            self.extract_month_list.blockSignals(False)
        self._update_extract_plan()

    def _extract_all_months(self):
        return [self.extract_month_list.item(i).data(Qt.ItemDataRole.UserRole)
                for i in range(self.extract_month_list.count())]

    def _extract_months_all(self):
        self._extract_set_months(self._extract_all_months())

    def _extract_months_none(self):
        self._extract_set_months([])

    def _extract_months_complete(self):
        """
        Only the months this file covers end to end.

        The natural companion to the part-month warning: on an export pulled
        mid-month this is the selection that makes the months comparable, and
        unlike the range trim it can drop a gap in the middle too.
        """
        partial = {c.period for c in self._extract_coverage or [] if c.partial}
        self._extract_set_months([m for m in self._extract_all_months()
                                  if m not in partial])

    def _on_extract_month_mode_changed(self):
        """
        Carry the current range into the tick list when switching to it.

        Ticking everything instead would silently widen the selection - a
        narrowed Jun..Sep became the whole file - which is the opposite of what
        reaching for a finer control means.
        """
        if self._extract_month_mode() == EXTRACT_MONTHS_PICK:
            start = self.extract_from_combo.currentData()
            end = self.extract_to_combo.currentData()
            if start and end:
                inside = [m for m in self._extract_all_months()
                          if start <= m <= end]
                if inside and inside != self._extract_selected_months():
                    self._extract_apply_month_mode()
                    self._extract_set_months(inside)
                    return
        self._extract_apply_month_mode()
        self._update_extract_plan()

    def _populate_extract_columns(self):
        """
        Every source column, checkable, starting from last time's selection.

        Falls back to the detected standard columns on a first run, which is a
        sensible starting point rather than a guess: they are the ones every
        other check in the app already relies on.

        A selection made by hand this session wins over the remembered one, or
        stepping into another view and back would silently undo 25 clicks. It is
        dropped on a sheet load, since it described the previous sheet.
        """
        session = getattr(self, "_extract_ticks", None)
        if session and all(c in self.df.columns for c in session):
            wanted = session
        else:
            remembered = self.config.get("extract_columns") or []
            wanted = [c for c in remembered if c in self.df.columns]
            if not wanted:
                wanted = self._extract_standard_columns()

        fill = self._extract_column_fill()
        self.extract_col_list.clear()
        for column in self.df.columns:
            share = fill.get(str(column))
            # The fill share is the guide to what is worth keeping: on a wide
            # export the columns nobody fills are exactly the ones to drop.
            suffix = "" if share is None else (
                "   (empty)" if share == 0
                else f"   ({share * 100:.0f}% filled)")
            item = QListWidgetItem(f"{column}{suffix}")
            # The column name lives in item data. The label carries the fill
            # share, so reading the name off the text would return "state
            # (100% filled)" and match nothing.
            item.setData(Qt.ItemDataRole.UserRole, str(column))
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if column in wanted
                               else Qt.CheckState.Unchecked)
            self.extract_col_list.addItem(item)
        self._extract_filter_columns(self.extract_col_filter.text())

    def _extract_standard_columns(self):
        """The columns the analyzer matched to a required category."""
        if self.analyzer is None:
            return list(self.df.columns)
        chosen = [info['chosen'] for info in self.analyzer.column_choices().values()
                  if info['chosen']]
        return chosen or list(self.df.columns)

    def _extract_column_fill(self):
        """
        Per-column fill shares, computed once per sheet.

        Reuses the analyzer's cached df.isnull().sum() when the null view has
        already built it - 191ms against 1ms at 80k x 105 - and caches the
        result either way, since refresh_extract runs on every visit.
        """
        cached = getattr(self, "_extract_fill", None)
        if cached is not None:
            return cached
        null_counts = None
        if self.analyzer is not None:
            null_counts = self.analyzer._cache.get("null_counts")
        try:
            fill = extract_core.column_fill(self.df, null_counts)
        except Exception:
            log.warning("could not measure column fill", exc_info=True)
            fill = {}
        self._extract_fill = fill
        return fill

    def _extract_column_name(self, item):
        """The column an item stands for, never its label."""
        return item.data(Qt.ItemDataRole.UserRole) or item.text()

    def _extract_visible_column_items(self):
        """
        The items the filter is currently showing.

        The presets act on these, which is the one rule that makes a filter
        useful: type "desc", press All, get the description columns. With no
        filter every item is visible, so it is the same rule either way.
        """
        return [self.extract_col_list.item(i)
                for i in range(self.extract_col_list.count())
                if not self.extract_col_list.item(i).isHidden()]

    def _extract_filter_columns(self, text):
        """Show only the columns matching `text`, ticks untouched."""
        needle = (text or "").strip().lower()
        shown = 0
        for i in range(self.extract_col_list.count()):
            item = self.extract_col_list.item(i)
            match = needle in self._extract_column_name(item).lower()
            item.setHidden(not match)
            shown += bool(match)
        if needle:
            # Said out loud, because a filtered list plus a preset that acts on
            # it could otherwise look like the preset had missed columns.
            self.extract_col_filter.setToolTip(
                f"{shown} of {self.extract_col_list.count()} columns shown")
        else:
            self.extract_col_filter.setToolTip("")

    def _extract_selected_columns(self):
        """Every ticked column, including ones the filter is hiding."""
        return [self._extract_column_name(self.extract_col_list.item(i))
                for i in range(self.extract_col_list.count())
                if self.extract_col_list.item(i).checkState()
                == Qt.CheckState.Checked]

    # -- presets ----------------------------------------------------------

    def _extract_set_checked(self, wanted):
        """
        Tick exactly `wanted` among the columns currently on show.

        Blocking is not cosmetic: itemChanged fires per item, so a 105-column
        'All' would otherwise replan 105 times.

        Filtered-out columns keep their ticks. That is what makes the filter
        worth having - narrow to "desc", press All, and the nine columns you had
        already chosen elsewhere are still chosen. With no filter every column is
        on show, so it is one rule rather than two.
        """
        keep = set(wanted)
        items = self._extract_visible_column_items()
        self.extract_col_list.blockSignals(True)
        try:
            for item in items:
                item.setCheckState(
                    Qt.CheckState.Checked
                    if self._extract_column_name(item) in keep
                    else Qt.CheckState.Unchecked)
        finally:
            self.extract_col_list.blockSignals(False)
        self._update_extract_plan()

    def _extract_select_standard(self):
        self._extract_set_checked(set(self._extract_standard_columns()))

    def _extract_select_all(self):
        self._extract_set_checked(set(self.df.columns))

    def _extract_select_none(self):
        self._extract_set_checked(set())

    def _extract_select_non_empty(self):
        """
        Every column that holds at least one value.

        The reduction that matters on a wide export: a ServiceNow dump carries
        columns nobody ever fills, and they cost cells in the output for nothing.
        Blank counts the empty string as well as null - see extract.column_fill -
        because a sheet full of "" is not data.
        """
        fill = self._extract_column_fill()
        self._extract_set_checked(
            {column for column in self.df.columns
             if fill.get(str(column), 1.0) > 0})

    def _on_extract_date_column_changed(self):
        """A different date column means different months."""
        self.extract_from_combo.blockSignals(True)
        self.extract_to_combo.blockSignals(True)
        try:
            self._populate_extract_months()
        finally:
            self.extract_from_combo.blockSignals(False)
            self.extract_to_combo.blockSignals(False)
        self._update_extract_plan()

    # -- the live count ---------------------------------------------------

    def _update_extract_plan(self):
        """
        Recompute the size of the current selection. Synchronous on purpose:
        ~20ms at 400k rows, and an async count would race the next click.
        """
        if self.df is None or self.analyzer is None:
            return

        columns = self._extract_selected_columns()
        if columns:
            # Remembered for the rest of the session, so leaving the view and
            # coming back does not discard it. See _populate_extract_columns.
            self._extract_ticks = columns
        dates = self._extract_dates()
        picking = (dates is not None
                   and self._extract_month_mode() == EXTRACT_MONTHS_PICK)
        months = self._extract_selected_months() if picking else None
        if picking and months:
            self._extract_month_ticks = months
        start = self.extract_from_combo.currentData()
        end = self.extract_to_combo.currentData()

        self._extract_show_format_warning()
        self._extract_show_partial_months(start, end, months, dates is not None)

        # Both labels are updated before any early return, or they keep saying
        # whatever the last valid selection said.
        self.extract_cols_title.setText(
            f"Columns — {len(columns)} of {len(self.df.columns)}")
        if dates is None:
            self.extract_undated_check.setText(
                "Also include rows with no date")

        # Each of these bails out before a plan exists, so the preview has to be
        # emptied here too: leaving the last good one on screen would show rows
        # for a selection that no longer produces any.
        if not columns:
            self.extract_summary.setText(
                "No columns selected — tick at least one.")
            set_severity(self.extract_result_card, "warn")
            self.btn_extract.setEnabled(False)
            self._clear_extract_preview()
            return

        if picking and not months:
            self.extract_summary.setText(
                "No months ticked - tick at least one.")
            set_severity(self.extract_result_card, "warn")
            self.btn_extract.setEnabled(False)
            self._clear_extract_preview()
            return

        if dates is not None and not picking and (not start or not end):
            self.extract_summary.setText("No months available in this column.")
            set_severity(self.extract_result_card, "warn")
            self.btn_extract.setEnabled(False)
            self._clear_extract_preview()
            return

        exclude = self._extract_exclusion_mask()
        plan = extract_core.plan_extract(
            self.df, dates, start=start or "", end=end or "", months=months,
            columns=columns, exclude=exclude,
            include_undated=self.extract_undated_check.isChecked())
        self._extract_plan = plan

        if dates is not None:
            self.extract_undated_check.setText(
                f"Also include {plan.undated_rows:,} rows with no date")
            # Hidden when there are none, for the same reason it is hidden with
            # no date filter at all: "Also include 0 rows" invites the reader to
            # think rows are being dropped, when the column has no gaps to drop.
            self.extract_undated_check.setVisible(plan.undated_rows > 0)

        lines = [f"{plan.rows:,} rows × {plan.columns} columns "
                 f"≈ {plan.cells:,} cells"]
        if dates is not None:
            # Named here because a picked set is not a range, and the count on
            # its own gives the reader no way to tell which it was.
            lines[0] += ("  ·  " + (extract_core.months_label(months)
                                    if picking
                                    else extract_core.range_label(start, end)))
        if plan.excluded_rows:
            # Said on the same line as the count, since it is the difference
            # between this number and the one the month totals imply.
            lines[0] += f"  ·  {plan.excluded_rows:,} cancelled left out"
        if plan.rows == 0:
            lines.append("Nothing matches this selection.")
            severity = "warn"
        elif plan.exceeds_excel_rows:
            lines.append(
                f"More rows than an Excel worksheet holds "
                f"({DETAIL_MAX_DATA_ROWS + 1:,}). Narrow the range, or choose "
                f"CSV, which has no limit.")
            severity = "error"
        elif plan.may_be_slow_in_excel:
            # A rule of thumb, and labelled as one - see EXCEL_SLOW_CELLS.
            lines.append("Excel will open this, but expect it to feel slow. "
                         "Dropping columns helps more than shortening the range.")
            severity = "warn"
        else:
            lines.append("Comfortable for Excel.")
            severity = "ok"

        chosen_format = self.extract_format_combo.currentData()
        if chosen_format == extract_core.FORMAT_XLSX and plan.rows:
            # From the measured ~59k text cells/s; a rough figure beats none,
            # because the alternative is the user waiting with no idea.
            lines.append(f"Writing .xlsx will take roughly "
                         f"{max(1, round(plan.cells / 59_000)):,}s. "
                         f"CSV would be near-instant.")

        self.extract_summary.setText("\n".join(lines))
        self._update_extract_preview(dates, start, end, months, columns, plan)
        set_severity(self.extract_result_card, severity)
        self.btn_extract.setEnabled(
            plan.rows > 0 and not (plan.exceeds_excel_rows
                                   and chosen_format == extract_core.FORMAT_XLSX))

    def _extract_show_partial_months(self, start, end, months,
                                     filtering: bool) -> None:
        """
        Name the months the file only partly covers, inside the chosen range.

        A partial month's row count is real but not comparable: an export taken
        on the 12th makes that month look like a collapse in volume when nothing
        happened except the month not being over. Judged on the span of days
        present rather than how many days have rows, so a weekday-only queue -
        which is missing eight or nine days of every month - is not accused of
        being incomplete.
        """
        coverage = getattr(self, "_extract_coverage", None)
        if not filtering or not coverage:
            self.extract_partial_row.setVisible(False)
            return

        if months is not None:
            chosen = set(months)
            partial = [c for c in coverage if c.partial and c.period in chosen]
        else:
            partial = extract_core.partial_months(coverage, start or "",
                                                  end or "")
        if not partial:
            self.extract_partial_row.setVisible(False)
            return

        if months is not None:
            # Ticking is not restricted to a span, so *any* partial month can be
            # dropped here - including one in the middle, which the range mode
            # cannot reach.
            trimmable, interior = partial, []
        else:
            # Only the ends can be dropped by narrowing a range; a partial month
            # in the middle is reported and left, since excluding it would put a
            # hole in the extract rather than shorten it.
            trimmable = [c for c in partial
                         if c.period in {(start or ""), (end or "")}]
            interior = [c for c in partial if c not in trimmable]

        self._extract_trimmable = [c.period for c in trimmable]
        # Offered only when pressing it would actually change the selection. A
        # button that does nothing is worse than no button - which happens when
        # the partial month is the only one chosen, in either mode.
        if months is not None:
            can_trim = bool(trimmable) and len(months) > len(trimmable)
        else:
            can_trim = self._extract_trim_targets() != (
                self.extract_from_combo.currentIndex(),
                self.extract_to_combo.currentIndex())

        where = "selection" if months is not None else "range"
        lines = [(f"Incomplete month in this {where}: "
                         if len(partial) == 1
                         else f"Incomplete months in this {where}: ")]
        lines[0] += "; ".join(c.describe() for c in partial) + "."
        if can_trim:
            names = " and ".join(extract_core.month_label(c.period)
                                 for c in trimmable)
            lines.append(f"Those rows are real, but counting them against a "
                         f"whole month is not. Exclude {names} to compare like "
                         f"with like.")
        elif trimmable:
            lines.append("It is the only month selected, so there is nothing to "
                         "narrow to - widen the selection, or read the total "
                         "knowing the month is not over.")
        if interior:
            lines.append("A gap in the middle cannot be excluded by shortening "
                         "the range, so it is named here instead.")

        self.extract_partial_warning.setText(" ".join(lines))
        set_state(self.extract_partial_warning, "warn")
        self.btn_extract_trim.setVisible(can_trim)
        self.extract_partial_row.setVisible(True)

    def _extract_untick_partial_months(self) -> bool:
        """
        In pick mode, excluding a partial month is just unticking it.

        Returns whether anything changed, so the caller can leave the offer up
        only while it would do something.
        """
        partial = {c.period for c in self._extract_coverage or [] if c.partial}
        ticked = self._extract_selected_months()
        keep = [m for m in ticked if m not in partial]
        if keep == ticked or not keep:
            # Refusing to untick everything: that would leave an empty selection
            # and no extract, which is not what "exclude" means.
            return False
        self._extract_set_months(keep)
        return True

    def _extract_trim_targets(self):
        """
        Where Exclude would leave the range, as (from index, to index).

        Steps one month in from each partial end rather than jumping to the
        first complete month: two partial months in a row is a real shape - an
        export that both begins and ends mid-month - and moving further than
        asked would silently drop whole months the user had chosen. It never
        collapses the range, so a single partial month comes back unchanged and
        the offer is withdrawn rather than being a button that does nothing.
        """
        from_combo, to_combo = self.extract_from_combo, self.extract_to_combo
        start_index, end_index = from_combo.currentIndex(), to_combo.currentIndex()
        trimmable = set(getattr(self, "_extract_trimmable", []) or [])
        if not trimmable:
            return start_index, end_index

        if (from_combo.currentData() in trimmable
                and start_index + 1 <= end_index - 1):
            start_index += 1
        if (to_combo.currentData() in trimmable
                and end_index - 1 >= start_index + 1):
            end_index -= 1
        return start_index, end_index

    def _extract_trim_partial_months(self) -> None:
        """Drop the partial months: untick them, or narrow the range past them."""
        if self._extract_month_mode() == EXTRACT_MONTHS_PICK:
            self._extract_untick_partial_months()
            return

        from_combo, to_combo = self.extract_from_combo, self.extract_to_combo
        start_index, end_index = self._extract_trim_targets()
        if (start_index, end_index) == (from_combo.currentIndex(),
                                        to_combo.currentIndex()):
            return

        # Both combos moved before one recount, or the intermediate range would
        # be planned and reported for a selection that never existed.
        from_combo.blockSignals(True)
        to_combo.blockSignals(True)
        try:
            from_combo.setCurrentIndex(start_index)
            to_combo.setCurrentIndex(end_index)
        finally:
            from_combo.blockSignals(False)
            to_combo.blockSignals(False)
        self._update_extract_plan()

    def _clear_extract_preview(self) -> None:
        """Empty the preview, so it never describes a stale selection."""
        set_table_model(self.extract_preview_table, QStandardItemModel())
        self.extract_preview_title.setText("Preview")

    def _update_extract_preview(self, dates, start, end, months, columns,
                                plan) -> None:
        """
        Show the first rows of the subset as it will actually be written.

        Built with build_extract and a limit, deliberately: the same mask and the
        same column order as the file, so what is on screen cannot disagree with
        what lands on disk. A preview from a second code path would be worse
        than none.
        """
        if plan.rows == 0 or not columns:
            self._clear_extract_preview()
            return

        try:
            head = extract_core.build_extract(
                self.df, dates, start=start or "", end=end or "",
                months=months, columns=columns,
                exclude=self._extract_exclusion_mask(),
                include_undated=self.extract_undated_check.isChecked(),
                limit=EXTRACT_PREVIEW_ROWS)
        except Exception:
            log.warning("could not build the extract preview", exc_info=True)
            self._clear_extract_preview()
            return

        model = QStandardItemModel(len(head), len(head.columns))
        model.setHorizontalHeaderLabels([str(c) for c in head.columns])
        for row, (_index, values) in enumerate(head.iterrows()):
            for col, value in enumerate(values):
                text = "" if pd.isna(value) else str(value)
                model.setItem(row, col, QStandardItem(text))
        set_table_model(self.extract_preview_table, model)
        # Bounded, not fitted to all of them: the table holds
        # EXTRACT_PREVIEW_ROWS and scrolls, so nothing is withdrawn - what goes
        # away is 480px of height that pushed the pickers off a 720px window.
        fit_table_height(self.extract_preview_table,
                         max_rows=EXTRACT_PREVIEW_VISIBLE)
        self.extract_preview_title.setText(
            f"Preview — first {len(head):,} of {plan.rows:,} rows")
        # The provenance a pasted preview needs: which rows these are of what.
        set_copy_context(self.extract_preview_table,
                         **{"Rows in the extract": f"{plan.rows:,}",
                            "Shown here": f"{len(head):,}"})

    def _extract_show_format_warning(self):
        """
        Month buckets rest on how the dates were read. If that was a guess, the
        buckets are wrong and no amount of picking fixes it.
        """
        column = self.extract_date_combo.currentData()
        if not column:
            self.extract_format_warning.setVisible(False)
            return

        report = self.analyzer.check_date_formats()
        problem = next((c for c in report.get('columns', [])
                        if c['column'] == column
                        and (c.get('needs_confirmation') or c.get('conflicting'))),
                       None)
        if problem is None:
            self.extract_format_warning.setVisible(False)
            return

        if problem.get('conflicting'):
            message = (f"{column} mixes DD/MM and MM/DD, so these months are "
                       f"unreliable however the range is set. Fix it at source.")
        else:
            message = (f"{problem['ambiguous']:,} values in {column} could be "
                       f"read as either DD/MM or MM/DD, so these month buckets "
                       f"rest on a guess. Set the order in Logic Checks first.")
        self.extract_format_warning.setText(message)
        set_state(self.extract_format_warning, "warn")
        self.extract_format_warning.setVisible(True)

    # -- writing ----------------------------------------------------------

    def run_extract(self):
        """
        Write the current selection.

        Mirrors export_report, including its hard-won detail: every input is
        resolved on the UI thread before dispatch. A closure over self.df would
        resolve when the worker *runs*, and this view exists to be fiddled with,
        so the user changing the selection mid-write is likely rather than
        theoretical.
        """
        if self.df is None or self.analyzer is None:
            return

        columns = self._extract_selected_columns()
        if not columns:
            return

        fmt = self.extract_format_combo.currentData()
        date_column = self.extract_date_combo.currentData() or ""
        start = self.extract_from_combo.currentData() or ""
        end = self.extract_to_combo.currentData() or ""
        include_undated = self.extract_undated_check.isChecked()
        exclude = self._extract_exclusion_mask()
        excluded_states = (list(self._cancelled_state_info()["values"])
                           if exclude is not None else [])
        picking = (date_column
                   and self._extract_month_mode() == EXTRACT_MONTHS_PICK)
        months = self._extract_selected_months() if picking else None
        if picking and not months:
            return

        suggested = extract_core.suggest_filename(
            self.filepath, date_column, start or "all", end or "all", fmt,
            months=months)
        start_dir = self.config.get("last_folder") or ""
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Extract", os.path.join(start_dir, suggested),
            "Excel Files (*.xlsx)" if fmt == extract_core.FORMAT_XLSX
            else "CSV Files (*.csv)")
        if not path:
            return

        # Remembered so the next extract of the same shape is one click. The
        # month mode is remembered but the months themselves are not: they
        # describe this file's span, and the next file may not contain them.
        self.config["extract_columns"] = columns
        self.config["extract_date_column"] = date_column or None
        self.config["extract_format"] = fmt
        self.config["extract_month_mode"] = self._extract_month_mode()
        self.config["extract_exclude_cancelled"] = \
            self.extract_cancelled_check.isChecked()
        self.config_mgr.config = self.config
        self.config_mgr.save()

        df = self.df
        dates = self._extract_dates()
        analyzer_filepath = self.filepath
        plan = getattr(self, "_extract_plan", None)
        # Snapshotted like everything else: this view exists to be fiddled with,
        # so the selection may well change while the write runs.
        if months is not None:
            chosen = set(months)
            incomplete = [c.period for c in self._extract_coverage or []
                          if c.partial and c.period in chosen]
        else:
            incomplete = [c.period for c in extract_core.partial_months(
                self._extract_coverage or [], start, end)]

        def task(progress=None):
            frame = extract_core.build_extract(
                df, dates, start=start, end=end, months=months,
                columns=columns, exclude=exclude,
                include_undated=include_undated)
            written = extract_core.write_extract(
                path, frame, fmt=fmt, date_column=date_column, start=start,
                end=end, plan=plan, source_path=analyzer_filepath,
                incomplete=incomplete, months=months,
                excluded_states=excluded_states, progress=progress,
                cancelled=lambda: self._extract_cancelled)
            return written

        def done(written):
            self._finish_extract()
            notes = []
            if plan is not None and plan.undated_rows and not include_undated:
                notes.append(f"{plan.undated_rows:,} rows with no date were left "
                             f"out. Tick the box above to include them.")
            if incomplete:
                # On screen this was a warning the user could act on. In the
                # finished file it is a caveat that has to travel with it, or
                # the short month reads as a fall in volume.
                names = ", ".join(extract_core.month_label(period)
                                  for period in incomplete)
                notes.append(
                    f"{names} {'is not a whole month' if len(incomplete) == 1 else 'are not whole months'} "
                    f"in this file, so those rows are not comparable with the "
                    f"other months. Recorded in the extract.")
            if plan is not None and plan.excluded_rows:
                # The one row count a reader cannot recover from the file, since
                # the excluded rows are simply absent from it.
                notes.append(
                    f"{plan.excluded_rows:,} cancelled tickets were left out "
                    f"({', '.join(excluded_states)}). Recorded in the extract.")
            if fmt == extract_core.FORMAT_CSV:
                notes.append(
                    "The date column and range are recorded beside it in "
                    + os.path.basename(extract_core.info_sidecar_path(path)) + ".")
            described = (extract_core.months_label(months) if months is not None
                         else extract_core.range_label(start, end))
            body = (f"{written:,} rows × {len(columns)} columns "
                    f"({described}) written to:\n{path}")
            # The button stays in the view rather than living on the dialog:
            # after a write measured in tens of seconds, the folder is still
            # what you want a minute later, once the dialog has been dismissed.
            self._extract_last_path = path
            self.btn_extract_reveal.setVisible(True)
            QMessageBox.information(
                self, "Extract complete",
                body + ("\n\n" + "\n\n".join(notes) if notes else ""))

        def failed(message):
            self._finish_extract()
            if self._extract_cancelled:
                # Not an error: the user asked. Said plainly, including that
                # nothing was left behind, since a half file would be worse than
                # no file and they cannot see which they got.
                self.update_status("Extract cancelled - no file was written.")
                QMessageBox.information(
                    self, "Extract cancelled",
                    "The extract was cancelled and the part-written file was "
                    "deleted, so nothing incomplete is left behind.")
                return
            log.warning("extract failed: %s", message)
            self.show_error(f"Extract failed: {message}")

        self._begin_extract()
        self.run_in_background(task, done, failed, report_progress=True)

    def _reveal_extract(self) -> None:
        """Open the folder holding the last extract."""
        path = getattr(self, "_extract_last_path", None)
        if not path:
            return
        folder = os.path.dirname(os.path.abspath(path))
        if not os.path.isdir(folder):
            self.show_error(f"That folder is no longer there:\n{folder}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(folder))

    def _cancel_extract(self) -> None:
        """
        Ask the worker to stop. It checks between blocks of rows, so the write
        ends at the next checkpoint rather than instantly.
        """
        self._extract_cancelled = True
        self.btn_extract_cancel.setEnabled(False)
        self.update_status("Cancelling the extract…")

    def _begin_extract(self) -> None:
        self._extract_cancelled = False
        self.btn_extract.setEnabled(False)
        self.btn_extract_cancel.setEnabled(True)
        self.btn_extract_cancel.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setVisible(True)

    def _finish_extract(self) -> None:
        self.btn_extract_cancel.setVisible(False)
        self.progress_bar.setVisible(False)
        # Via the planner, so the button's enabled state stays consistent with
        # whatever the selection now is rather than being blindly re-enabled.
        self._update_extract_plan()

    # 12. Report - the document this review produces
    #
    # The one view whose result is legitimately plain text: it *is* a text
    # document, written to be pasted into a ticket or a mail client, so a table
    # would be a worse rendering of it rather than a better one. Nothing here is
    # hand-built markup either - the text comes out of core/reporter.py and goes
    # into a read-only QPlainTextEdit unchanged.
    REPORT_FULL = "Full report"
    REPORT_SUMMARY = "Summary"
    REPORT_FLAVOURS = [REPORT_FULL, REPORT_SUMMARY]

    def create_report_tab(self, parent):
        """
        Deliberately not scrollable. A QPlainTextEdit scrolls itself, so wrapping
        the page would nest two scrollbars and leave the document's own on the
        inside, where it is hardest to reach. The text area takes the layout's
        stretch instead, and nothing in this view pins a height - so the rule
        that a pinning view must scroll does not apply.
        """
        layout = make_view(parent)

        card, card_layout = make_card(margin=SPACE_SM)

        bar = QHBoxLayout()
        bar.setSpacing(SPACE_XS)
        # Not "REPORT" - the chrome already prints the view name above this, and
        # three views used to say theirs twice in two different styles.
        title = QLabel("DOCUMENT")
        title.setObjectName("CardTitle")
        bar.addWidget(title)
        bar.addSpacing(SPACE_MD)
        # Two documents, not two renderings of one: the full report walks every
        # section in order, the summary sorts the same findings into critical /
        # warning / passed and names the actions they imply. Both are model-free.
        self.report_tabs = make_seg_tabs(bar, self.REPORT_FLAVOURS,
                                         self._show_report_flavour)
        bar.addStretch()
        self.btn_report_copy = QPushButton("Copy to clipboard")
        self.btn_report_copy.setObjectName("GhostButton")
        self.btn_report_copy.setToolTip(
            "Put the report on the clipboard exactly as it reads here")
        self.btn_report_copy.clicked.connect(self._copy_report)
        bar.addWidget(self.btn_report_copy)
        self.btn_report_save = QPushButton("Save as .txt")
        self.btn_report_save.setObjectName("GhostButton")
        self.btn_report_save.setToolTip(
            "Write the report to a text file beside your data")
        self.btn_report_save.clicked.connect(self._save_report)
        bar.addWidget(self.btn_report_save)
        card_layout.addLayout(bar)

        # What is on screen and what it was measured over. It carries the copy
        # context as well, so a sentence pasted out of it can be traced back to
        # a sheet and a flavour.
        self.report_note = QLabel("")
        self.report_note.setWordWrap(True)
        set_state(self.report_note, "muted")
        card_layout.addWidget(self.report_note)

        # Both generators run every check in the app - ~1.2s at 100k rows and
        # ~3.7s at 300k - so arriving here does not start one. Until the user
        # asks, this card is what the panel holds: it states what has not
        # happened rather than showing an empty pane, which reads as a report
        # that found nothing to say.
        self.report_prompt, prompt_layout = make_card(margin=SPACE_MD)
        prompt_title = QLabel("No report has been written yet")
        prompt_title.setObjectName("CardTitle")
        prompt_layout.addWidget(prompt_title)
        self.report_prompt_body = QLabel()
        self.report_prompt_body.setWordWrap(True)
        prompt_layout.addWidget(self.report_prompt_body)
        prompt_actions = QHBoxLayout()
        prompt_actions.setSpacing(SPACE_SM)
        self.btn_report_run = QPushButton("Write the report")
        self.btn_report_run.setObjectName("PrimaryButton")
        self.btn_report_run.setToolTip(
            "Run every check on this sheet and write the report")
        self.btn_report_run.clicked.connect(self._start_report)
        prompt_actions.addWidget(self.btn_report_run)
        prompt_actions.addStretch()
        prompt_layout.addLayout(prompt_actions)
        self.report_prompt_hint = QLabel(
            "Language findings are included only for the columns Language Check "
            "has already run on. Detection is far too slow to repeat here, and "
            "repeating it would report different numbers than the screen.")
        self.report_prompt_hint.setObjectName("Caption")
        self.report_prompt_hint.setWordWrap(True)
        prompt_layout.addWidget(self.report_prompt_hint)
        card_layout.addWidget(self.report_prompt)

        # Read-only and mono: the report rules its headings with "=" and lines
        # its counts up in columns, neither of which survives a proportional
        # face. setPlainText and never setHtml - the text is the artifact and
        # has to reach the clipboard as written. Wrapping is set per flavour in
        # _populate_report, because the two documents are different shapes.
        self.report_text = QPlainTextEdit()
        self.report_text.setObjectName("ReportText")
        self.report_text.setReadOnly(True)
        self.report_text.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        card_layout.addWidget(self.report_text, 1)

        # Collects the leftover height while the document is hidden, for the
        # reason Overview's register card carries one: a QVBoxLayout hands that
        # space to whatever will take it, and a QLabel's vertical policy is
        # Preferred rather than Fixed, so the prompt drifts down the panel under
        # a band of white. A spacer rather than a filler widget - a plain
        # QWidget is matched by the `QMainWindow, QWidget` rule and paints as a
        # grey block - and its factor is toggled rather than left at 1, or it
        # would compete with the document's own stretch.
        self._report_card_layout = card_layout
        card_layout.addStretch(0)
        self._report_filler_index = card_layout.count() - 1

        layout.addWidget(card, 1)
        # Paints the prompt rather than leaving an empty document on screen.
        # refresh_report runs straight after this, but a view that is created
        # and not yet refreshed is briefly visible.
        self._show_report_idle()

    def refresh_report(self):
        """
        Write the report off the UI thread, once the user has asked for it.

        Both generators run every check in the app - ~1.2s at 100k rows and
        ~3.7s at 300k - so this can never be synchronous, and it must say it is
        working rather than showing a pane that looks like a finished document.

        It also does not start on its own, for the reason Overview's register
        does not: `_report_started` is reset per sheet in load_sheet, because
        having written a report for one sheet says nothing about the next.
        """
        if not self._report_started:
            self._show_report_idle()
            return

        # Snapshotted before dispatch, for the reason _refresh_async binds the
        # analyzer: a closure over self.df resolves when the worker *runs*, so a
        # sheet switch mid-flight would write a document about other data and
        # label it with this sheet's name.
        frame = self.df
        filepath = self.filepath
        lang_results = list(self.language_results.values())
        threshold = float(self.config.get("null_threshold", 20))
        desc_min = int(self.config.get("desc_min_length", 20))
        desc_max = int(self.config.get("desc_max_length", 5000))
        flavour = self._report_flavour
        summary = flavour == self.REPORT_SUMMARY

        self._report_text = None
        self._set_report_running(True)
        self.report_text.setPlainText("")
        self.report_note.setText(f"Writing the {flavour.lower()}...")
        set_state(self.report_note, "busy")

        def task(analyzer):
            reporter = ReportGenerator(
                frame, analyzer, filepath, null_threshold=threshold,
                desc_min_length=desc_min, desc_max_length=desc_max,
                lang_results=lang_results)
            text = (reporter.generate_summary_report() if summary
                    else reporter.generate_text_report())
            return {'flavour': flavour, 'text': text, 'rows': len(frame)}

        # The flavour is part of the key, so pressing the tab that is already
        # current coalesces while switching to the other one is not dropped as
        # a duplicate refresh.
        self._refresh_async(f"Report:{flavour}", task, self._populate_report,
                            busy_msg=f"Writing the {flavour.lower()}...",
                            err_callback=self._report_failed)

    def _populate_report(self, result) -> None:
        """Put the finished document on screen."""
        text = result['text']
        self._report_text = text
        # What is actually in the pane, which is not necessarily the tab that is
        # current: two flavours dispatched in quick succession both complete.
        self._report_flavour_shown = result['flavour']

        # Wrapping follows the document, not a preference. The full report is
        # column-aligned - a reflowed line breaks the alignment it spent its
        # width on - while the summary is prose paragraphs, which unwrapped run
        # off the right edge and have to be read on a horizontal scrollbar.
        self.report_text.setLineWrapMode(
            QPlainTextEdit.LineWrapMode.WidgetWidth
            if result['flavour'] == self.REPORT_SUMMARY
            else QPlainTextEdit.LineWrapMode.NoWrap)
        self.report_text.setPlainText(text)
        # setPlainText leaves the cursor at the end of the document; the reader
        # wants the top of it.
        self.report_text.moveCursor(QTextCursor.MoveOperation.Start)
        self._set_report_running(True)

        lines = text.count("\n") + 1
        self.report_note.setText(
            f"{result['flavour']}, {lines:,} lines over {result['rows']:,} "
            f"rows. Copy it or save it as a .txt - the text is the deliverable.")
        set_state(self.report_note, "muted")
        set_copy_context(self.report_note, **{
            "Report": result['flavour'],
            "Rows analyzed": f"{result['rows']:,}",
        })
        self.update_status(f"{result['flavour']} written.")

    def _start_report(self) -> None:
        """The button under the prompt. Writes the report for this sheet."""
        if self.df is None or self.analyzer is None:
            return
        self._report_started = True
        self.refresh_report()

    def _report_failed(self, err) -> None:
        """
        A failed run returns to the prompt rather than to an empty pane.

        Without this the note says it is working for ever and the button that
        would retry is hidden behind the document, so the only way out is
        reloading the sheet.
        """
        log.warning("could not write the report: %s", err)
        self._report_started = False
        self._show_report_idle()
        self.report_prompt_body.setText(
            f"The report could not be written for this sheet: {err}\n\n"
            f"Every other view still works. Press the button to try again.")
        set_state(self.report_prompt_body, "error")
        self.btn_report_run.setText("Try again")
        self.update_status("The report could not be written.")

    def _set_report_running(self, running: bool) -> None:
        """
        Swap the panel between the prompt and the document.

        The pane is hidden rather than left empty for the reason Duplicate Check
        hides its splitter: a blank document is a worse answer than one sentence
        saying what has not happened. The flavour strip goes with it - choosing
        between two documents, neither of which exists, is a control that cannot
        do anything.

        The two actions are gated on there being text rather than on `running`,
        because between dispatch and completion the pane is on screen and still
        empty. A Copy that puts nothing on the clipboard is worse than one that
        is briefly greyed out.
        """
        self.report_prompt.setVisible(not running)
        self._report_card_layout.setStretch(
            self._report_filler_index, 0 if running else 1)
        self.report_text.setVisible(running)
        self.report_note.setVisible(running)
        for button in self.report_tabs.values():
            button.setVisible(running)

        has_text = bool(self._report_text)
        self.btn_report_copy.setVisible(running)
        self.btn_report_save.setVisible(running)
        self.btn_report_copy.setEnabled(has_text)
        self.btn_report_save.setEnabled(has_text)

    def _show_report_idle(self) -> None:
        """
        The view before anyone has asked for a report.

        The pane is emptied as well as hidden: leaving the previous sheet's
        document in it would mean a report describing one file sat behind the
        prompt for another.
        """
        self._report_text = None
        self.report_text.setPlainText("")
        self._set_report_running(False)
        self.btn_report_run.setText("Write the report")
        set_state(self.report_prompt_body, "muted")

        rows = len(self.df) if self.df is not None else 0
        self.report_prompt_body.setText(
            f"This sheet has {rows:,} rows. Both reports work through every "
            f"check in the app and can take a few seconds on a file this size, "
            f"so they wait until you ask.\n\n"
            f"The full report walks every section in order and is meant to be "
            f"filed alongside the data. The summary sorts the same findings "
            f"into critical, warnings and passed and names the actions they "
            f"imply, which is the one to paste into a mail.")

    def _show_report_flavour(self, name: str) -> None:
        """
        Switching flavour re-generates.

        They are two documents built by different code, not two views of one
        frame, so there is nothing cached to filter - and a stale document under
        the other tab's label would be the worst of both.
        """
        if name == self._report_flavour:
            return
        self._report_flavour = name
        if self._report_started:
            self.refresh_report()

    def _copy_report(self) -> None:
        """
        The report, verbatim.

        No provenance block on top, unlike a copied table: the full report
        already names the file and when it was written, and the summary is a
        mail draft whose first line is its subject - a header above that would
        have to be deleted by hand every time. The note above the document
        carries the provenance instead, and copies with it.
        """
        if not self._report_text:
            return
        QApplication.clipboard().setText(self._report_text)
        # The status bar, not a message box: a modal on every copy is noise, and
        # this is where the app says this kind of thing everywhere else.
        self.update_status(
            f"{self._report_flavour_shown or 'Report'} copied to the clipboard.")

    def _save_report(self) -> None:
        """
        Write the report to a text file.

        utf-8-sig, for the reason the CSV sidecars use it: without the BOM
        Notepad and Excel mis-decode non-ASCII on a double-click and mangle
        exactly the non-English text this tool exists to find.
        """
        if not self._report_text:
            return
        stem = (os.path.splitext(os.path.basename(self.filepath))[0]
                if self.filepath else "report")
        slug = (self._report_flavour_shown or "report").lower().replace(" ", "_")
        start_dir = self.config.get("last_folder") or ""
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Report",
            os.path.join(start_dir, f"{stem}_{slug}.txt"),
            "Text Files (*.txt)")
        if not path:
            return

        try:
            with open(path, "w", encoding="utf-8-sig", newline="\n") as handle:
                handle.write(self._report_text)
        except OSError as err:
            log.exception("could not write the report to %s", path)
            self.update_status(f"Could not save the report: {err}")
            return
        self.update_status(f"Report saved to {os.path.basename(path)}.")

    # 13. Settings
    def create_settings_tab(self, parent):
        layout = QVBoxLayout(parent)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        # Form layout manually
        # make_card, for the reason Description Quality needed it: QSS padding
        # does not move a laid-out QFrame's children, so a bare Card frame has
        # none.
        form, fl = make_card(margin=SPACE_MD)
        
        # Section: Analysis Settings
        analysis_header = QLabel("Analysis Settings")
        analysis_header.setObjectName("SectionHeader")
        fl.addWidget(analysis_header)
        
        fl.addWidget(QLabel("Null Threshold (%):"))
        # A spin box rather than a free-text field: the old QLineEdit swallowed
        # invalid input in a bare except, so a typo silently kept the old value.
        self.sett_thresh = QDoubleSpinBox()
        self.sett_thresh.setRange(0.0, 100.0)
        self.sett_thresh.setDecimals(1)
        self.sett_thresh.setSingleStep(1.0)
        self.sett_thresh.setSuffix(" %")
        self.sett_thresh.setValue(float(self.config.get("null_threshold", 20)))
        self.sett_thresh.setToolTip("Columns with a higher share of nulls are flagged")
        self.sett_thresh.setMaximumWidth(120)
        fl.addWidget(self.sett_thresh)
        
        fl.addSpacing(SPACE_XL)
        
        # Save button
        btn_save = QPushButton("Save Settings")
        btn_save.setObjectName("PrimaryButton")
        btn_save.clicked.connect(self.save_settings)
        fl.addWidget(btn_save)
        
        layout.addWidget(form)
        layout.addStretch()
        
    def save_settings(self):
        self.config["null_threshold"] = self.sett_thresh.value()

        self.config_mgr.config = self.config
        self.config_mgr.save()
        QMessageBox.information(self, "Saved", "Settings saved!")


    # --- Other Methods ---
    
    def export_report(self):
        if self.df is None or self.analyzer is None:
            return

        rows = len(self.df)
        if rows > DETAIL_MAX_DATA_ROWS:
            # Excel cannot hold the detail sheet, and the choice of what to do
            # about it is the user's, so it cannot live in core/.
            if not self._confirm_row_limit(rows):
                return

        path, _ = QFileDialog.getSaveFileName(
            self, "Save Report", "sanity_report.xlsx", "Excel Files (*.xlsx)")
        if not path:
            return

        # Everything the worker needs, resolved on the UI thread.
        #
        # df and analyzer are snapshotted rather than read through self: a
        # closure resolves them when the worker *runs*, so a sheet switch during
        # a long export would silently redirect the remaining sheets onto the
        # new sheet's data and produce one workbook describing two sheets. Same
        # reason _refresh_async binds the analyzer at dispatch.
        df = self.df
        analyzer = self.analyzer
        filepath = self.filepath
        seq = self._data_seq
        lang_results = list(self.language_results.values())
        null_threshold = float(self.config.get("null_threshold", 20))
        desc_min = int(self.config.get("desc_min_length", 20))
        desc_max = int(self.config.get("desc_max_length", 5000))
        native_pivots = bool(self.config.get("report_native_pivots", True))

        def task(progress=None):
            reporter = ReportGenerator(
                df, analyzer, filepath,
                null_threshold=null_threshold,
                desc_min_length=desc_min,
                desc_max_length=desc_max,
                lang_results=lang_results,
                native_pivots=native_pivots,
            )
            reporter.export_to_excel(path, progress=progress)
            return reporter

        def done(reporter):
            self._finish_export()
            if seq != self._data_seq:
                # The sheet changed mid-export. The file is still a valid report
                # of the sheet it was started for, so say which one.
                self.update_status(f"Report written for the previous sheet: {path}")
            self._show_export_summary(path, reporter)

        def failed(message):
            self._finish_export()
            log.warning("export failed: %s", message)
            self.show_error(f"Export failed: {message}")

        self._begin_export()
        self.run_in_background(task, done, failed, report_progress=True)

    def _begin_export(self) -> None:
        """Gate the export controls and show activity for the duration."""
        # Both were left live, so two exports could target the same path at once.
        self.btn_export.setEnabled(False)
        self.export_action.setEnabled(False)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setVisible(True)

    def _finish_export(self) -> None:
        self.progress_bar.setVisible(False)
        can_export = self.df is not None
        self.btn_export.setEnabled(can_export)
        self.export_action.setEnabled(can_export)

    def _confirm_row_limit(self, rows: int) -> bool:
        """
        Warn before exporting more rows than an Excel worksheet can hold.

        Naming both numbers matters: the detail data has to be truncated either
        way, and a reader who is not told will derive wrong totals from it. The
        companion CSV is not subject to the limit, which is the useful part of
        the answer.
        """
        limit = DETAIL_MAX_DATA_ROWS
        answer = QMessageBox.warning(
            self, "More rows than a worksheet holds",
            f"This sheet has {rows:,} rows. An Excel worksheet holds "
            f"{limit + 1:,} including the header.\n\n"
            f"The per-ticket detail will be written as a CSV alongside the "
            f"report, which has no such limit, and the report will state what "
            f"was included.\n\nContinue?",
            QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Ok)
        return answer == QMessageBox.StandardButton.Ok

    def _show_export_summary(self, path: str, reporter) -> None:
        """
        Report what was written, including anything the workbook could not hold.

        The old dialog said "Done!" and nothing else - not the path, not that a
        second file had appeared, not that a check had been skipped.
        """
        lines = [f"Report written to:\n{path}"]
        if getattr(reporter, "detail_path", None):
            lines.append(
                f"\nPer-ticket detail ({reporter.detail_rows_written:,} rows):\n"
                f"{os.path.basename(reporter.detail_path)}")
        disclosures = getattr(reporter, "export_disclosures", None) or []
        if disclosures:
            shown = disclosures[:6]
            lines.append("\nNotes:")
            lines.extend(f"  • {note}" for note in shown)
            if len(disclosures) > len(shown):
                lines.append(f"  • …and {len(disclosures) - len(shown)} more "
                             f"(see the Overview sheet)")
        QMessageBox.information(self, "Export complete", "\n".join(lines))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    apply_dark_palette(app)

    window = SanityCheckApp()
    window.show()
    
    sys.exit(app.exec())
