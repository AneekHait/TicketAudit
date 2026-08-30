"""
Render every view to PNG, for looking at.

Development tool, not part of the app or the test suite. Run it before and after
a visual change and compare the two folders:

    python scripts/ui_shots.py before
    ... make changes ...
    python scripts/ui_shots.py after

Two things make this work, and both are easy to get wrong:

1. **The native platform plugin, not offscreen.** QT_QPA_PLATFORM=offscreen has
   *zero* font families in this environment, so every glyph renders as a tofu
   box and the screenshot is worthless for judging type. This script clears that
   variable deliberately.
2. **WA_DontShowOnScreen instead of not showing the window.** An unshown window
   never runs a layout pass - resize() is ignored and every page keeps its
   default height - so the capture would not match what a user sees. This
   attribute gives the full layout pass and real fonts while keeping the window
   off the desktop.

Charts are captured with animation disabled *before* the chart is built. Bars
animate up from zero height, and a processEvents loop never advances the
animation clock, so a grab would otherwise show an empty plot area with the
value labels stacked at the bottom (see CLAUDE.md, Charts).
"""
import os
import sys

# Must happen before QApplication: see the note above.
os.environ.pop("QT_QPA_PLATFORM", None)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import numpy as np                                                  # noqa: E402
import pandas as pd                                                 # noqa: E402
from PySide6.QtCharts import QChart                                 # noqa: E402
from PySide6.QtCore import Qt                                       # noqa: E402
from PySide6.QtWidgets import QApplication                          # noqa: E402

# Disabled before gui.app_pyside builds any chart.
QChart.setAnimationOptions = lambda self, options: None

from config import ConfigManager                                    # noqa: E402
from core.analyzer import SanityAnalyzer                            # noqa: E402
import gui.app_pyside as ui                                         # noqa: E402

# Logical sizes, not pixel sizes - this machine reports a 1.5x device pixel
# ratio, so a 1280x720 window grabs as a 1920x1080 image.
#
# 1280x720 is the one that matters most and was missing: a 14-inch 1080p laptop
# at Windows' default 150% scaling, maximised. That is the machine this tool is
# actually used on, and it has *less* usable height than the 1100x700 "small
# window" case while being wider - which is exactly the shape that hides a
# pinned table's last rows. 1536x864 is the same panel at 125%.
SIZES = ((1500, 950), (1280, 720), (1536, 864), (1100, 700))


def sample_frame(rows: int = 5000) -> pd.DataFrame:
    """
    A frame shaped like the real thing: wide, with repeated code fields, mostly
    unique free text, a part-month at the end, some undated rows, cancellations
    and a little non-English text. A narrow, tidy fixture hides most of what
    these screenshots are for.
    """
    rng = np.random.default_rng(7)
    stamps = pd.date_range("2025-08-01", periods=rows, freq="90min")
    frame = pd.DataFrame({
        "Number": [f"INC{i:07d}" for i in range(rows)],
        "State": rng.choice(["Closed", "Resolved", "In Progress", "New",
                             "Cancelled"], rows, p=[.42, .2, .18, .13, .07]),
        "Priority": rng.choice(["1 - Critical", "2 - High", "3 - Moderate",
                                "4 - Low"], rows, p=[.08, .22, .5, .2]),
        "Assignment group": rng.choice(["Service Desk", "Network", "Apps",
                                        "Infrastructure"], rows),
        "Configuration item": rng.choice([f"SRV-{i:03d}" for i in range(40)], rows),
        "Category": rng.choice(["Software", "Hardware", "Network", "Access"], rows),
        "Created": stamps,
        "Resolved": stamps + pd.to_timedelta(rng.integers(1, 96, rows), unit="h"),
        "Short description": rng.choice([
            "Password reset needed for the account",
            "Printer offline on floor three refusing all jobs",
            "VPN disconnects every few minutes on this laptop",
            "Le serveur ne repond plus depuis ce matin",
            "Der Drucker funktioniert nicht mehr richtig",
        ], rows),
        "Description": [f"A sufficiently detailed description, ref {i}"
                        for i in range(rows)],
    })
    # Columns nobody fills, so the fill share and the Non-empty preset have
    # something to say.
    for i in range(6):
        frame[f"u_custom_field_{i}"] = rng.choice(["", "", "value"], rows)
    frame.loc[frame.sample(40, random_state=2).index, "Created"] = pd.NaT
    return frame


def main() -> int:
    label = sys.argv[1] if len(sys.argv) > 1 else "shots"
    outdir = os.path.join(os.environ.get("TEMP", "."), f"dp_ui_{label}")
    os.makedirs(outdir, exist_ok=True)

    # Never the developer's real config.
    import tempfile
    ui.ConfigManager = lambda *a, **k: ConfigManager(
        os.path.join(tempfile.mkdtemp(), "config.json"))

    app = QApplication(sys.argv[:1])
    app.setStyle("Fusion")
    ui.apply_dark_palette(app)

    frame = sample_frame()
    window = ui.SanityCheckApp()
    window.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen, True)
    window.show()

    def settle(rounds: int = 20) -> None:
        for _ in range(rounds):
            app.processEvents()
            window.threadpool.waitForDone(30000)
            app.processEvents()

    written = 0
    for width, height in SIZES:
        window.resize(width, height)
        settle(6)
        tag = f"{width}x{height}"

        window.df = None
        window._set_content_enabled(False)
        settle(4)
        window.grab().save(os.path.join(outdir, f"{tag}_00_empty.png"))
        written += 1

        window.df = frame
        window.analyzer = SanityAnalyzer(frame)
        window.filepath = "Incident Data dump_Last 12 Months.xlsx"
        window._set_content_enabled(True)
        # The chrome names the dataset in three places and they all come from
        # here; without it the shots show the no-file state around real content.
        window.set_dataset_badge()
        settle()

        for index, (name, _creator) in enumerate(window.tab_defs, start=1):
            if name in ui.NON_NAV_VIEWS:
                continue
            window.show_view(name)
            settle(8)
            # Language Check and Description Quality only fill on a click.
            if name == "Language Check":
                window.lang_col_combo.setCurrentText("Short description")
                window.run_language_check()
                settle(30)
            safe = name.replace(" ", "_")
            window.grab().save(os.path.join(outdir, f"{tag}_{index:02d}_{safe}.png"))
            written += 1

    print(f"{written} shots -> {outdir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
