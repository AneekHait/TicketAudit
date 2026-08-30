#!/usr/bin/env python3
"""
TicketAudit - Main Entry Point

A GUI application for validating and analyzing ITSM ticket data (Excel/CSV).
Features:
- Column validation against required fields
- Null/missing value analysis
- Date logic checks
- Language detection (non-English text)
- Monthly inflow charts
- Pivot tables
- Duplicate detection
- Export reports to Excel

Usage:
    python main.py

Requirements:
    See requirements.txt
"""

import sys
import os

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def _verify_about_integrity():
    """Verify About section contains required credits."""
    _base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    _app_file = os.path.join(_base, "gui", "app_pyside.py")
    
    # Required strings that MUST be in the About section (base64 encoded versions)
    # These are the base64 strings for "Aneek Hait" and "TicketAudit"
    _required = [
        "".join([chr(x) for x in [81, 87, 53, 108, 90, 87, 115, 103, 83, 71, 70, 112, 100, 65, 61, 61]]),  # base64 of author
        "".join([chr(x) for x in [86, 71, 108, 106, 97, 50, 86, 48, 81, 88, 86, 107, 97, 88, 81, 61]]),  # base64 of app name
    ]
    
    try:
        if os.path.exists(_app_file):
            with open(_app_file, "r", encoding="utf-8") as f:
                _content = f.read()
            
            # Check all required strings are present
            for _req in _required:
                if _req not in _content:
                    return False
            return True
        return False
    except Exception:
        return False


def main():
    """Run the TicketAudit GUI application (PySide6)."""
    # Integrity check before launching
    if not _verify_about_integrity():
        try:
            from PySide6.QtWidgets import QApplication, QMessageBox
            app = QApplication(sys.argv)
            QMessageBox.critical(
                None,
                "Integrity Check Failed",
                "APPLICATION INTEGRITY COMPROMISED\n\n"
                "The application code has been modified.\n"
                "This copy of TicketAudit is not authentic.\n\n"
                "Please obtain an original copy from the official source.\n\n"
                "The application will now terminate."
            )
        except Exception:
            print("INTEGRITY CHECK FAILED: Application code has been tampered with!")
        sys.exit(1)
    
    try:
        from PySide6.QtWidgets import QApplication
        from PySide6.QtCore import Qt
        from gui.app_pyside import SanityCheckApp, apply_dark_palette
        
        app = QApplication(sys.argv)
        app.setStyle("Fusion")
        apply_dark_palette(app)

        window = SanityCheckApp()
        window.show()
        
        sys.exit(app.exec())
        
    except ImportError as e:
        print(f"Error: PySide6 is required to run TicketAudit.")
        print(f"Install it with: pip install PySide6")
        print(f"Details: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
