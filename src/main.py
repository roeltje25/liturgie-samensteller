"""Main entry point for the Liturgie Samensteller application."""

import sys
import os

# Minimal imports first - these are fast
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt


def main():
    """Run the application."""
    # Determine base path early (no heavy imports needed)
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Create application first (required for any Qt widgets)
    app = QApplication(sys.argv)

    # Show splash screen immediately (before heavy imports)
    from .ui.splash_screen import show_splash
    splash = show_splash(app)

    # Now do the heavy imports while splash is visible
    from .logging_config import log_startup_info
    from .ui import MainWindow
    from . import __version__

    # Log startup banner and info
    log_startup_info()

    # Set application metadata
    app.setApplicationName("Liturgie Samensteller")
    app.setApplicationVersion(__version__)
    app.setOrganizationName("PowerPoint Mixer")

    # Create main window (this can take time due to services initialization)
    splash.showMessage("Initializing...", Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter)
    app.processEvents()

    window = MainWindow(base_path)

    # Close splash and show main window
    splash.finish(window)
    window.show()

    # Run event loop
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
