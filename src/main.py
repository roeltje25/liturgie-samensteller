"""Main entry point for the Liturgie Samensteller application."""

import sys
import os

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt

from .logging_config import log_startup_info
from .ui import MainWindow
from . import __version__


def main():
    """Run the application."""
    # Log startup banner and info
    log_startup_info()

    # Determine base path (where the app is run from)
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        base_path = os.path.dirname(sys.executable)
    else:
        # Running as script
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Create application
    app = QApplication(sys.argv)

    # Set application metadata
    app.setApplicationName("Liturgie Samensteller")
    app.setApplicationVersion(__version__)
    app.setOrganizationName("PowerPoint Mixer")

    # Create and show main window
    window = MainWindow(base_path)
    window.show()

    # Run event loop
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
