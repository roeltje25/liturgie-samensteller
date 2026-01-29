"""Main entry point for the Liturgie Samensteller application."""

import sys
import os

# Minimal PyQt6 imports - get splash visible ASAP
from PyQt6.QtWidgets import QApplication, QSplashScreen, QMessageBox, QFileDialog
from PyQt6.QtGui import QPixmap, QPainter, QFont, QColor, QLinearGradient
from PyQt6.QtCore import Qt


def _create_splash_pixmap():
    """Create splash pixmap inline to avoid module import delay."""
    width, height = 500, 350
    pixmap = QPixmap(width, height)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    # Background gradient
    gradient = QLinearGradient(0, 0, 0, height)
    gradient.setColorAt(0, QColor(40, 60, 100))
    gradient.setColorAt(1, QColor(20, 30, 50))
    painter.fillRect(0, 0, width, height, gradient)

    # Border
    painter.setPen(QColor(100, 140, 200))
    painter.drawRect(0, 0, width - 1, height - 1)

    # Title
    painter.setPen(QColor(255, 255, 255))
    painter.setFont(QFont("Segoe UI", 28, QFont.Weight.Bold))
    painter.drawText(0, 60, width, 50, Qt.AlignmentFlag.AlignCenter, "Liturgie")
    painter.drawText(0, 100, width, 50, Qt.AlignmentFlag.AlignCenter, "Samensteller")

    # Subtitle
    painter.setFont(QFont("Segoe UI", 12))
    painter.setPen(QColor(180, 200, 230))
    painter.drawText(0, 160, width, 30, Qt.AlignmentFlag.AlignCenter, "PowerPoint Mixer")

    # Cross decoration
    painter.setFont(QFont("Consolas", 10))
    painter.setPen(QColor(120, 160, 220))
    cross = ["      ┃      ", "      ┃      ", "  ━━━━╋━━━━  ", "      ┃      ", "      ┃      "]
    for i, line in enumerate(cross):
        painter.drawText(0, 200 + i * 16, width, 20, Qt.AlignmentFlag.AlignCenter, line)

    # Loading text
    painter.setPen(QColor(150, 180, 220))
    painter.setFont(QFont("Segoe UI", 10))
    painter.drawText(0, height - 40, width, 30, Qt.AlignmentFlag.AlignCenter, "Loading...")

    painter.end()
    return pixmap


def _handle_first_run(base_path: str, settings) -> None:
    """Handle first-run setup (called after splash is closed)."""
    from .i18n import tr

    # Show welcome message
    msg = QMessageBox()
    msg.setWindowTitle(tr("dialog.firstrun.title"))
    msg.setText(tr("dialog.firstrun.message"))
    msg.setIcon(QMessageBox.Icon.Information)
    msg.addButton(tr("dialog.firstrun.select_folder"), QMessageBox.ButtonRole.AcceptRole)
    msg.exec()

    # Open folder selection
    folder = QFileDialog.getExistingDirectory(
        None,
        tr("dialog.settings.base_folder"),
        base_path
    )

    if folder:
        settings.base_folder = folder
        settings.save()  # Saves to AppData/LiturgieSamensteller/settings.json


def main():
    """Run the application."""
    # Determine base path early
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Create app and show splash immediately
    app = QApplication(sys.argv)
    splash = QSplashScreen(_create_splash_pixmap())
    splash.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()

    # Heavy imports while splash is visible
    from .logging_config import log_startup_info
    from .models import Settings
    from .ui import MainWindow
    from . import __version__

    log_startup_info()

    app.setApplicationName("Liturgie Samensteller")
    app.setApplicationVersion(__version__)
    app.setOrganizationName("PowerPoint Mixer")

    # Load settings and check for first run
    splash.showMessage("Loading settings...", Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter)
    app.processEvents()

    settings = Settings.load()  # Loads from AppData/LiturgieSamensteller/settings.json

    # Handle first run BEFORE creating main window (close splash first)
    if settings.is_first_run():
        splash.close()
        _handle_first_run(base_path, settings)
        # Recreate splash for remaining initialization
        splash = QSplashScreen(_create_splash_pixmap())
        splash.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        splash.show()
        app.processEvents()

    # Create main window
    splash.showMessage("Initializing...", Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter)
    app.processEvents()

    window = MainWindow(base_path, skip_first_run_check=True)

    splash.finish(window)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
