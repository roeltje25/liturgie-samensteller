"""Splash screen for application startup."""

from PyQt6.QtWidgets import QSplashScreen, QApplication
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QPainter, QFont, QColor, QLinearGradient


def create_splash() -> QSplashScreen:
    """Create and return a splash screen with the app banner."""
    # Create a pixmap for the splash
    width, height = 500, 350
    pixmap = QPixmap(width, height)

    # Paint the splash content
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

    # Title text
    painter.setPen(QColor(255, 255, 255))
    title_font = QFont("Segoe UI", 28, QFont.Weight.Bold)
    painter.setFont(title_font)
    painter.drawText(0, 60, width, 50, Qt.AlignmentFlag.AlignCenter, "Liturgie")
    painter.drawText(0, 100, width, 50, Qt.AlignmentFlag.AlignCenter, "Samensteller")

    # Subtitle
    subtitle_font = QFont("Segoe UI", 12)
    painter.setFont(subtitle_font)
    painter.setPen(QColor(180, 200, 230))
    painter.drawText(0, 160, width, 30, Qt.AlignmentFlag.AlignCenter, "PowerPoint Mixer")

    # ASCII art style decoration (simplified cross)
    deco_font = QFont("Consolas", 10)
    painter.setFont(deco_font)
    painter.setPen(QColor(120, 160, 220))

    cross = [
        "      ┃      ",
        "      ┃      ",
        "  ━━━━╋━━━━  ",
        "      ┃      ",
        "      ┃      ",
    ]
    y_start = 200
    for i, line in enumerate(cross):
        painter.drawText(0, y_start + i * 16, width, 20, Qt.AlignmentFlag.AlignCenter, line)

    # Loading text
    painter.setPen(QColor(150, 180, 220))
    loading_font = QFont("Segoe UI", 10)
    painter.setFont(loading_font)
    painter.drawText(0, height - 40, width, 30, Qt.AlignmentFlag.AlignCenter, "Loading...")

    painter.end()

    # Create splash screen
    splash = QSplashScreen(pixmap)
    splash.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)

    return splash


def show_splash(app: QApplication) -> QSplashScreen:
    """Create, show splash screen and process events to display it immediately."""
    splash = create_splash()
    splash.show()
    app.processEvents()
    return splash
