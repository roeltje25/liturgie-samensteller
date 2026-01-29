"""About dialog for the application."""

import getpass
import platform

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QLabel,
    QDialogButtonBox,
    QTextEdit,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from .. import __version__, __revision__, __build_date__
from ..i18n import tr


class AboutDialog(QDialog):
    """About dialog showing version and system information."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("dialog.about.title"))
        self.setMinimumSize(600, 500)
        self._setup_ui()

    def _setup_ui(self) -> None:
        """Setup the dialog UI."""
        layout = QVBoxLayout(self)

        # ASCII banner in a text widget with monospace font
        banner_text = QTextEdit()
        banner_text.setReadOnly(True)
        banner_text.setFont(QFont("Consolas", 9))
        banner_text.setMinimumHeight(320)

        # Build the banner content
        banner = self._get_banner()
        info = self._get_system_info()

        banner_text.setPlainText(banner + "\n" + info)
        layout.addWidget(banner_text)

        # Description label
        desc_label = QLabel(tr("dialog.about.description"))
        desc_label.setWordWrap(True)
        desc_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(desc_label)

        # OK button
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        button_box.accepted.connect(self.accept)
        layout.addWidget(button_box)

    def _get_banner(self) -> str:
        """Get the ASCII art banner."""
        return r"""
  ╔═══════════════════════════════════════════════════════════════════╗
  ║                                                                   ║
  ║   ██╗     ██╗████████╗██╗   ██╗██████╗  ██████╗ ██╗███████╗       ║
  ║   ██║     ██║╚══██╔══╝██║   ██║██╔══██╗██╔════╝ ██║██╔════╝       ║
  ║   ██║     ██║   ██║   ██║   ██║██████╔╝██║  ███╗██║█████╗         ║
  ║   ██║     ██║   ██║   ██║   ██║██╔══██╗██║   ██║██║██╔══╝         ║
  ║   ███████╗██║   ██║   ╚██████╔╝██║  ██║╚██████╔╝██║███████╗       ║
  ║   ╚══════╝╚═╝   ╚═╝    ╚═════╝ ╚═╝  ╚═╝ ╚═════╝ ╚═╝╚══════╝       ║
  ║                                                                   ║
  ║              ███████╗ █████╗ ███╗   ███╗███████╗███╗   ██╗        ║
  ║              ██╔════╝██╔══██╗████╗ ████║██╔════╝████╗  ██║        ║
  ║              ███████╗███████║██╔████╔██║█████╗  ██╔██╗ ██║        ║
  ║              ╚════██║██╔══██║██║╚██╔╝██║██╔══╝  ██║╚██╗██║        ║
  ║              ███████║██║  ██║██║ ╚═╝ ██║███████╗██║ ╚████║        ║
  ║              ╚══════╝╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝╚═╝  ╚═══╝        ║
  ║                                                                   ║
  ║                    Liturgie Samensteller                          ║
  ║                     PowerPoint Mixer                              ║
  ║                                                                   ║
  ╚═══════════════════════════════════════════════════════════════════╝
        """.strip()

    def _get_system_info(self) -> str:
        """Get system and version information."""
        # Get user info
        try:
            username = getpass.getuser()
        except Exception:
            username = "unknown"

        # Get system info
        try:
            hostname = platform.node()
        except Exception:
            hostname = "unknown"

        lines = [
            "=" * 71,
            f"  Version: {__version__} (revision: {__revision__})",
            f"  Build date: {__build_date__}",
            f"  User: {username}@{hostname}",
            f"  Python: {platform.python_version()}",
            f"  Platform: {platform.system()} {platform.release()}",
            "=" * 71,
        ]

        return "\n".join(lines)
