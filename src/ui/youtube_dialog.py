"""Dialog for searching and selecting YouTube videos."""

import subprocess
import sys
from typing import List, Optional

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QPushButton,
    QListWidget,
    QListWidgetItem,
    QLabel,
    QDialogButtonBox,
    QProgressBar,
    QMessageBox,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal

from ..services import YouTubeService, YouTubeResult
from ..i18n import tr


class SearchWorker(QThread):
    """Worker thread for YouTube search."""

    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, youtube_service: YouTubeService, query: str):
        super().__init__()
        self.youtube_service = youtube_service
        self.query = query

    def run(self):
        try:
            results = self.youtube_service.search(self.query, max_results=10)
            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))


class YouTubeDialog(QDialog):
    """Dialog for searching and selecting YouTube videos for a song."""

    def __init__(self, song_title: str, youtube_service: YouTubeService, parent=None):
        super().__init__(parent)
        self.song_title = song_title
        self.youtube_service = youtube_service
        self._selected_urls: List[str] = []
        self._worker: Optional[SearchWorker] = None
        self._results: List[YouTubeResult] = []

        self._setup_ui()
        self._connect_signals()

        # Auto-search on open
        self._do_search()

    def _setup_ui(self) -> None:
        """Setup the dialog UI."""
        self.setWindowTitle(tr("dialog.youtube.title"))
        self.setMinimumSize(600, 500)
        self.resize(700, 600)

        layout = QVBoxLayout(self)

        # Search section
        search_label = QLabel(tr("dialog.youtube.search_for", title=self.song_title))
        layout.addWidget(search_label)

        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setText(self.song_title)
        self.search_button = QPushButton(tr("button.search"))
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_button)
        layout.addLayout(search_layout)

        # Progress
        self.progress_label = QLabel(tr("dialog.youtube.searching"))
        self.progress_label.setVisible(False)
        layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 0)  # Indeterminate
        layout.addWidget(self.progress_bar)

        # Results section
        self.results_label = QLabel(tr("dialog.youtube.select_video"))
        layout.addWidget(self.results_label)

        self.results_list = QListWidget()
        self.results_list.setAlternatingRowColors(True)
        self.results_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(self.results_list)

        # No results label
        self.no_results_label = QLabel(tr("dialog.youtube.no_results"))
        self.no_results_label.setVisible(False)
        self.no_results_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.no_results_label)

        # yt-dlp warning
        self.warning_label = QLabel()
        self.warning_label.setWordWrap(True)
        self.warning_label.setVisible(False)
        layout.addWidget(self.warning_label)

        # Button box
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.save_button = self.button_box.button(QDialogButtonBox.StandardButton.Ok)
        self.save_button.setText(tr("dialog.youtube.save_link"))
        self.save_button.setEnabled(False)
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText(tr("button.cancel"))
        layout.addWidget(self.button_box)

        # Check if yt-dlp is available
        if not self.youtube_service.is_yt_dlp_available():
            self._show_install_option()

    def _connect_signals(self) -> None:
        """Connect widget signals."""
        self.search_button.clicked.connect(self._do_search)
        self.search_input.returnPressed.connect(self._do_search)
        self.results_list.itemSelectionChanged.connect(self._on_selection_changed)
        self.results_list.itemDoubleClicked.connect(self._on_double_click)
        self.button_box.accepted.connect(self._on_accept)
        self.button_box.rejected.connect(self.reject)

    def _do_search(self) -> None:
        """Perform YouTube search."""
        query = self.search_input.text().strip()
        if not query:
            return

        if not self.youtube_service.is_yt_dlp_available():
            return

        # Show progress
        self.progress_label.setVisible(True)
        self.progress_bar.setVisible(True)
        self.no_results_label.setVisible(False)
        self.results_list.clear()
        self.search_button.setEnabled(False)

        # Start search worker
        self._worker = SearchWorker(self.youtube_service, query)
        self._worker.finished.connect(self._on_search_finished)
        self._worker.error.connect(self._on_search_error)
        self._worker.start()

    def _on_search_finished(self, results: List[YouTubeResult]) -> None:
        """Handle search completion."""
        self.progress_label.setVisible(False)
        self.progress_bar.setVisible(False)
        self.search_button.setEnabled(True)

        self._results = results
        self.results_list.clear()

        if not results:
            self.no_results_label.setVisible(True)
            return

        for result in results:
            item = QListWidgetItem()
            item.setText(f"{result.title}\n  {result.channel} - {result.duration}")
            item.setData(Qt.ItemDataRole.UserRole, result)
            self.results_list.addItem(item)

    def _on_search_error(self, error: str) -> None:
        """Handle search error."""
        self.progress_label.setVisible(False)
        self.progress_bar.setVisible(False)
        self.search_button.setEnabled(True)

        self.warning_label.setText(f"Search error: {error}")
        self.warning_label.setVisible(True)

    def _on_selection_changed(self) -> None:
        """Handle selection change in results list."""
        selected = self.results_list.selectedItems()
        self.save_button.setEnabled(len(selected) > 0)

    def _on_double_click(self, item: QListWidgetItem) -> None:
        """Handle double-click on result."""
        result = item.data(Qt.ItemDataRole.UserRole)
        if result:
            self._selected_urls = [result.url]
            self.accept()

    def _on_accept(self) -> None:
        """Handle dialog acceptance."""
        selected_items = self.results_list.selectedItems()
        self._selected_urls = []

        for item in selected_items:
            result = item.data(Qt.ItemDataRole.UserRole)
            if result:
                self._selected_urls.append(result.url)

        self.accept()

    def get_selected_urls(self) -> List[str]:
        """Get the selected YouTube URLs."""
        return self._selected_urls

    def _show_install_option(self) -> None:
        """Show install option when yt-dlp is not available."""
        self.warning_label.setText(
            "yt-dlp is not installed. YouTube search requires this package."
        )
        self.warning_label.setVisible(True)
        self.search_button.setEnabled(False)
        self.results_label.setVisible(False)
        self.results_list.setVisible(False)

        # Add install button
        self.install_button = QPushButton("Install yt-dlp now")
        self.install_button.clicked.connect(self._install_ytdlp)
        # Insert before button box
        layout = self.layout()
        layout.insertWidget(layout.count() - 1, self.install_button)

    def _install_ytdlp(self) -> None:
        """Install yt-dlp package."""
        self.install_button.setEnabled(False)
        self.install_button.setText("Installing...")
        self.warning_label.setText("Installing yt-dlp, please wait...")

        # Force UI update
        from PyQt6.QtWidgets import QApplication
        QApplication.processEvents()

        try:
            # Install yt-dlp using pip
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", "yt-dlp"],
                capture_output=True,
                text=True,
                timeout=120,
            )

            if result.returncode == 0:
                # Reset the yt-dlp availability check
                self.youtube_service._yt_dlp_available = None

                if self.youtube_service.is_yt_dlp_available():
                    self.warning_label.setText("yt-dlp installed successfully!")
                    self.install_button.setVisible(False)
                    self.search_button.setEnabled(True)
                    self.results_label.setVisible(True)
                    self.results_list.setVisible(True)

                    QMessageBox.information(
                        self,
                        "Installation Complete",
                        "yt-dlp has been installed. You can now search for YouTube videos."
                    )
                else:
                    self.warning_label.setText(
                        "Installation completed but yt-dlp is still not available. "
                        "Please restart the application."
                    )
                    self.install_button.setText("Installed - Restart app")
            else:
                self.warning_label.setText(
                    f"Installation failed: {result.stderr[:200] if result.stderr else 'Unknown error'}"
                )
                self.install_button.setEnabled(True)
                self.install_button.setText("Retry installation")

        except subprocess.TimeoutExpired:
            self.warning_label.setText("Installation timed out. Please try again or install manually.")
            self.install_button.setEnabled(True)
            self.install_button.setText("Retry installation")
        except Exception as e:
            self.warning_label.setText(f"Installation error: {str(e)}")
            self.install_button.setEnabled(True)
            self.install_button.setText("Retry installation")
