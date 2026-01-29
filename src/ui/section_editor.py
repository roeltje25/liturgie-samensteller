"""Unified editor dialog for sections with tabs for fields and YouTube."""

import os
import subprocess
import sys
from typing import Dict, List, Optional

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QTabWidget,
    QWidget,
    QTableWidget,
    QTableWidgetItem,
    QLineEdit,
    QPushButton,
    QLabel,
    QDialogButtonBox,
    QHeaderView,
    QComboBox,
    QMessageBox,
    QListWidget,
    QListWidgetItem,
    QProgressBar,
    QSizePolicy,
    QStyledItemDelegate,
    QPlainTextEdit,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QUrl, QModelIndex
from PyQt6.QtMultimedia import QMediaPlayer, QAudioOutput

from ..models import LiturgySection, LiturgySlide
from ..services import PptxService, YouTubeService, YouTubeResult, SlideField
from ..i18n import tr


class MultiLineDelegate(QStyledItemDelegate):
    """Delegate that provides a multiline text editor for table cells."""

    def createEditor(self, parent, option, index):
        """Create a QPlainTextEdit for multiline editing."""
        editor = QPlainTextEdit(parent)
        editor.setMinimumHeight(60)
        return editor

    def setEditorData(self, editor: QPlainTextEdit, index: QModelIndex):
        """Set the editor's text from the model."""
        value = index.model().data(index, Qt.ItemDataRole.EditRole)
        if value:
            editor.setPlainText(str(value))

    def setModelData(self, editor: QPlainTextEdit, model, index: QModelIndex):
        """Save the editor's text to the model."""
        model.setData(index, editor.toPlainText(), Qt.ItemDataRole.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        """Set editor geometry to be larger for multiline editing."""
        rect = option.rect
        # Make the editor taller for comfortable multiline editing
        rect.setHeight(max(100, rect.height()))
        editor.setGeometry(rect)


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


class AudioUrlWorker(QThread):
    """Worker thread to extract audio URL from YouTube."""

    finished = pyqtSignal(str, str)  # url, audio_url
    error = pyqtSignal(str, str)  # url, error_message

    def __init__(self, url: str):
        super().__init__()
        self.url = url

    def run(self):
        try:
            # Use yt-dlp to get the audio stream URL
            result = subprocess.run(
                [sys.executable, "-m", "yt_dlp",
                 "--get-url",
                 "-f", "bestaudio",
                 "--no-warnings",
                 self.url],
                capture_output=True,
                text=True,
                timeout=30,
            )
            if result.returncode == 0 and result.stdout.strip():
                self.finished.emit(self.url, result.stdout.strip())
            else:
                self.error.emit(self.url, result.stderr.strip() or "Failed to get audio URL")
        except Exception as e:
            self.error.emit(self.url, str(e))


class YouTubeItemWidget(QWidget):
    """Custom widget for YouTube list items with play/stop button."""

    play_requested = pyqtSignal(str)  # url
    stop_requested = pyqtSignal()

    def __init__(self, url: str, display_text: str, parent=None):
        super().__init__(parent)
        self.url = url
        self._is_playing = False
        self._is_loading = False

        layout = QHBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 2)

        # Play/Stop button
        self.play_btn = QPushButton("▶")
        self.play_btn.setFixedSize(28, 28)
        self.play_btn.setToolTip("Play preview")
        self.play_btn.clicked.connect(self._on_play_clicked)
        layout.addWidget(self.play_btn)

        # Text label
        self.label = QLabel(display_text)
        self.label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        layout.addWidget(self.label)

    def _on_play_clicked(self):
        if self._is_playing:
            self.stop_requested.emit()
        else:
            self.play_requested.emit(self.url)

    def set_playing(self, playing: bool):
        """Update the button state."""
        self._is_playing = playing
        self._is_loading = False
        if playing:
            self.play_btn.setText("⏹")
            self.play_btn.setToolTip("Stop")
        else:
            self.play_btn.setText("▶")
            self.play_btn.setToolTip("Play preview")

    def set_loading(self, loading: bool):
        """Show loading state."""
        self._is_loading = loading
        if loading:
            self.play_btn.setText("...")
            self.play_btn.setToolTip("Loading...")
            self.play_btn.setEnabled(False)
        else:
            self.play_btn.setEnabled(True)
            if not self._is_playing:
                self.play_btn.setText("▶")
                self.play_btn.setToolTip("Play preview")


class SectionEditorDialog(QDialog):
    """Unified editor dialog for sections with tabs for fields and YouTube."""

    TAB_FIELDS = 0
    TAB_YOUTUBE = 1

    def __init__(
        self,
        section: LiturgySection,
        slide: Optional[LiturgySlide],
        pptx_service: PptxService,
        youtube_service: YouTubeService,
        initial_tab: int = 0,
        parent=None
    ):
        super().__init__(parent)
        self.section = section
        self.slide = slide
        self.pptx_service = pptx_service
        self.youtube_service = youtube_service
        self.initial_tab = initial_tab

        # Fields data
        self._fields: Dict[str, str] = {}
        self._available_fields: List[SlideField] = []

        # YouTube data
        self._youtube_urls: List[str] = list(section.youtube_links) if section.youtube_links else []
        self._worker: Optional[SearchWorker] = None
        self._results: List[YouTubeResult] = []

        # Audio playback
        self._audio_player: Optional[QMediaPlayer] = None
        self._audio_output: Optional[QAudioOutput] = None
        self._audio_worker: Optional[AudioUrlWorker] = None
        self._currently_playing_url: Optional[str] = None
        self._item_widgets: Dict[str, YouTubeItemWidget] = {}  # url -> widget

        self._setup_ui()
        self._setup_audio_player()
        self._connect_signals()
        self._load_data()

        # Set initial tab
        self.tab_widget.setCurrentIndex(initial_tab)

    def _setup_audio_player(self):
        """Setup the audio player."""
        try:
            self._audio_player = QMediaPlayer()
            self._audio_output = QAudioOutput()
            self._audio_player.setAudioOutput(self._audio_output)
            self._audio_output.setVolume(0.7)

            # Connect player signals
            self._audio_player.mediaStatusChanged.connect(self._on_media_status_changed)
            self._audio_player.errorOccurred.connect(self._on_player_error)
        except Exception as e:
            print(f"Warning: Could not initialize audio player: {e}")
            self._audio_player = None
            self._audio_output = None

    def _setup_ui(self) -> None:
        """Setup the dialog UI."""
        title = self.section.name or tr("dialog.editor.title")
        self.setWindowTitle(tr("dialog.editor.title_with_name", name=title))
        self.setMinimumSize(650, 550)
        self.resize(700, 600)

        layout = QVBoxLayout(self)

        # Tab widget
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)

        # Fields tab (always add, even if empty - user might want to add custom fields)
        self.fields_tab = QWidget()
        self._setup_fields_tab()
        self.tab_widget.addTab(self.fields_tab, tr("dialog.editor.tab_fields"))

        # YouTube tab (only for songs)
        if self.section.is_song:
            self.youtube_tab = QWidget()
            self._setup_youtube_tab()
            self.tab_widget.addTab(self.youtube_tab, tr("dialog.editor.tab_youtube"))

        # Button box
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setText(tr("button.save"))
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText(tr("button.cancel"))
        layout.addWidget(self.button_box)

    def _setup_fields_tab(self) -> None:
        """Setup the fields editing tab."""
        layout = QVBoxLayout(self.fields_tab)

        # Section name edit
        section_layout = QHBoxLayout()
        section_layout.addWidget(QLabel(tr("dialog.editor.section_name")))
        self.section_name_edit = QLineEdit()
        self.section_name_edit.setText(self.section.name or "")
        section_layout.addWidget(self.section_name_edit)
        layout.addLayout(section_layout, 0)

        # Slide title edit (if we have a slide)
        if self.slide:
            slide_layout = QHBoxLayout()
            slide_layout.addWidget(QLabel(tr("dialog.editor.slide_title")))
            self.slide_title_edit = QLineEdit()
            self.slide_title_edit.setText(self.slide.title or "")
            slide_layout.addWidget(self.slide_title_edit)
            layout.addLayout(slide_layout, 0)
        else:
            self.slide_title_edit = None

        # Fields table (stretches)
        self.fields_table = QTableWidget()
        self.fields_table.setColumnCount(2)
        self.fields_table.setHorizontalHeaderLabels([
            tr("dialog.fields.field_name"),
            tr("dialog.fields.value")
        ])
        self.fields_table.horizontalHeader().setStretchLastSection(True)
        self.fields_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        # Enable multiline editing for value column
        self._multiline_delegate = MultiLineDelegate(self.fields_table)
        self.fields_table.setItemDelegateForColumn(1, self._multiline_delegate)
        # Enable word wrap and adjust row heights
        self.fields_table.setWordWrap(True)
        self.fields_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.fields_table, 1)

        # Add field controls (no stretch)
        add_layout = QHBoxLayout()
        self.add_field_combo = QComboBox()
        self.add_field_combo.setMinimumWidth(200)
        self.add_field_btn = QPushButton(tr("button.add_field"))
        add_layout.addWidget(QLabel(tr("dialog.fields.add_field")))
        add_layout.addWidget(self.add_field_combo)
        add_layout.addWidget(self.add_field_btn)
        add_layout.addStretch()
        layout.addLayout(add_layout, 0)

        # No fields message (no stretch)
        self.no_fields_label = QLabel(tr("dialog.editor.no_fields"))
        self.no_fields_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.no_fields_label.setVisible(False)
        layout.addWidget(self.no_fields_label, 0)

    def _setup_youtube_tab(self) -> None:
        """Setup the YouTube search/edit tab."""
        layout = QVBoxLayout(self.youtube_tab)

        # Current links section (no stretch for label)
        current_group_label = QLabel(tr("dialog.editor.current_links"))
        current_group_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(current_group_label, 0)

        # Current links list (stretches)
        self.current_links_list = QListWidget()
        self.current_links_list.setMinimumHeight(80)
        layout.addWidget(self.current_links_list, 1)

        # Remove link button (no stretch)
        remove_layout = QHBoxLayout()
        self.remove_link_btn = QPushButton(tr("dialog.editor.remove_link"))
        self.remove_link_btn.setEnabled(False)
        remove_layout.addWidget(self.remove_link_btn)
        remove_layout.addStretch()
        layout.addLayout(remove_layout, 0)

        # Search section (no stretch)
        search_label = QLabel(tr("dialog.youtube.search_for", title=self.section.name))
        layout.addWidget(search_label, 0)

        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setText(self.section.name or "")
        self.search_button = QPushButton(tr("button.search"))
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_button)
        layout.addLayout(search_layout, 0)

        # Progress (no stretch)
        self.progress_label = QLabel(tr("dialog.youtube.searching"))
        self.progress_label.setVisible(False)
        layout.addWidget(self.progress_label, 0)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 0)  # Indeterminate
        layout.addWidget(self.progress_bar, 0)

        # Results section (no stretch for label)
        self.results_label = QLabel(tr("dialog.youtube.select_video"))
        layout.addWidget(self.results_label, 0)

        # Results list (stretches)
        self.results_list = QListWidget()
        self.results_list.setAlternatingRowColors(True)
        self.results_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(self.results_list, 1)

        # No results label (no stretch)
        self.no_results_label = QLabel(tr("dialog.youtube.no_results"))
        self.no_results_label.setVisible(False)
        self.no_results_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.no_results_label, 0)

        # Add selected button (no stretch)
        add_layout = QHBoxLayout()
        self.add_selected_btn = QPushButton(tr("dialog.editor.add_selected"))
        self.add_selected_btn.setEnabled(False)
        add_layout.addWidget(self.add_selected_btn)
        add_layout.addStretch()
        layout.addLayout(add_layout, 0)

        # yt-dlp warning (no stretch)
        self.warning_label = QLabel()
        self.warning_label.setWordWrap(True)
        self.warning_label.setVisible(False)
        layout.addWidget(self.warning_label, 0)

        # Check if yt-dlp is available
        if not self.youtube_service.is_yt_dlp_available():
            self._show_install_option()

    def _connect_signals(self) -> None:
        """Connect widget signals."""
        # Dialog buttons
        self.button_box.accepted.connect(self._on_save)
        self.button_box.rejected.connect(self.reject)

        # Tab change - stop audio
        self.tab_widget.currentChanged.connect(self._on_tab_changed)

        # Fields tab
        self.add_field_btn.clicked.connect(self._on_add_field)

        # YouTube tab (only if it's a song)
        if self.section.is_song:
            self.search_button.clicked.connect(self._do_search)
            self.search_input.returnPressed.connect(self._do_search)
            self.results_list.itemSelectionChanged.connect(self._on_results_selection_changed)
            self.results_list.itemDoubleClicked.connect(self._on_result_double_click)
            self.current_links_list.itemSelectionChanged.connect(self._on_current_links_selection_changed)
            self.remove_link_btn.clicked.connect(self._on_remove_link)
            self.add_selected_btn.clicked.connect(self._on_add_selected)

    def _on_tab_changed(self, index: int):
        """Handle tab change - stop any playing audio."""
        self._stop_playback()

    def _load_data(self) -> None:
        """Load data for all tabs."""
        self._load_fields()
        if self.section.is_song:
            self._load_youtube_links()

    def _load_fields(self) -> None:
        """Load fields from the slide source."""
        if self.slide:
            self._fields = dict(self.slide.fields)

            # Extract available fields from the source PPTX
            if self.slide.source_path and os.path.exists(self.slide.source_path):
                self._available_fields = self.pptx_service.extract_fields(
                    self.slide.source_path,
                    self.slide.slide_index
                )

            # Auto-add text pattern fields
            for field in self._available_fields:
                if field.field_type == "text_pattern" and field.name not in self._fields:
                    self._fields[field.name] = ""

        # Populate combo box
        self.add_field_combo.clear()
        added_names = set(self._fields.keys())
        for field in self._available_fields:
            if field.field_type == "placeholder" and field.name not in added_names:
                self.add_field_combo.addItem(field.name, field)
        self.add_field_combo.addItem(tr("dialog.fields.custom"), None)

        # Populate table
        self._populate_fields_table()

        # Show no fields message if empty
        if not self._fields and not self._available_fields:
            self.no_fields_label.setVisible(True)
            self.fields_table.setVisible(False)

    def _populate_fields_table(self) -> None:
        """Populate the fields table."""
        self.fields_table.setRowCount(len(self._fields))

        for row, (name, value) in enumerate(self._fields.items()):
            name_item = QTableWidgetItem(name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.fields_table.setItem(row, 0, name_item)

            value_item = QTableWidgetItem(value)
            self.fields_table.setItem(row, 1, value_item)

    def _load_youtube_links(self) -> None:
        """Load current YouTube links with play buttons."""
        self.current_links_list.clear()
        self._item_widgets.clear()

        for url in self._youtube_urls:
            self._add_link_item(self.current_links_list, url, url)

    def _add_link_item(self, list_widget: QListWidget, url: str, display_text: str):
        """Add a link item with play button to the list."""
        item = QListWidgetItem()
        widget = YouTubeItemWidget(url, display_text)
        widget.play_requested.connect(self._on_play_requested)
        widget.stop_requested.connect(self._stop_playback)

        item.setSizeHint(widget.sizeHint())
        list_widget.addItem(item)
        list_widget.setItemWidget(item, widget)

        self._item_widgets[url] = widget

    def _add_result_item(self, result: YouTubeResult):
        """Add a search result item with play button."""
        item = QListWidgetItem()
        display_text = f"{result.title}\n  {result.channel} - {result.duration}"
        widget = YouTubeItemWidget(result.url, display_text)
        widget.play_requested.connect(self._on_play_requested)
        widget.stop_requested.connect(self._stop_playback)

        item.setData(Qt.ItemDataRole.UserRole, result)
        item.setSizeHint(widget.sizeHint())
        self.results_list.addItem(item)
        self.results_list.setItemWidget(item, widget)

        self._item_widgets[result.url] = widget

    def _on_play_requested(self, url: str):
        """Handle play request for a URL."""
        if not self._audio_player:
            return  # Audio not available

        # Stop any current playback
        self._stop_playback()

        # Set loading state for the requested URL
        if url in self._item_widgets:
            self._item_widgets[url].set_loading(True)

        # Start worker to get audio URL
        self._audio_worker = AudioUrlWorker(url)
        self._audio_worker.finished.connect(self._on_audio_url_ready)
        self._audio_worker.error.connect(self._on_audio_url_error)
        self._audio_worker.start()

    def _on_audio_url_ready(self, url: str, audio_url: str):
        """Handle when audio URL is extracted."""
        if url in self._item_widgets:
            self._item_widgets[url].set_loading(False)

        if not self._audio_player:
            return

        # Play the audio
        self._currently_playing_url = url
        self._audio_player.setSource(QUrl(audio_url))
        self._audio_player.play()

        # Update button state
        if url in self._item_widgets:
            self._item_widgets[url].set_playing(True)

    def _on_audio_url_error(self, url: str, error: str):
        """Handle audio URL extraction error."""
        if url in self._item_widgets:
            self._item_widgets[url].set_loading(False)
        # Silently fail - user can try again

    def _on_media_status_changed(self, status: QMediaPlayer.MediaStatus):
        """Handle media status changes."""
        if status == QMediaPlayer.MediaStatus.EndOfMedia:
            # Playback finished
            self._stop_playback()

    def _on_player_error(self, error, error_string: str):
        """Handle player errors."""
        self._stop_playback()

    def _stop_playback(self):
        """Stop any current playback."""
        if self._audio_player:
            # Clear the source first to avoid blocking on network streams
            self._audio_player.setSource(QUrl())

        # Reset button state for currently playing URL
        if self._currently_playing_url and self._currently_playing_url in self._item_widgets:
            self._item_widgets[self._currently_playing_url].set_playing(False)

        self._currently_playing_url = None

    def _on_add_field(self) -> None:
        """Add a new field to the table."""
        field = self.add_field_combo.currentData()

        if field is None:
            from PyQt6.QtWidgets import QInputDialog
            name, ok = QInputDialog.getText(
                self,
                tr("dialog.fields.custom_title"),
                tr("dialog.fields.custom_name")
            )
            if not ok or not name.strip():
                return
            field_name = name.strip().upper().replace(" ", "_")
        else:
            field_name = field.name

        if field_name in self._fields:
            QMessageBox.warning(
                self,
                tr("dialog.fields.error"),
                tr("dialog.fields.field_exists", name=field_name)
            )
            return

        self._fields[field_name] = ""

        row = self.fields_table.rowCount()
        self.fields_table.setRowCount(row + 1)

        name_item = QTableWidgetItem(field_name)
        name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.fields_table.setItem(row, 0, name_item)

        value_item = QTableWidgetItem("")
        self.fields_table.setItem(row, 1, value_item)

        idx = self.add_field_combo.findText(field_name)
        if idx >= 0:
            self.add_field_combo.removeItem(idx)

    def _do_search(self) -> None:
        """Perform YouTube search."""
        query = self.search_input.text().strip()
        if not query:
            return

        if not self.youtube_service.is_yt_dlp_available():
            return

        # Stop any playback before clearing results
        self._stop_playback()

        self.progress_label.setVisible(True)
        self.progress_bar.setVisible(True)
        self.no_results_label.setVisible(False)
        self.results_list.clear()
        # Clear widget references for old results
        self._item_widgets = {k: v for k, v in self._item_widgets.items() if k in self._youtube_urls}
        self.search_button.setEnabled(False)

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
            self._add_result_item(result)

    def _on_search_error(self, error: str) -> None:
        """Handle search error."""
        self.progress_label.setVisible(False)
        self.progress_bar.setVisible(False)
        self.search_button.setEnabled(True)

        self.warning_label.setText(f"Search error: {error}")
        self.warning_label.setVisible(True)

    def _on_results_selection_changed(self) -> None:
        """Handle selection change in results list."""
        selected = self.results_list.selectedItems()
        self.add_selected_btn.setEnabled(len(selected) > 0)

    def _on_result_double_click(self, item: QListWidgetItem) -> None:
        """Handle double-click on search result - add it to links."""
        result = item.data(Qt.ItemDataRole.UserRole)
        if result and result.url not in self._youtube_urls:
            self._youtube_urls.append(result.url)
            self._add_link_item(self.current_links_list, result.url, result.url)

    def _on_current_links_selection_changed(self) -> None:
        """Handle selection change in current links list."""
        selected = self.current_links_list.selectedItems()
        self.remove_link_btn.setEnabled(len(selected) > 0)

    def _on_remove_link(self) -> None:
        """Remove selected link from current links."""
        selected = self.current_links_list.selectedItems()
        for item in selected:
            row = self.current_links_list.row(item)
            widget = self.current_links_list.itemWidget(item)
            if widget and hasattr(widget, 'url'):
                url = widget.url
                if url in self._youtube_urls:
                    self._youtube_urls.remove(url)
                if url in self._item_widgets:
                    del self._item_widgets[url]
                # Stop if this was playing
                if url == self._currently_playing_url:
                    self._stop_playback()
            self.current_links_list.takeItem(row)

    def _on_add_selected(self) -> None:
        """Add selected search results to current links."""
        selected = self.results_list.selectedItems()
        for item in selected:
            result = item.data(Qt.ItemDataRole.UserRole)
            if result and result.url not in self._youtube_urls:
                self._youtube_urls.append(result.url)
                self._add_link_item(self.current_links_list, result.url, result.url)

    def _on_save(self) -> None:
        """Save all changes and close."""
        # Stop any playback
        self._stop_playback()

        # Save section name
        new_section_name = self.section_name_edit.text().strip()
        if new_section_name:
            self.section.name = new_section_name

        # Save slide title
        if self.slide and self.slide_title_edit:
            new_slide_title = self.slide_title_edit.text().strip()
            if new_slide_title:
                self.slide.title = new_slide_title

        # Save fields
        if self.slide:
            self._fields.clear()
            for row in range(self.fields_table.rowCount()):
                name = self.fields_table.item(row, 0).text()
                value = self.fields_table.item(row, 1).text()
                if value:
                    self._fields[name] = value
            self.slide.fields = self._fields

        # Save YouTube links
        if self.section.is_song:
            self.section.youtube_links = self._youtube_urls

        self.accept()

    def reject(self):
        """Handle dialog rejection - stop playback."""
        self._stop_playback()
        super().reject()

    def closeEvent(self, event):
        """Handle dialog close - stop playback."""
        self._stop_playback()
        super().closeEvent(event)

    def _show_install_option(self) -> None:
        """Show install option when yt-dlp is not available."""
        self.warning_label.setText(
            "yt-dlp is not installed. YouTube search requires this package."
        )
        self.warning_label.setVisible(True)
        self.search_button.setEnabled(False)
        self.results_label.setVisible(False)
        self.results_list.setVisible(False)

        self.install_button = QPushButton("Install yt-dlp now")
        self.install_button.clicked.connect(self._install_ytdlp)
        layout = self.youtube_tab.layout()
        layout.addWidget(self.install_button)

    def _install_ytdlp(self) -> None:
        """Install yt-dlp package."""
        self.install_button.setEnabled(False)
        self.install_button.setText("Installing...")
        self.warning_label.setText("Installing yt-dlp, please wait...")

        from PyQt6.QtWidgets import QApplication
        QApplication.processEvents()

        try:
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", "yt-dlp"],
                capture_output=True,
                text=True,
                timeout=120,
            )

            if result.returncode == 0:
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

    def get_fields(self) -> Dict[str, str]:
        """Get the edited fields."""
        return self._fields

    def get_youtube_urls(self) -> List[str]:
        """Get the YouTube URLs."""
        return self._youtube_urls
