"""Dialog for selecting songs from the Songs folder."""

import re
import unicodedata
from typing import List, Optional

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QTreeWidget,
    QTreeWidgetItem,
    QLineEdit,
    QPushButton,
    QCheckBox,
    QLabel,
    QDialogButtonBox,
    QGroupBox,
    QFileDialog,
    QInputDialog,
    QFrame,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QPixmap

from ..models import Song, SongLiturgyItem
from ..services import PptxService
from ..i18n import tr


def _normalize_for_search(text: str) -> str:
    """Normalize text for fuzzy multilingual search.

    Handles phonetic equivalences common in Dutch/German/English/Arabic songs:
    - y ↔ j (ya/ja, Yeshua/Jeshua)
    - i ↔ ee ↔ ie (Jesus/Jezus)
    - c ↔ k (Christ/Krist)
    - ou ↔ oe ↔ u (Dutch sounds)
    - ei ↔ ij ↔ y (Dutch)
    - ph ↔ f, th ↔ t
    - sch ↔ s
    - Removes diacritics (é→e, ü→u, etc.)
    """
    # Lowercase
    text = text.lower()

    # Remove diacritics (é→e, ü→u, ñ→n, etc.)
    text = unicodedata.normalize('NFD', text)
    text = ''.join(c for c in text if unicodedata.category(c) != 'Mn')

    # Multi-character replacements (order matters - do longer patterns first)
    replacements = [
        ('sch', 's'),      # German/Dutch sch → s
        ('gh', 'g'),       # Arabic transliteration
        ('ch', 'k'),       # Christ → Krist
        ('ph', 'f'),       # Pharao → Farao
        ('th', 't'),       # Thomas → Tomas
        ('oe', 'u'),       # Dutch oe → u sound
        ('ou', 'u'),       # French/Dutch ou → u
        ('ee', 'i'),       # ee → i (Geest → Gist)
        ('ie', 'i'),       # ie → i
        ('ei', 'y'),       # Dutch ei → y
        ('ij', 'y'),       # Dutch ij → y
        ('aa', 'a'),       # Double vowels → single
        ('oo', 'o'),
        ('uu', 'u'),
    ]

    for old, new in replacements:
        text = text.replace(old, new)

    # Single character replacements
    char_map = str.maketrans({
        'j': 'y',          # j → y (Dutch j sounds like English y)
        'c': 'k',          # c → k
        'q': 'k',          # q → k
        'x': 'ks',         # x → ks
        'z': 's',          # z → s (soften)
        'v': 'f',          # v → f (Dutch v often sounds like f)
        'w': 'v',          # w → v (German w)
    })
    text = text.translate(char_map)

    # Remove non-alphanumeric
    text = re.sub(r'[^a-z0-9]', '', text)

    return text


def _fuzzy_match(query: str, text: str) -> float:
    """Calculate fuzzy match score between query and text.

    Returns a score from 0.0 (no match) to 1.0 (perfect match).
    Uses normalized forms for phonetic matching.
    """
    if not query:
        return 1.0

    query_norm = _normalize_for_search(query)
    text_norm = _normalize_for_search(text)

    if not query_norm:
        return 1.0

    # Exact normalized substring match - best score
    if query_norm in text_norm:
        return 1.0

    # Check if all query characters appear in order (subsequence match)
    # This handles typos and partial matches
    query_idx = 0
    matches = 0
    last_match_pos = -1
    consecutive_bonus = 0

    for i, char in enumerate(text_norm):
        if query_idx < len(query_norm) and char == query_norm[query_idx]:
            matches += 1
            # Bonus for consecutive matches
            if last_match_pos == i - 1:
                consecutive_bonus += 0.1
            last_match_pos = i
            query_idx += 1

    if matches == len(query_norm):
        # All characters found in order - good match
        # Score based on how compact the match is
        base_score = 0.6
        # Bonus for consecutive matches
        score = min(1.0, base_score + consecutive_bonus)
        return score

    # Partial match - some characters found
    if matches > 0:
        return 0.3 * (matches / len(query_norm))

    return 0.0


class SongPickerDialog(QDialog):
    """Dialog for selecting a song to add to the liturgy."""

    def __init__(self, songs: List[Song], pptx_service: Optional[PptxService] = None, parent=None):
        super().__init__(parent)
        self.songs = songs
        self.pptx_service = pptx_service
        self._selected_song: Optional[Song] = None
        self._stub_title: Optional[str] = None
        self._external_path: Optional[str] = None
        self._external_title: Optional[str] = None

        self._setup_ui()
        self._populate_tree()
        self._connect_signals()

    def _setup_ui(self) -> None:
        """Setup the dialog UI."""
        self.setWindowTitle(tr("dialog.song.title"))
        self.setMinimumSize(600, 500)
        self.resize(700, 650)

        layout = QVBoxLayout(self)

        # Search field
        search_layout = QHBoxLayout()
        search_label = QLabel(tr("dialog.song.search"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(tr("dialog.song.search"))
        self.search_input.setClearButtonEnabled(True)
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)

        # Song tree
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setAlternatingRowColors(True)
        layout.addWidget(self.tree)

        # Preview area with thumbnail and info
        preview_layout = QHBoxLayout()

        # Thumbnail preview
        self.thumbnail_label = QLabel()
        self.thumbnail_label.setFixedSize(120, 90)
        self.thumbnail_label.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Sunken)
        self.thumbnail_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.thumbnail_label.setStyleSheet("background-color: #f0f0f0;")
        self.thumbnail_label.setText(tr("dialog.song.no_preview"))
        preview_layout.addWidget(self.thumbnail_label)

        # Info label
        self.info_label = QLabel()
        self.info_label.setWordWrap(True)
        preview_layout.addWidget(self.info_label, 1)

        layout.addLayout(preview_layout)

        # Action buttons group
        action_group = QGroupBox(tr("dialog.song.actions"))
        action_layout = QHBoxLayout(action_group)

        self.browse_button = QPushButton(tr("button.browse_file"))
        self.browse_button.setToolTip(tr("dialog.song.browse_tooltip"))
        action_layout.addWidget(self.browse_button)

        self.stub_button = QPushButton(tr("button.create_stub"))
        self.stub_button.setToolTip(tr("dialog.song.stub_tooltip"))
        action_layout.addWidget(self.stub_button)

        action_layout.addStretch()
        layout.addWidget(action_group)

        # Status label for external/stub selection
        self.status_label = QLabel()
        self.status_label.setStyleSheet("color: blue;")
        layout.addWidget(self.status_label)

        # Button box
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setText(tr("button.select"))
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText(tr("button.cancel"))
        layout.addWidget(self.button_box)

        # Initially disable OK button
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setEnabled(False)

    def _populate_tree(self) -> None:
        """Populate the tree with songs organized by folder hierarchy."""
        self.tree.clear()
        self._song_items = {}  # Map from relative_path to QTreeWidgetItem

        # Build tree structure
        folder_items = {}  # Map from folder path to QTreeWidgetItem

        for song in self.songs:
            # Split path into parts
            parts = song.relative_path.replace("\\", "/").split("/")

            # Create folder nodes
            current_path = ""
            parent_item = None

            for i, part in enumerate(parts[:-1]):
                current_path = f"{current_path}/{part}" if current_path else part

                if current_path not in folder_items:
                    folder_item = QTreeWidgetItem()
                    folder_item.setText(0, part)
                    folder_item.setData(0, Qt.ItemDataRole.UserRole, None)  # Not a song

                    if parent_item:
                        parent_item.addChild(folder_item)
                    else:
                        self.tree.addTopLevelItem(folder_item)

                    folder_items[current_path] = folder_item

                parent_item = folder_items[current_path]

            # Create song node
            song_item = QTreeWidgetItem()
            display_text = self._format_song_display(song)
            song_item.setText(0, display_text)
            song_item.setData(0, Qt.ItemDataRole.UserRole, song)

            if parent_item:
                parent_item.addChild(song_item)
            else:
                self.tree.addTopLevelItem(song_item)

            self._song_items[song.relative_path] = song_item

        # Expand all folders
        self.tree.expandAll()

    def _format_song_display(self, song: Song) -> str:
        """Format song display text with status indicators."""
        indicators = []

        if song.has_pptx:
            indicators.append("📊")  # PPT icon
        else:
            indicators.append(f"[{tr('dialog.song.no_pptx')}]")

        if song.has_pdf:
            indicators.append("📕")  # PDF icon

        if song.has_youtube:
            indicators.append("📺")  # YouTube icon

        status = " ".join(indicators)
        # Use display_title which prefers song.properties title over folder name
        return f"{song.display_title}  {status}"

    def _connect_signals(self) -> None:
        """Connect widget signals."""
        self.search_input.textChanged.connect(self._on_search_text_changed)
        self.tree.itemSelectionChanged.connect(self._on_selection_changed)
        self.tree.itemDoubleClicked.connect(self._on_double_click)
        self.browse_button.clicked.connect(self._on_browse_file)
        self.stub_button.clicked.connect(self._on_create_stub)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

    def _on_search_text_changed(self, text: str) -> None:
        """Filter songs based on search text using fuzzy multilingual matching."""
        # Minimum score to consider a match (0.0 to 1.0)
        min_score = 0.3

        def get_match_score(song: Song) -> float:
            """Get best fuzzy match score for a song across all searchable fields."""
            if not text:
                return 1.0

            # Check multiple fields and return best score
            scores = [
                _fuzzy_match(text, song.display_title),
                _fuzzy_match(text, song.name),
                _fuzzy_match(text, song.relative_path),
            ]
            return max(scores)

        def set_item_visibility(item: QTreeWidgetItem) -> bool:
            """Recursively set visibility. Returns True if item or any child is visible."""
            song = item.data(0, Qt.ItemDataRole.UserRole)

            if song is not None:
                # This is a song item - use fuzzy matching
                score = get_match_score(song)
                visible = score >= min_score
                item.setHidden(not visible)
                return visible
            else:
                # This is a folder item
                any_child_visible = False
                for i in range(item.childCount()):
                    child = item.child(i)
                    if set_item_visibility(child):
                        any_child_visible = True

                item.setHidden(not any_child_visible)
                return any_child_visible

        # Process all top-level items
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            set_item_visibility(item)

        # Expand all visible folders
        if text:
            self.tree.expandAll()

    def _on_selection_changed(self) -> None:
        """Handle selection change in tree."""
        selected_items = self.tree.selectedItems()

        if selected_items:
            item = selected_items[0]
            song = item.data(0, Qt.ItemDataRole.UserRole)

            if song is not None:
                self._selected_song = song
                self._stub_title = None
                self._external_path = None
                self._external_title = None
                self._update_info_label(song)
                self.status_label.clear()
            else:
                self._selected_song = None
                self.info_label.setText("")
        else:
            self._selected_song = None
            self.info_label.setText("")

        self._update_ok_button()

    def _update_info_label(self, song: Song) -> None:
        """Update the info label and thumbnail with song details."""
        info_parts = []

        if song.has_pptx:
            info_parts.append("PowerPoint: ✓")
        else:
            info_parts.append("PowerPoint: ✗")

        if song.has_pdf:
            info_parts.append("PDF: ✓")
        else:
            info_parts.append("PDF: ✗")

        if song.has_youtube:
            info_parts.append(f"YouTube: {len(song.youtube_links)} link(s)")
        else:
            info_parts.append("YouTube: -")

        self.info_label.setText("  |  ".join(info_parts))

        # Update thumbnail
        self._update_thumbnail(song)

    def _update_thumbnail(self, song: Song) -> None:
        """Update the thumbnail preview for a song."""
        if not self.pptx_service or not song.has_pptx:
            self.thumbnail_label.setPixmap(QPixmap())
            self.thumbnail_label.setText(tr("dialog.song.no_preview"))
            return

        # Try to get thumbnail from pptx
        thumb_data = self.pptx_service.get_thumbnail(song.pptx_path)
        if thumb_data:
            pixmap = QPixmap()
            pixmap.loadFromData(thumb_data)
            if not pixmap.isNull():
                scaled = pixmap.scaled(
                    self.thumbnail_label.size(),
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation
                )
                self.thumbnail_label.setPixmap(scaled)
                self.thumbnail_label.setText("")
                return

        self.thumbnail_label.setPixmap(QPixmap())
        self.thumbnail_label.setText(tr("dialog.song.no_preview"))

    def _on_double_click(self, item: QTreeWidgetItem, column: int) -> None:
        """Handle double-click on item."""
        song = item.data(0, Qt.ItemDataRole.UserRole)
        if song is not None:
            self._selected_song = song
            self._stub_title = None
            self._external_path = None
            self._external_title = None
            self.accept()

    def _on_browse_file(self) -> None:
        """Open file dialog to select external PowerPoint file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("dialog.song.browse_title"),
            "",
            "PowerPoint Files (*.pptx *.ppt);;All Files (*)"
        )

        if file_path:
            # Ask for a title for this external song
            import os
            default_title = os.path.splitext(os.path.basename(file_path))[0]

            title, ok = QInputDialog.getText(
                self,
                tr("dialog.song.external_title_dialog"),
                tr("dialog.song.external_enter_title"),
                text=default_title
            )

            if ok and title.strip():
                self._external_path = file_path
                self._external_title = title.strip()
                self._selected_song = None
                self._stub_title = None
                self.tree.clearSelection()

                self.status_label.setText(tr("dialog.song.external_selected", title=self._external_title))
                self.info_label.clear()
                self._update_ok_button()

    def _on_create_stub(self) -> None:
        """Open dialog to create a stub with custom title."""
        title, ok = QInputDialog.getText(
            self,
            tr("dialog.song.stub_dialog_title"),
            tr("dialog.song.stub_enter_title"),
        )

        if ok and title.strip():
            self._stub_title = title.strip()
            self._selected_song = None
            self._external_path = None
            self._external_title = None
            self.tree.clearSelection()

            self.status_label.setText(tr("dialog.song.stub_selected", title=self._stub_title))
            self.info_label.clear()
            self._update_ok_button()

    def _update_ok_button(self) -> None:
        """Update OK button enabled state."""
        enabled = (
            self._selected_song is not None
            or self._stub_title is not None
            or self._external_path is not None
        )
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setEnabled(enabled)

    def get_selected_item(self) -> Optional[SongLiturgyItem]:
        """Get the selected song as a SongLiturgyItem."""
        if self._stub_title:
            return SongLiturgyItem(
                title=self._stub_title,
                is_stub=True,
            )

        if self._external_path:
            return SongLiturgyItem(
                title=self._external_title,
                source_path=self._external_path,
                pptx_path=self._external_path,
                is_stub=False,
            )

        if self._selected_song:
            return SongLiturgyItem(
                title=self._selected_song.display_title,
                source_path=self._selected_song.relative_path,
                pptx_path=self._selected_song.pptx_path,
                pdf_path=self._selected_song.pdf_path,
                youtube_links=list(self._selected_song.youtube_links),
                is_stub=not self._selected_song.has_pptx,
            )

        return None

    # Backwards compatibility alias
    def get_selected_element(self) -> Optional[SongLiturgyItem]:
        """Get the selected item. (Backwards compatibility alias)"""
        return self.get_selected_item()
