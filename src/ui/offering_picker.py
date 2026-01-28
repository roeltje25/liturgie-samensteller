"""Dialog for selecting a slide from the Offering (Collecte) PowerPoint."""

from typing import List, Optional

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QLineEdit,
    QPushButton,
    QLabel,
    QDialogButtonBox,
    QTextEdit,
    QGroupBox,
    QFileDialog,
    QInputDialog,
    QFrame,
)
from PyQt6.QtCore import Qt, QRunnable, QThreadPool, pyqtSignal, QObject
from PyQt6.QtGui import QPixmap

from ..models import OfferingSlide, OfferingLiturgyItem, Settings
from ..services import PptxService, FolderScanner
from ..i18n import tr


class ThumbnailSignals(QObject):
    """Signals for thumbnail worker."""
    finished = pyqtSignal(int, bytes)  # slide_index, image_data
    error = pyqtSignal(int)  # slide_index


class ThumbnailRunnable(QRunnable):
    """Runnable for loading slide thumbnails in background."""

    def __init__(self, pptx_service: PptxService, pptx_path: str, slide_index: int):
        super().__init__()
        self.pptx_service = pptx_service
        self.pptx_path = pptx_path
        self.slide_index = slide_index
        self.signals = ThumbnailSignals()
        self._cancelled = False

    def cancel(self):
        """Mark this runnable as cancelled."""
        self._cancelled = True

    def run(self):
        """Load the thumbnail in background."""
        if self._cancelled:
            return
        try:
            thumb_data = self.pptx_service.get_slide_thumbnail(
                self.pptx_path, self.slide_index
            )
            if self._cancelled:
                return
            if thumb_data:
                self.signals.finished.emit(self.slide_index, thumb_data)
            else:
                self.signals.error.emit(self.slide_index)
        except Exception:
            if not self._cancelled:
                self.signals.error.emit(self.slide_index)


class OfferingPickerDialog(QDialog):
    """Dialog for selecting an Offering slide to add to the liturgy."""

    def __init__(self, slides: List[OfferingSlide], settings: Settings, base_path: str,
                 folder_scanner: Optional['FolderScanner'] = None, parent=None):
        super().__init__(parent)
        self.slides = slides
        self.settings = settings
        self.base_path = base_path
        self.folder_scanner = folder_scanner
        self._selected_slide: Optional[OfferingSlide] = None
        self._pptx_service = PptxService(settings, base_path)
        self._custom_pptx_path: Optional[str] = None
        self._stub_title: Optional[str] = None

        # Background thumbnail loading
        self._thread_pool = QThreadPool.globalInstance()
        self._current_runnable: Optional[ThumbnailRunnable] = None
        self._loading_slide_index: Optional[int] = None

        self._setup_ui()
        self._populate_list()
        self._connect_signals()

    def _setup_ui(self) -> None:
        """Setup the dialog UI."""
        self.setWindowTitle(tr("dialog.offering.title"))
        self.setMinimumSize(550, 550)
        self.resize(650, 650)

        layout = QVBoxLayout(self)

        # Instructions label
        self.instruction_label = QLabel(tr("dialog.offering.select_slide"))
        layout.addWidget(self.instruction_label)

        # Search field
        search_layout = QHBoxLayout()
        search_label = QLabel(tr("dialog.offering.search"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(tr("dialog.offering.search"))
        self.search_input.setClearButtonEnabled(True)
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)

        # Slide list
        self.list_widget = QListWidget()
        self.list_widget.setAlternatingRowColors(True)
        layout.addWidget(self.list_widget)

        # Preview area with thumbnail
        preview_label = QLabel(tr("dialog.offering.preview"))
        layout.addWidget(preview_label)

        preview_layout = QHBoxLayout()

        # Thumbnail
        self.thumbnail_label = QLabel()
        self.thumbnail_label.setFixedSize(120, 90)
        self.thumbnail_label.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Sunken)
        self.thumbnail_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.thumbnail_label.setStyleSheet("background-color: #f0f0f0;")
        preview_layout.addWidget(self.thumbnail_label)

        # Text preview
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMaximumHeight(100)
        preview_layout.addWidget(self.preview_text, 1)

        layout.addLayout(preview_layout)

        # Action buttons group
        action_group = QGroupBox(tr("dialog.offering.actions"))
        action_layout = QHBoxLayout(action_group)

        self.browse_button = QPushButton(tr("button.browse_file"))
        self.browse_button.setToolTip(tr("dialog.offering.browse_tooltip"))
        action_layout.addWidget(self.browse_button)

        self.stub_button = QPushButton(tr("button.create_stub"))
        self.stub_button.setToolTip(tr("dialog.offering.stub_tooltip"))
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

    def _populate_list(self) -> None:
        """Populate the list with Offering slides."""
        self.list_widget.clear()

        for slide in self.slides:
            list_item = QListWidgetItem()
            list_item.setText(f"{slide.index + 1}. {slide.title}")
            list_item.setData(Qt.ItemDataRole.UserRole, slide)
            self.list_widget.addItem(list_item)

    def _connect_signals(self) -> None:
        """Connect widget signals."""
        self.search_input.textChanged.connect(self._on_search_text_changed)
        self.list_widget.itemSelectionChanged.connect(self._on_selection_changed)
        self.list_widget.itemDoubleClicked.connect(self._on_double_click)
        self.browse_button.clicked.connect(self._on_browse_file)
        self.stub_button.clicked.connect(self._on_create_stub)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

    def _on_search_text_changed(self, text: str) -> None:
        """Filter list items based on search text."""
        search_lower = text.lower()
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            matches = search_lower in item.text().lower()
            item.setHidden(not matches)

    def _on_selection_changed(self) -> None:
        """Handle selection change in list."""
        selected_items = self.list_widget.selectedItems()

        if selected_items:
            item = selected_items[0]
            self._selected_slide = item.data(Qt.ItemDataRole.UserRole)
            self._custom_pptx_path = None
            self._stub_title = None
            self._update_preview()
            self.status_label.clear()
        else:
            self._selected_slide = None
            self.preview_text.clear()

        self._update_ok_button()

    def _update_preview(self) -> None:
        """Update preview with slide content and thumbnail."""
        if self._selected_slide:
            pptx_path = self._custom_pptx_path or self.settings.get_collecte_path(self.base_path)
            preview_text = self._pptx_service.get_slide_thumbnail_text(
                pptx_path, self._selected_slide.index
            )
            self.preview_text.setText(preview_text if preview_text else tr("dialog.offering.no_text"))

            # Update thumbnail
            self._update_thumbnail(pptx_path)
        else:
            self.thumbnail_label.setPixmap(QPixmap())

    def _update_thumbnail(self, pptx_path: str) -> None:
        """Update the thumbnail preview for the selected slide (async)."""
        if not self._selected_slide:
            self.thumbnail_label.setPixmap(QPixmap())
            return

        slide_index = self._selected_slide.index

        # Cancel any existing thumbnail load
        self._cancel_thumbnail_load()

        # Show loading indicator
        self.thumbnail_label.setText("...")
        self._loading_slide_index = slide_index

        # Start background loading using thread pool
        runnable = ThumbnailRunnable(self._pptx_service, pptx_path, slide_index)
        runnable.signals.finished.connect(self._on_thumbnail_loaded)
        runnable.signals.error.connect(self._on_thumbnail_error)
        self._current_runnable = runnable
        self._thread_pool.start(runnable)

    def _cancel_thumbnail_load(self) -> None:
        """Cancel any in-progress thumbnail loading."""
        if self._current_runnable:
            self._current_runnable.cancel()
        self._current_runnable = None
        self._loading_slide_index = None

    def _on_thumbnail_loaded(self, slide_index: int, thumb_data: bytes) -> None:
        """Handle thumbnail loaded from background thread."""
        # Only update if this is still the slide we're waiting for
        if self._loading_slide_index != slide_index:
            return

        pixmap = QPixmap()
        pixmap.loadFromData(thumb_data)
        if not pixmap.isNull():
            scaled = pixmap.scaled(
                self.thumbnail_label.size(),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
            self.thumbnail_label.setPixmap(scaled)
        else:
            self.thumbnail_label.setPixmap(QPixmap())

        self._loading_slide_index = None

    def _on_thumbnail_error(self, slide_index: int) -> None:
        """Handle thumbnail load error."""
        if self._loading_slide_index == slide_index:
            self.thumbnail_label.setPixmap(QPixmap())
            self._loading_slide_index = None

    def _on_double_click(self, item: QListWidgetItem) -> None:
        """Handle double-click on item."""
        self._selected_slide = item.data(Qt.ItemDataRole.UserRole)
        self._custom_pptx_path = None
        self._stub_title = None
        self.accept()

    def _on_browse_file(self) -> None:
        """Open file dialog to select different offerings PowerPoint file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("dialog.offering.browse_title"),
            "",
            "PowerPoint Files (*.pptx *.ppt);;All Files (*)"
        )

        if file_path:
            self._custom_pptx_path = file_path
            self._selected_slide = None
            self._stub_title = None
            self.list_widget.clearSelection()

            # Load slides from the new file
            if self.folder_scanner:
                new_slides = self.folder_scanner.get_offering_slides(file_path)
            else:
                # Fallback: create temporary scanner
                from ..services import FolderScanner
                temp_scanner = FolderScanner(self.settings, self.base_path)
                new_slides = temp_scanner.get_offering_slides(file_path)

            if new_slides:
                self.slides = new_slides
                self._populate_list()

                import os
                filename = os.path.basename(file_path)
                self.status_label.setText(tr("dialog.offering.external_loaded", filename=filename))
            else:
                self.status_label.setText(tr("dialog.offering.no_slides_found"))

            self.preview_text.clear()
            self._update_ok_button()

    def _on_create_stub(self) -> None:
        """Open dialog to create a stub offering slide."""
        title, ok = QInputDialog.getText(
            self,
            tr("dialog.offering.stub_dialog_title"),
            tr("dialog.offering.stub_enter_title"),
        )

        if ok and title.strip():
            self._stub_title = title.strip()
            self._selected_slide = None
            self._custom_pptx_path = None
            self.list_widget.clearSelection()

            self.status_label.setText(tr("dialog.offering.stub_selected", title=self._stub_title))
            self.preview_text.clear()
            self._update_ok_button()

    def _update_ok_button(self) -> None:
        """Update OK button enabled state."""
        enabled = (
            self._selected_slide is not None
            or self._stub_title is not None
        )
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setEnabled(enabled)

    def get_selected_item(self) -> Optional[OfferingLiturgyItem]:
        """Get the selected slide as an OfferingLiturgyItem."""
        if self._stub_title:
            return OfferingLiturgyItem(
                title=self._stub_title,
                is_stub=True,
            )

        if self._selected_slide:
            return OfferingLiturgyItem(
                title=tr("item.offering"),
                slide_index=self._selected_slide.index,
                slide_title=self._selected_slide.title,
                pptx_path=self._custom_pptx_path,  # None if using default
                is_stub=False,
            )

        return None

    # Backwards compatibility alias
    def get_selected_element(self) -> Optional[OfferingLiturgyItem]:
        """Get the selected item. (Backwards compatibility alias)"""
        return self.get_selected_item()

    def closeEvent(self, event) -> None:
        """Clean up background thread on close."""
        self._cancel_thumbnail_load()
        super().closeEvent(event)


# Backwards compatibility alias
CollectePickerDialog = OfferingPickerDialog
