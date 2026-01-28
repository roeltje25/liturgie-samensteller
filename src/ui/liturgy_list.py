"""Liturgy list widget with drag-and-drop support."""

from typing import List, Optional

from PyQt6.QtWidgets import (
    QListWidget,
    QListWidgetItem,
    QAbstractItemView,
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QIcon

from ..models import Liturgy, LiturgyItem, SongLiturgyItem, GenericLiturgyItem, OfferingLiturgyItem, ItemType
from ..i18n import tr


class LiturgyListWidget(QListWidget):
    """List widget for displaying and reordering liturgy items."""

    order_changed = pyqtSignal()
    item_double_clicked = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._liturgy: Optional[Liturgy] = None
        self._items: List[LiturgyItem] = []

        self._setup_ui()
        self._connect_signals()

    def _setup_ui(self) -> None:
        """Setup the widget appearance and behavior."""
        # Enable drag and drop
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        # Visual settings
        self.setAlternatingRowColors(True)
        self.setSpacing(2)

    def _connect_signals(self) -> None:
        """Connect internal signals."""
        self.model().rowsMoved.connect(self._on_rows_moved)
        self.itemDoubleClicked.connect(self._on_item_double_clicked)

    def set_liturgy(self, liturgy: Liturgy) -> None:
        """Set the liturgy to display."""
        self._liturgy = liturgy
        self._items = list(liturgy.items)
        self._update_display()

    def _update_display(self) -> None:
        """Refresh the list display."""
        self.clear()

        for i, item in enumerate(self._items):
            list_item = QListWidgetItem()
            list_item.setText(f"{i + 1}. {self._format_item(item)}")
            list_item.setData(Qt.ItemDataRole.UserRole, item)

            # Set icon based on item type
            # (icons can be added later if desired)

            self.addItem(list_item)

    def _format_item(self, item: LiturgyItem) -> str:
        """Format an item for display."""
        type_label = tr(f"item.{item.item_type.value}")

        if item.item_type == ItemType.SONG:
            song: SongLiturgyItem = item
            suffix = ""
            if song.is_stub:
                suffix = " (stub)"
            elif not song.pptx_path:
                suffix = f" ({tr('dialog.song.no_pptx')})"
            return f"{type_label}: {item.title}{suffix}"

        elif item.item_type == ItemType.OFFERING:
            offering: OfferingLiturgyItem = item
            if offering.is_stub:
                return f"{type_label}: {offering.title} (stub)"
            return f"{type_label}: {offering.slide_title}"

        else:
            # Generic item
            generic: GenericLiturgyItem = item
            suffix = " (stub)" if generic.is_stub else ""
            return f"{type_label}: {item.title}{suffix}"

    def _on_rows_moved(self) -> None:
        """Handle row reordering via drag and drop."""
        # Rebuild items list from current order
        new_items = []
        for i in range(self.count()):
            list_item = self.item(i)
            item = list_item.data(Qt.ItemDataRole.UserRole)
            new_items.append(item)

        self._items = new_items

        # Update numbering
        self._update_display()

        # Emit signal
        self.order_changed.emit()

    def _on_item_double_clicked(self, item: QListWidgetItem) -> None:
        """Handle double-click on item."""
        self.item_double_clicked.emit()

    def get_elements(self) -> List[LiturgyItem]:
        """Get the current list of items in order. (Backwards compatibility name)"""
        return list(self._items)

    def get_items(self) -> List[LiturgyItem]:
        """Get the current list of items in order."""
        return list(self._items)

    def get_selected_index(self) -> int:
        """Get the index of the currently selected item, or -1 if none."""
        current = self.currentRow()
        return current if current >= 0 else -1

    def select_index(self, index: int) -> None:
        """Select the item at the given index."""
        if 0 <= index < self.count():
            self.setCurrentRow(index)

    def refresh(self) -> None:
        """Refresh display (e.g., after language change)."""
        self._update_display()
