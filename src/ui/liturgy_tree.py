"""Liturgy tree widget with hierarchical section/slide support and constrained drag-drop."""

from typing import List, Optional, Tuple

from PyQt6.QtWidgets import (
    QTreeWidget,
    QTreeWidgetItem,
    QAbstractItemView,
    QMenu,
)
from PyQt6.QtCore import Qt, pyqtSignal, QMimeData
from PyQt6.QtGui import QDragEnterEvent, QDropEvent

import os

from ..models import (
    Liturgy,
    LiturgySection,
    LiturgySlide,
    SectionType,
    # V1 compatibility
    LiturgyItem,
    SongLiturgyItem,
    GenericLiturgyItem,
    OfferingLiturgyItem,
    ItemType,
)
from ..services import PptxService
from ..i18n import tr


class LiturgyTreeWidget(QTreeWidget):
    """Tree widget for displaying and reordering liturgy sections and slides."""

    # Signals
    order_changed = pyqtSignal()
    section_selected = pyqtSignal(str)  # section_id
    slide_selected = pyqtSignal(str, str)  # section_id, slide_id
    item_double_clicked = pyqtSignal()

    # Item types for data storage
    ITEM_TYPE_SECTION = 1
    ITEM_TYPE_SLIDE = 2

    def __init__(self, pptx_service: PptxService = None, parent=None):
        super().__init__(parent)
        self._liturgy: Optional[Liturgy] = None
        self._pptx_service = pptx_service
        # Cache for unfilled fields check: {(source_path, slide_index): [field_names]}
        self._field_cache: dict = {}

        self._setup_ui()
        self._connect_signals()

    def set_pptx_service(self, pptx_service: PptxService) -> None:
        """Set the PPTX service for field checking."""
        self._pptx_service = pptx_service
        self._field_cache.clear()

    def _get_field_status(self, slide: LiturgySlide) -> Tuple[List[str], List[str]]:
        """Get field status for a slide.
        Returns (all_fields, unfilled_fields) tuple.
        """
        if not self._pptx_service or not slide.source_path:
            return ([], [])

        if not os.path.exists(slide.source_path):
            return ([], [])

        # Check cache
        cache_key = (slide.source_path, slide.slide_index)
        if cache_key not in self._field_cache:
            try:
                fields = self._pptx_service.extract_fields(
                    slide.source_path, slide.slide_index
                )
                # Only track text_pattern fields
                self._field_cache[cache_key] = [
                    f.name for f in fields if f.field_type == "text_pattern"
                ]
            except Exception:
                self._field_cache[cache_key] = []

        required_fields = self._field_cache[cache_key]
        unfilled = []
        for field_name in required_fields:
            value = slide.fields.get(field_name, "")
            if not value.strip():
                unfilled.append(field_name)

        return (required_fields, unfilled)

    def _get_unfilled_fields(self, slide: LiturgySlide) -> List[str]:
        """Get list of unfilled text pattern fields for a slide."""
        _, unfilled = self._get_field_status(slide)
        return unfilled

    @staticmethod
    def _clean_title(title: str) -> str:
        """Clean a title by removing newlines and control characters."""
        if not title:
            return title
        # Replace various newline and control characters with spaces
        cleaned = title.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
        cleaned = ''.join(c if c.isprintable() or c == ' ' else ' ' for c in cleaned)
        return ' '.join(cleaned.split()).strip()

    def _setup_ui(self) -> None:
        """Setup the widget appearance and behavior."""
        # Single column for display
        self.setHeaderHidden(True)
        self.setColumnCount(1)

        # Enable drag and drop
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        # Visual settings
        self.setAlternatingRowColors(True)
        self.setIndentation(20)
        self.setRootIsDecorated(True)

        # Enable context menu
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)

    def _connect_signals(self) -> None:
        """Connect internal signals."""
        self.itemSelectionChanged.connect(self._on_selection_changed)
        self.itemDoubleClicked.connect(self._on_item_double_clicked)
        self.customContextMenuRequested.connect(self._on_context_menu)

    def set_liturgy(self, liturgy: Liturgy) -> None:
        """Set the liturgy to display."""
        self._liturgy = liturgy
        self._update_display()

    def _update_display(self) -> None:
        """Refresh the tree display."""
        self.clear()

        if not self._liturgy:
            return

        # Check if using v2 (sections) or v1 (items)
        if self._liturgy.sections:
            self._display_sections()
        else:
            self._display_items_as_sections()

    def _display_sections(self) -> None:
        """Display v2 sections and their slides."""
        for section_idx, section in enumerate(self._liturgy.sections):
            section_item = self._create_section_item(section, section_idx)
            self.addTopLevelItem(section_item)

            # Add slides as children
            for slide_idx, slide in enumerate(section.slides):
                slide_item = self._create_slide_item(slide, section.id, slide_idx, len(section.slides))
                section_item.addChild(slide_item)

            # Expand by default
            section_item.setExpanded(True)

    def _display_items_as_sections(self) -> None:
        """Display v1 items as pseudo-sections (for backwards compatibility)."""
        for item_idx, item in enumerate(self._liturgy._items):
            # Create a pseudo-section for each item
            section_item = self._create_item_as_section(item, item_idx)
            self.addTopLevelItem(section_item)

    def _create_section_item(self, section: LiturgySection, index: int) -> QTreeWidgetItem:
        """Create a tree item for a section."""
        item = QTreeWidgetItem()

        # Format display text
        icon = "🎵" if section.is_song else "📁"
        indicators = []
        if section.is_song:
            if section.has_pptx:
                indicators.append("📊")  # PPT icon
            if section.has_youtube:
                indicators.append("📺")  # YouTube icon
            if section.has_pdf:
                indicators.append("📕")  # PDF icon

        indicator_text = " ".join(indicators)
        clean_name = self._clean_title(section.name)
        display_text = f"{index + 1}. {icon} {clean_name}"
        if indicator_text:
            display_text += f"  {indicator_text}"

        item.setText(0, display_text)
        item.setData(0, Qt.ItemDataRole.UserRole, self.ITEM_TYPE_SECTION)
        item.setData(0, Qt.ItemDataRole.UserRole + 1, section.id)

        # Style: bold for sections
        font = item.font(0)
        font.setBold(True)
        item.setFont(0, font)

        # Allow children (slides) but not dropping of other sections
        item.setFlags(
            Qt.ItemFlag.ItemIsEnabled |
            Qt.ItemFlag.ItemIsSelectable |
            Qt.ItemFlag.ItemIsDragEnabled |
            Qt.ItemFlag.ItemIsDropEnabled
        )

        return item

    def _create_slide_item(
        self, slide: LiturgySlide, section_id: str, index: int, total: int
    ) -> QTreeWidgetItem:
        """Create a tree item for a slide."""
        item = QTreeWidgetItem()

        # Format display text with tree-like prefix
        prefix = "└─" if index == total - 1 else "├─"
        clean_title = self._clean_title(slide.title) or f'Slide {index + 1}'
        display_text = f"{prefix} {clean_title}"
        if slide.is_stub:
            display_text += " (stub)"

        # Check for field status
        all_fields, unfilled = self._get_field_status(slide)
        if all_fields:  # Has text pattern fields
            if unfilled:
                display_text += " ⚠"  # Warning: unfilled fields
                item.setToolTip(0, f"Unfilled fields: {', '.join(unfilled)}")
            else:
                display_text += " ✓"  # All fields filled
                item.setToolTip(0, f"All fields filled: {', '.join(all_fields)}")

        item.setText(0, display_text)
        item.setData(0, Qt.ItemDataRole.UserRole, self.ITEM_TYPE_SLIDE)
        item.setData(0, Qt.ItemDataRole.UserRole + 1, section_id)
        item.setData(0, Qt.ItemDataRole.UserRole + 2, slide.id)

        # Store unfilled status for easy access
        item.setData(0, Qt.ItemDataRole.UserRole + 4, unfilled)

        # Slides can be dragged
        item.setFlags(
            Qt.ItemFlag.ItemIsEnabled |
            Qt.ItemFlag.ItemIsSelectable |
            Qt.ItemFlag.ItemIsDragEnabled
        )

        return item

    def _create_item_as_section(self, item: LiturgyItem, index: int) -> QTreeWidgetItem:
        """Create a tree item for a v1 item (displayed as a section)."""
        tree_item = QTreeWidgetItem()

        # Format display text based on item type
        type_label = tr(f"item.{item.item_type.value}")

        if item.item_type == ItemType.SONG:
            song: SongLiturgyItem = item
            icon = "🎵"
            indicators = []
            if song.pptx_path:
                indicators.append("📊")  # PPT icon
            if song.youtube_links:
                indicators.append("📺")  # YouTube icon
            if song.pdf_path:
                indicators.append("📕")  # PDF icon
            suffix = ""
            if song.is_stub:
                suffix = " (stub)"
            elif not song.pptx_path:
                suffix = f" ({tr('dialog.song.no_pptx')})"

            indicator_text = " ".join(indicators)
            display_text = f"{index + 1}. {icon} {item.title}{suffix}"
            if indicator_text:
                display_text += f"  {indicator_text}"

        elif item.item_type == ItemType.OFFERING:
            offering: OfferingLiturgyItem = item
            icon = "💰"
            if offering.is_stub:
                display_text = f"{index + 1}. {icon} {offering.title} (stub)"
            else:
                display_text = f"{index + 1}. {icon} {offering.slide_title or offering.title}"

        else:
            # Generic item
            generic: GenericLiturgyItem = item
            icon = "📁"
            suffix = " (stub)" if generic.is_stub else ""
            display_text = f"{index + 1}. {icon} {item.title}{suffix}"

        tree_item.setText(0, display_text)
        tree_item.setData(0, Qt.ItemDataRole.UserRole, self.ITEM_TYPE_SECTION)
        tree_item.setData(0, Qt.ItemDataRole.UserRole + 1, str(index))  # Use index as pseudo-ID
        tree_item.setData(0, Qt.ItemDataRole.UserRole + 3, item)  # Store original item

        # Style: bold for items
        font = tree_item.font(0)
        font.setBold(True)
        tree_item.setFont(0, font)

        tree_item.setFlags(
            Qt.ItemFlag.ItemIsEnabled |
            Qt.ItemFlag.ItemIsSelectable |
            Qt.ItemFlag.ItemIsDragEnabled
        )

        return tree_item

    def _on_selection_changed(self) -> None:
        """Handle selection change in tree."""
        selected = self.selectedItems()
        if not selected:
            return

        item = selected[0]
        item_type = item.data(0, Qt.ItemDataRole.UserRole)

        if item_type == self.ITEM_TYPE_SECTION:
            section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
            self.section_selected.emit(section_id)
        elif item_type == self.ITEM_TYPE_SLIDE:
            section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
            slide_id = item.data(0, Qt.ItemDataRole.UserRole + 2)
            self.slide_selected.emit(section_id, slide_id)

    def _on_item_double_clicked(self, item: QTreeWidgetItem, column: int) -> None:
        """Handle double-click on item."""
        self.item_double_clicked.emit()

    def _on_context_menu(self, position) -> None:
        """Show context menu."""
        item = self.itemAt(position)
        if not item:
            return

        menu = QMenu()

        item_type = item.data(0, Qt.ItemDataRole.UserRole)

        if item_type == self.ITEM_TYPE_SECTION:
            menu.addAction(tr("context.duplicate"), self._duplicate_selected_section)
            menu.addAction(tr("menu.edit.delete"), self._delete_selected_section)
            menu.addSeparator()
            menu.addAction(tr("menu.edit.move_up"), self._move_section_up)
            menu.addAction(tr("menu.edit.move_down"), self._move_section_down)
        elif item_type == self.ITEM_TYPE_SLIDE:
            menu.addAction(tr("button.edit"), self._edit_selected_slide)

        menu.exec(self.viewport().mapToGlobal(position))

    def dropEvent(self, event: QDropEvent) -> None:
        """Handle drop event with constraints."""
        dragged_item = self.currentItem()
        if not dragged_item or not self._liturgy:
            event.ignore()
            return

        dragged_type = dragged_item.data(0, Qt.ItemDataRole.UserRole)

        # Get drop position
        pos = event.position().toPoint()
        target_item = self.itemAt(pos)
        drop_indicator = self.dropIndicatorPosition()

        if dragged_type == self.ITEM_TYPE_SECTION:
            self._do_section_drop(dragged_item, target_item, drop_indicator)
        elif dragged_type == self.ITEM_TYPE_SLIDE:
            self._do_slide_drop(dragged_item, target_item, drop_indicator)

        # Always ignore the default Qt handling - we do it manually
        event.ignore()

    def _do_section_drop(self, dragged_item: QTreeWidgetItem, target_item: QTreeWidgetItem, drop_indicator) -> None:
        """Handle section reordering."""
        if not self._liturgy.sections:
            return

        # Get dragged section index
        dragged_idx = self.indexOfTopLevelItem(dragged_item)
        if dragged_idx < 0:
            return

        # Determine target index
        if target_item is None:
            # Dropped at the very end
            target_idx = len(self._liturgy.sections) - 1
        else:
            target_type = target_item.data(0, Qt.ItemDataRole.UserRole)

            if target_type == self.ITEM_TYPE_SLIDE:
                # Dropped on a slide - use its parent section
                parent_section = target_item.parent()
                if not parent_section:
                    return
                target_idx = self.indexOfTopLevelItem(parent_section)
                # When dropping on/between slides, position relative to that section
                if drop_indicator == QAbstractItemView.DropIndicatorPosition.BelowItem:
                    target_idx += 1
            else:
                target_idx = self.indexOfTopLevelItem(target_item)
                if target_idx < 0:
                    return

                # Adjust for drop indicator position
                if drop_indicator == QAbstractItemView.DropIndicatorPosition.BelowItem:
                    target_idx += 1
                elif drop_indicator == QAbstractItemView.DropIndicatorPosition.OnItem:
                    # Dropping ON a section - treat as dropping below it
                    target_idx += 1

        # Adjust target if dragging from above
        if dragged_idx < target_idx:
            target_idx -= 1

        # Clamp to valid range
        target_idx = max(0, min(target_idx, len(self._liturgy.sections) - 1))

        if dragged_idx != target_idx:
            self._liturgy.move_section(dragged_idx, target_idx)
            self._update_display()
            # Select the moved item
            self.setCurrentItem(self.topLevelItem(target_idx))
            self.order_changed.emit()

    def _do_slide_drop(self, dragged_item: QTreeWidgetItem, target_item: QTreeWidgetItem, drop_indicator) -> None:
        """Handle slide reordering - within section or between sections."""
        if not self._liturgy.sections:
            return

        dragged_parent = dragged_item.parent()
        if not dragged_parent:
            return

        # Get source section and slide indices
        source_section_idx = self.indexOfTopLevelItem(dragged_parent)
        if source_section_idx < 0 or source_section_idx >= len(self._liturgy.sections):
            return

        source_section = self._liturgy.sections[source_section_idx]
        dragged_slide_idx = dragged_parent.indexOfChild(dragged_item)
        if dragged_slide_idx < 0 or dragged_slide_idx >= len(source_section.slides):
            return

        # Determine target section and position
        if target_item is None:
            return

        target_type = target_item.data(0, Qt.ItemDataRole.UserRole)

        if target_type == self.ITEM_TYPE_SECTION:
            # Dropped on a section header - add to end of that section
            target_section_idx = self.indexOfTopLevelItem(target_item)
            if target_section_idx < 0:
                return
            target_section = self._liturgy.sections[target_section_idx]
            target_slide_idx = len(target_section.slides)  # Add at end
        else:
            # Dropped on another slide
            target_parent = target_item.parent()
            if not target_parent:
                return

            target_section_idx = self.indexOfTopLevelItem(target_parent)
            if target_section_idx < 0:
                return
            target_section = self._liturgy.sections[target_section_idx]

            target_slide_idx = target_parent.indexOfChild(target_item)
            if target_slide_idx < 0:
                return

            # Adjust for drop indicator
            if drop_indicator == QAbstractItemView.DropIndicatorPosition.BelowItem:
                target_slide_idx += 1
            elif drop_indicator == QAbstractItemView.DropIndicatorPosition.OnItem:
                # Dropping ON a slide - insert after it
                target_slide_idx += 1

        # Check if moving within same section or between sections
        if source_section_idx == target_section_idx:
            # Moving within same section
            # Adjust target if dragging from above
            if dragged_slide_idx < target_slide_idx:
                target_slide_idx -= 1

            # Clamp to valid range
            target_slide_idx = max(0, min(target_slide_idx, len(source_section.slides) - 1))

            if dragged_slide_idx != target_slide_idx:
                self._liturgy.move_slide_within_section(source_section_idx, dragged_slide_idx, target_slide_idx)
                self._update_display()
                # Select the moved item
                new_section_item = self.topLevelItem(source_section_idx)
                if new_section_item and target_slide_idx < new_section_item.childCount():
                    self.setCurrentItem(new_section_item.child(target_slide_idx))
                self.order_changed.emit()
        else:
            # Moving between sections
            self._liturgy.move_slide_to_section(
                source_section_idx, dragged_slide_idx,
                target_section_idx, target_slide_idx
            )
            self._update_display()
            # Select the moved item in new location
            new_section_item = self.topLevelItem(target_section_idx)
            # Clamp target_slide_idx for selection
            actual_idx = min(target_slide_idx, new_section_item.childCount() - 1) if new_section_item else 0
            if new_section_item and actual_idx >= 0 and actual_idx < new_section_item.childCount():
                self.setCurrentItem(new_section_item.child(actual_idx))
            self.order_changed.emit()

    def _rebuild_from_tree(self) -> None:
        """Rebuild the liturgy from the current tree state."""
        if not self._liturgy:
            return

        if self._liturgy.sections:
            # V2 mode: rebuild sections
            new_sections = []

            for i in range(self.topLevelItemCount()):
                tree_section = self.topLevelItem(i)
                section_id = tree_section.data(0, Qt.ItemDataRole.UserRole + 1)

                # Find the original section
                section = self._liturgy.get_section_by_id(section_id)
                if section:
                    # Rebuild slides order from tree
                    new_slides = []
                    for j in range(tree_section.childCount()):
                        tree_slide = tree_section.child(j)
                        slide_id = tree_slide.data(0, Qt.ItemDataRole.UserRole + 2)

                        # Find the slide in the section
                        for slide in section.slides:
                            if slide.id == slide_id:
                                new_slides.append(slide)
                                break

                    section.slides = new_slides
                    new_sections.append(section)

            self._liturgy.sections = new_sections
        else:
            # V1 mode: rebuild items
            new_items = []

            for i in range(self.topLevelItemCount()):
                tree_item = self.topLevelItem(i)
                item = tree_item.data(0, Qt.ItemDataRole.UserRole + 3)
                if item:
                    new_items.append(item)

            self._liturgy._items = new_items

        # Refresh display to update numbering
        self._update_display()

    def _delete_selected_section(self) -> None:
        """Delete the selected section."""
        selected = self.selectedItems()
        if not selected:
            return

        item = selected[0]
        item_type = item.data(0, Qt.ItemDataRole.UserRole)

        if item_type != self.ITEM_TYPE_SECTION:
            return

        section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)

        if self._liturgy.sections:
            # V2 mode
            for i, section in enumerate(self._liturgy.sections):
                if section.id == section_id:
                    self._liturgy.remove_section(i)
                    break
        else:
            # V1 mode
            try:
                index = int(section_id)
                self._liturgy.remove_item(index)
            except ValueError:
                pass

        self._update_display()
        self.order_changed.emit()

    def _duplicate_selected_section(self) -> None:
        """Duplicate the selected section."""
        selected = self.selectedItems()
        if not selected:
            return

        item = selected[0]
        item_type = item.data(0, Qt.ItemDataRole.UserRole)

        if item_type != self.ITEM_TYPE_SECTION:
            return

        section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)

        if self._liturgy.sections:
            # V2 mode - find and duplicate section
            for i, section in enumerate(self._liturgy.sections):
                if section.id == section_id:
                    # Create a copy and insert after original
                    copy = section.copy()
                    copy.name = f"{section.name} (kopie)"
                    self._liturgy.insert_section(i + 1, copy)
                    break

            self._update_display()
            self.order_changed.emit()

    def _move_section_up(self) -> None:
        """Move selected section up."""
        selected = self.selectedItems()
        if not selected:
            return

        item = selected[0]
        index = self.indexOfTopLevelItem(item)
        if index > 0:
            if self._liturgy.sections:
                self._liturgy.move_section(index, index - 1)
            else:
                self._liturgy.move_item(index, index - 1)
            self._update_display()
            self.setCurrentItem(self.topLevelItem(index - 1))
            self.order_changed.emit()

    def _move_section_down(self) -> None:
        """Move selected section down."""
        selected = self.selectedItems()
        if not selected:
            return

        item = selected[0]
        index = self.indexOfTopLevelItem(item)
        max_index = self.topLevelItemCount() - 1
        if index < max_index:
            if self._liturgy.sections:
                self._liturgy.move_section(index, index + 1)
            else:
                self._liturgy.move_item(index, index + 1)
            self._update_display()
            self.setCurrentItem(self.topLevelItem(index + 1))
            self.order_changed.emit()

    def _edit_selected_slide(self) -> None:
        """Trigger edit for selected slide."""
        self.item_double_clicked.emit()

    def get_selected_section_index(self) -> int:
        """Get the index of the currently selected section, or -1 if none."""
        selected = self.selectedItems()
        if not selected:
            return -1

        item = selected[0]
        item_type = item.data(0, Qt.ItemDataRole.UserRole)

        if item_type == self.ITEM_TYPE_SECTION:
            return self.indexOfTopLevelItem(item)
        elif item_type == self.ITEM_TYPE_SLIDE:
            # Return parent section index
            parent = item.parent()
            if parent:
                return self.indexOfTopLevelItem(parent)

        return -1

    def get_selected_slide_info(self) -> Optional[Tuple[str, str]]:
        """Get the selected slide info as (section_id, slide_id) or None."""
        selected = self.selectedItems()
        if not selected:
            return None

        item = selected[0]
        item_type = item.data(0, Qt.ItemDataRole.UserRole)

        if item_type == self.ITEM_TYPE_SLIDE:
            section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
            slide_id = item.data(0, Qt.ItemDataRole.UserRole + 2)
            return (section_id, slide_id)

        return None

    def get_selected_index(self) -> int:
        """Get the index of the currently selected item (v1 compatibility)."""
        return self.get_selected_section_index()

    def select_section(self, index: int) -> None:
        """Select the section at the given index."""
        if 0 <= index < self.topLevelItemCount():
            self.setCurrentItem(self.topLevelItem(index))

    def select_index(self, index: int) -> None:
        """Select the item at the given index (v1 compatibility)."""
        self.select_section(index)

    # V1 compatibility methods
    def get_elements(self) -> List[LiturgyItem]:
        """Get the current list of items in order. (Backwards compatibility)"""
        if self._liturgy:
            return self._liturgy.items
        return []

    def get_items(self) -> List[LiturgyItem]:
        """Get the current list of items in order."""
        return self.get_elements()

    def refresh(self) -> None:
        """Refresh display (e.g., after language change)."""
        self._field_cache.clear()  # Clear cache to re-check fields
        self._update_display()

    def get_slides_with_unfilled_fields(self) -> List[Tuple[LiturgySection, LiturgySlide, List[str]]]:
        """Get all slides that have unfilled fields.
        Returns list of (section, slide, unfilled_field_names) tuples.
        """
        result = []
        if not self._liturgy:
            return result

        for section in self._liturgy.sections:
            for slide in section.slides:
                unfilled = self._get_unfilled_fields(slide)
                if unfilled:
                    result.append((section, slide, unfilled))

        return result
