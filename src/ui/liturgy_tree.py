"""Liturgy tree widget with hierarchical section/slide support and constrained drag-drop."""

from typing import List, Optional, Tuple

from PyQt6.QtWidgets import (
    QTreeWidget,
    QTreeWidgetItem,
    QAbstractItemView,
    QMenu,
    QLineEdit,
    QStyledItemDelegate,
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QKeyEvent, QKeySequence, QDropEvent, QDragEnterEvent, QDragMoveEvent

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
from ..logging_config import get_logger

logger = get_logger("liturgy_tree")


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
        # Clipboard for copy/paste
        self._clipboard_section: Optional[LiturgySection] = None
        self._clipboard_slide: Optional[Tuple[LiturgySlide, str]] = None  # (slide, source_section_id)
        # Inline editor
        self._edit_widget: Optional[QLineEdit] = None

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
            except Exception as e:
                logger.debug(f"Could not extract fields from {slide.source_path}[{slide.slide_index}]: {e}")
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

        # Style drop indicator for better visibility
        self.setStyleSheet("""
            QTreeWidget::item:selected {
                background-color: #0078d4;
                color: white;
            }
            QTreeWidget {
                show-decoration-selected: 1;
            }
        """)

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
            # Determine if this is a song section
            is_song_section = section.is_song
            if not is_song_section and section.slides:
                is_song_section = any(
                    slide.youtube_links or slide.pdf_path
                    for slide in section.slides
                )

            section_item = self._create_section_item(section, section_idx, is_song_section)
            self.addTopLevelItem(section_item)

            # Add slides as children
            for slide_idx, slide in enumerate(section.slides):
                slide_item = self._create_slide_item(
                    slide, section.id, slide_idx, len(section.slides), is_song_section
                )
                section_item.addChild(slide_item)

            # Expand by default
            section_item.setExpanded(True)

    def _display_items_as_sections(self) -> None:
        """Display v1 items as pseudo-sections (for backwards compatibility)."""
        for item_idx, item in enumerate(self._liturgy._items):
            # Create a pseudo-section for each item
            section_item = self._create_item_as_section(item, item_idx)
            self.addTopLevelItem(section_item)

    def _create_section_item(
        self, section: LiturgySection, index: int, is_song_section: bool = None
    ) -> QTreeWidgetItem:
        """Create a tree item for a section."""
        item = QTreeWidgetItem()

        # Determine if this is a song section (if not passed)
        if is_song_section is None:
            is_song_section = section.is_song
            if not is_song_section and section.slides:
                is_song_section = any(
                    slide.youtube_links or slide.pdf_path
                    for slide in section.slides
                )

        # Track section-level warnings
        warnings = []

        # Check for missing PowerPoint in non-stub slides
        slides_missing_pptx = [
            s for s in section.slides
            if not s.is_stub and (not s.source_path or not os.path.exists(s.source_path))
        ]
        has_pptx_error = len(slides_missing_pptx) > 0

        # Format display text
        icon = "🎵" if is_song_section else "📁"
        indicators = []

        # PowerPoint status
        if section.has_pptx and not has_pptx_error:
            indicators.append("📊")  # PPT icon - all present
        elif has_pptx_error:
            indicators.append("❌")  # Missing PowerPoint error
            warnings.append(f"{tr('warning.section_missing_pptx')}: {len(slides_missing_pptx)}")

        # YouTube status (for songs)
        if is_song_section:
            has_youtube = section.has_youtube or any(slide.youtube_links for slide in section.slides)
            slides_missing_youtube = [
                s for s in section.slides if not s.is_stub and not s.youtube_links
            ]
            if has_youtube:
                indicators.append("📺")  # YouTube icon - present
            elif slides_missing_youtube:
                indicators.append("🔇")  # No YouTube links
                warnings.append(f"{tr('warning.section_missing_youtube')}: {len(slides_missing_youtube)}")

            # PDF status (for songs)
            has_pdf = section.has_pdf or any(slide.pdf_path for slide in section.slides)
            slides_missing_pdf = [
                s for s in section.slides if not s.is_stub and not s.pdf_path
            ]
            if has_pdf:
                indicators.append("📕")  # PDF icon - present
            elif slides_missing_pdf:
                indicators.append("📃")  # No music PDF
                warnings.append(f"{tr('warning.section_missing_pdf')}: {len(slides_missing_pdf)}")

        indicator_text = " ".join(indicators)
        clean_name = self._clean_title(section.name)
        display_text = f"{index + 1}. {icon} {clean_name}"
        if indicator_text:
            display_text += f"  {indicator_text}"

        # Add error indicator if there are critical warnings
        if has_pptx_error:
            display_text += " ⚠"

        item.setText(0, display_text)
        item.setData(0, Qt.ItemDataRole.UserRole, self.ITEM_TYPE_SECTION)
        item.setData(0, Qt.ItemDataRole.UserRole + 1, section.id)

        # Set tooltip with warnings
        if warnings:
            item.setToolTip(0, "\n".join(warnings))

        # Style: bold for sections, red foreground if error
        font = item.font(0)
        font.setBold(True)
        item.setFont(0, font)

        if has_pptx_error:
            item.setForeground(0, Qt.GlobalColor.red)

        # Allow children (slides) but not dropping of other sections
        item.setFlags(
            Qt.ItemFlag.ItemIsEnabled |
            Qt.ItemFlag.ItemIsSelectable |
            Qt.ItemFlag.ItemIsDragEnabled |
            Qt.ItemFlag.ItemIsDropEnabled
        )

        return item

    def _create_slide_item(
        self, slide: LiturgySlide, section_id: str, index: int, total: int,
        is_song_section: bool = False
    ) -> QTreeWidgetItem:
        """Create a tree item for a slide."""
        item = QTreeWidgetItem()

        # Format display text with tree-like prefix
        prefix = "└─" if index == total - 1 else "├─"
        clean_title = self._clean_title(slide.title) or f'Slide {index + 1}'

        # Track warnings for tooltip
        warnings = []

        # Check PowerPoint status
        has_pptx = slide.source_path and os.path.exists(slide.source_path)
        pptx_missing = not slide.is_stub and not has_pptx

        # Check if this is a song slide (section is song or slide has song metadata)
        is_song_slide = is_song_section or slide.youtube_links or slide.pdf_path

        # Add indicators for slide-level properties
        indicators = []

        # PowerPoint indicator
        if has_pptx:
            indicators.append("📊")  # PPT icon - present
        elif pptx_missing:
            indicators.append("❌")  # Missing PowerPoint - critical error
            warnings.append(tr("warning.missing_pptx") if slide.source_path else tr("warning.no_pptx"))

        # YouTube indicator (for songs)
        if slide.youtube_links:
            indicators.append("📺")  # YouTube icon - present
        elif is_song_slide and not slide.is_stub:
            indicators.append("🔇")  # No YouTube link
            warnings.append(tr("warning.missing_youtube"))

        # PDF indicator (for songs)
        if slide.pdf_path:
            indicators.append("📕")  # PDF icon - present
        elif is_song_slide and not slide.is_stub:
            indicators.append("📃")  # No music PDF
            warnings.append(tr("warning.missing_pdf"))

        display_text = f"{prefix} {clean_title}"
        if indicators:
            display_text += f"  {' '.join(indicators)}"
        if slide.is_stub:
            display_text += " (stub)"

        # Check for field status
        all_fields, unfilled = self._get_field_status(slide)
        if all_fields:  # Has text pattern fields
            if unfilled:
                display_text += " ⚠"  # Warning: unfilled fields
                warnings.append(f"{tr('warning.unfilled_fields')}: {', '.join(unfilled)}")
            else:
                display_text += " ✓"  # All fields filled

        # Build comprehensive tooltip
        if warnings:
            item.setToolTip(0, "\n".join(warnings))
        elif all_fields and not unfilled:
            item.setToolTip(0, f"{tr('tooltip.fields_filled')}: {', '.join(all_fields)}")

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

        # Red text for missing PowerPoint (critical error)
        if pptx_missing:
            item.setForeground(0, Qt.GlobalColor.red)

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

    def keyPressEvent(self, event: QKeyEvent) -> None:
        """Handle key press events for F2 (rename), Ctrl+C (copy), Ctrl+V (paste)."""
        if event.key() == Qt.Key.Key_F2:
            self._start_inline_edit()
            return
        elif event.matches(QKeySequence.StandardKey.Copy):
            self._copy_selected()
            return
        elif event.matches(QKeySequence.StandardKey.Paste):
            self._paste()
            return

        super().keyPressEvent(event)

    def _start_inline_edit(self) -> None:
        """Start inline editing of the selected item's title."""
        selected = self.selectedItems()
        if not selected:
            return

        item = selected[0]
        item_type = item.data(0, Qt.ItemDataRole.UserRole)

        # Get the current title
        if item_type == self.ITEM_TYPE_SECTION:
            section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
            section = self._liturgy.get_section_by_id(section_id) if self._liturgy else None
            if not section:
                return
            current_title = section.name
        elif item_type == self.ITEM_TYPE_SLIDE:
            section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
            slide_id = item.data(0, Qt.ItemDataRole.UserRole + 2)
            slide = self._get_slide_by_ids(section_id, slide_id)
            if not slide:
                return
            current_title = slide.title
        else:
            return

        # Create inline editor
        rect = self.visualItemRect(item)
        self._edit_widget = QLineEdit(self)
        self._edit_widget.setText(current_title)
        self._edit_widget.setGeometry(rect)
        self._edit_widget.selectAll()
        self._edit_widget.setFocus()
        self._edit_widget.show()

        # Store item info for when editing finishes
        self._edit_widget.setProperty("item_type", item_type)
        self._edit_widget.setProperty("section_id", section_id if item_type == self.ITEM_TYPE_SECTION else section_id)
        self._edit_widget.setProperty("slide_id", slide_id if item_type == self.ITEM_TYPE_SLIDE else None)

        # Connect signals
        self._edit_widget.editingFinished.connect(self._finish_inline_edit)
        self._edit_widget.installEventFilter(self)

    def eventFilter(self, obj, event) -> bool:
        """Filter events for the inline editor."""
        if obj == self._edit_widget and event.type() == event.Type.KeyPress:
            if event.key() == Qt.Key.Key_Escape:
                self._cancel_inline_edit()
                return True
        return super().eventFilter(obj, event)

    def _finish_inline_edit(self) -> None:
        """Finish inline editing and apply the new title."""
        if not self._edit_widget:
            return

        new_title = self._edit_widget.text().strip()
        item_type = self._edit_widget.property("item_type")
        section_id = self._edit_widget.property("section_id")
        slide_id = self._edit_widget.property("slide_id")

        self._edit_widget.deleteLater()
        self._edit_widget = None

        if not new_title:
            self._update_display()
            return

        if item_type == self.ITEM_TYPE_SECTION:
            section = self._liturgy.get_section_by_id(section_id) if self._liturgy else None
            if section and section.name != new_title:
                section.name = new_title
                self._update_display()
                self.order_changed.emit()
        elif item_type == self.ITEM_TYPE_SLIDE:
            slide = self._get_slide_by_ids(section_id, slide_id)
            if slide and slide.title != new_title:
                slide.title = new_title
                self._update_display()
                self.order_changed.emit()

    def _cancel_inline_edit(self) -> None:
        """Cancel inline editing."""
        if self._edit_widget:
            self._edit_widget.deleteLater()
            self._edit_widget = None

    def _copy_selected(self) -> None:
        """Copy the selected section or slide to clipboard."""
        selected = self.selectedItems()
        if not selected or not self._liturgy:
            return

        item = selected[0]
        item_type = item.data(0, Qt.ItemDataRole.UserRole)

        if item_type == self.ITEM_TYPE_SECTION:
            section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
            section = self._liturgy.get_section_by_id(section_id)
            if section:
                self._clipboard_section = section.copy()
                self._clipboard_slide = None
                logger.debug(f"Copied section: {section.name}")
        elif item_type == self.ITEM_TYPE_SLIDE:
            section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
            slide_id = item.data(0, Qt.ItemDataRole.UserRole + 2)
            slide = self._get_slide_by_ids(section_id, slide_id)
            if slide:
                self._clipboard_slide = (slide.copy(), section_id)
                self._clipboard_section = None
                logger.debug(f"Copied slide: {slide.title}")

    def _paste(self) -> None:
        """Paste the copied section or slide."""
        if not self._liturgy:
            return

        selected = self.selectedItems()
        target_section_idx = -1
        target_slide_idx = -1

        if selected:
            item = selected[0]
            item_type = item.data(0, Qt.ItemDataRole.UserRole)

            if item_type == self.ITEM_TYPE_SECTION:
                target_section_idx = self.indexOfTopLevelItem(item)
            elif item_type == self.ITEM_TYPE_SLIDE:
                parent = item.parent()
                if parent:
                    target_section_idx = self.indexOfTopLevelItem(parent)
                    target_slide_idx = parent.indexOfChild(item)

        if self._clipboard_section:
            # Paste section after the selected section (or at end)
            new_section = self._clipboard_section.copy()
            new_section.name = f"{new_section.name} (kopie)"

            if target_section_idx >= 0:
                self._liturgy.insert_section(target_section_idx + 1, new_section)
            else:
                self._liturgy.add_section(new_section)

            self._update_display()
            self.order_changed.emit()
            logger.debug(f"Pasted section: {new_section.name}")

        elif self._clipboard_slide:
            slide, _ = self._clipboard_slide
            new_slide = slide.copy()

            if target_section_idx >= 0 and target_section_idx < len(self._liturgy.sections):
                target_section = self._liturgy.sections[target_section_idx]
                if target_slide_idx >= 0:
                    target_section.slides.insert(target_slide_idx + 1, new_slide)
                else:
                    target_section.slides.append(new_slide)

                self._update_display()
                self.order_changed.emit()
                logger.debug(f"Pasted slide: {new_slide.title}")

    def _on_context_menu(self, position) -> None:
        """Show context menu."""
        item = self.itemAt(position)
        if not item:
            return

        menu = QMenu()

        item_type = item.data(0, Qt.ItemDataRole.UserRole)

        if item_type == self.ITEM_TYPE_SECTION:
            section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
            section = self._liturgy.get_section_by_id(section_id) if self._liturgy else None

            # Add "Open PowerPoint" if section has slides with a source path
            if section and section.has_pptx:
                # Get the pptx path from the first slide that has one
                pptx_path = None
                for slide in section.slides:
                    if slide.source_path and os.path.exists(slide.source_path):
                        pptx_path = slide.source_path
                        break
                if pptx_path:
                    menu.addAction(tr("context.open_pptx"), lambda p=pptx_path: self._open_pptx_file(p))
                    menu.addSeparator()

            menu.addAction(tr("context.duplicate"), self._duplicate_selected_section)
            menu.addAction(tr("menu.edit.delete"), self._delete_selected_section)
            menu.addSeparator()
            menu.addAction(tr("menu.edit.move_up"), self._move_section_up)
            menu.addAction(tr("menu.edit.move_down"), self._move_section_down)
        elif item_type == self.ITEM_TYPE_SLIDE:
            section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
            slide_id = item.data(0, Qt.ItemDataRole.UserRole + 2)
            slide = self._get_slide_by_ids(section_id, slide_id)

            # Add "Open PowerPoint" if slide has a source path
            if slide and slide.source_path and os.path.exists(slide.source_path):
                menu.addAction(tr("context.open_pptx"), lambda: self._open_pptx_file(slide.source_path))
                menu.addSeparator()

            menu.addAction(tr("button.edit"), self._edit_selected_slide)
            menu.addSeparator()
            menu.addAction(tr("menu.edit.move_up"), self._move_slide_up)
            menu.addAction(tr("menu.edit.move_down"), self._move_slide_down)
            menu.addAction(tr("menu.edit.delete"), self._delete_selected_slide)

        menu.exec(self.viewport().mapToGlobal(position))

    def _get_slide_by_ids(self, section_id: str, slide_id: str) -> Optional[LiturgySlide]:
        """Get a slide by section and slide IDs."""
        if not self._liturgy:
            return None
        section = self._liturgy.get_section_by_id(section_id)
        if not section:
            return None
        for slide in section.slides:
            if slide.id == slide_id:
                return slide
        return None

    def _open_pptx_file(self, path: str) -> None:
        """Open a PowerPoint file with the default application."""
        if os.path.exists(path):
            os.startfile(path)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        """Accept drag enter events for internal moves."""
        # Always call super to handle visual feedback
        super().dragEnterEvent(event)
        if event.source() == self:
            event.acceptProposedAction()

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:
        """Accept drag move events for internal moves."""
        # Always call super to draw drop indicator
        super().dragMoveEvent(event)
        if event.source() == self:
            event.acceptProposedAction()

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
                # AboveItem/OnItem = take target's position
                # BelowItem = insert after target
                if drop_indicator == QAbstractItemView.DropIndicatorPosition.BelowItem:
                    target_idx += 1
                # OnItem: no adjustment, section takes target's position

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
        """Handle slide reordering - within section or between sections.

        Drop indicator logic:
        - AboveItem: Insert BEFORE the target (at target's current index)
        - BelowItem: Insert AFTER the target (at target's index + 1)
        - OnItem: Insert AFTER the target (at target's index + 1)
        """
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

            # Adjust for drop indicator position
            # AboveItem = take target's position (most intuitive for "drop on item")
            # OnItem = same as AboveItem (take target's position)
            # BelowItem = insert after target
            if drop_indicator == QAbstractItemView.DropIndicatorPosition.BelowItem:
                target_slide_idx += 1
            # AboveItem and OnItem: no adjustment, item takes target's position

        # Check if moving within same section or between sections
        if source_section_idx == target_section_idx:
            # Moving within same section
            # target_slide_idx is the insertion index (0 to N)

            # If dragging from above the insertion point, the removal will shift indices down
            # So we need to adjust the target index
            if dragged_slide_idx < target_slide_idx:
                target_slide_idx -= 1

            # Now target_slide_idx is the final position after the move
            # Clamp to valid range (0 to N-1)
            max_idx = len(source_section.slides) - 1
            target_slide_idx = max(0, min(target_slide_idx, max_idx))

            if dragged_slide_idx != target_slide_idx:
                self._liturgy.move_slide_within_section(source_section_idx, dragged_slide_idx, target_slide_idx)
                self._update_display()
                # Select the moved item
                new_section_item = self.topLevelItem(source_section_idx)
                if new_section_item and target_slide_idx < new_section_item.childCount():
                    self.setCurrentItem(new_section_item.child(target_slide_idx))
                self.order_changed.emit()
        else:
            # Moving between sections - target_slide_idx is insertion index
            # Clamp to valid range for the target section (0 to N, where N is current length)
            max_idx = len(target_section.slides)
            target_slide_idx = max(0, min(target_slide_idx, max_idx))

            self._liturgy.move_slide_to_section(
                source_section_idx, dragged_slide_idx,
                target_section_idx, target_slide_idx
            )
            self._update_display()
            # Select the moved item in new location
            new_section_item = self.topLevelItem(target_section_idx)
            # After insertion, the slide is at target_slide_idx
            actual_idx = min(target_slide_idx, new_section_item.childCount() - 1) if new_section_item else 0
            if new_section_item and actual_idx >= 0 and actual_idx < new_section_item.childCount():
                self.setCurrentItem(new_section_item.child(actual_idx))
            self.order_changed.emit()
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

    def _move_slide_up(self) -> None:
        """Move selected slide up within its section."""
        selected = self.selectedItems()
        if not selected or not self._liturgy:
            return

        item = selected[0]
        item_type = item.data(0, Qt.ItemDataRole.UserRole)
        if item_type != self.ITEM_TYPE_SLIDE:
            return

        parent = item.parent()
        if not parent:
            return

        section_idx = self.indexOfTopLevelItem(parent)
        slide_idx = parent.indexOfChild(item)

        if section_idx < 0 or section_idx >= len(self._liturgy.sections):
            return

        if slide_idx > 0:
            self._liturgy.move_slide_within_section(section_idx, slide_idx, slide_idx - 1)
            self._update_display()
            # Reselect the moved slide
            new_parent = self.topLevelItem(section_idx)
            if new_parent and slide_idx - 1 < new_parent.childCount():
                self.setCurrentItem(new_parent.child(slide_idx - 1))
            self.order_changed.emit()

    def _move_slide_down(self) -> None:
        """Move selected slide down within its section."""
        selected = self.selectedItems()
        if not selected or not self._liturgy:
            return

        item = selected[0]
        item_type = item.data(0, Qt.ItemDataRole.UserRole)
        if item_type != self.ITEM_TYPE_SLIDE:
            return

        parent = item.parent()
        if not parent:
            return

        section_idx = self.indexOfTopLevelItem(parent)
        slide_idx = parent.indexOfChild(item)

        if section_idx < 0 or section_idx >= len(self._liturgy.sections):
            return

        section = self._liturgy.sections[section_idx]
        if slide_idx < len(section.slides) - 1:
            self._liturgy.move_slide_within_section(section_idx, slide_idx, slide_idx + 1)
            self._update_display()
            # Reselect the moved slide
            new_parent = self.topLevelItem(section_idx)
            if new_parent and slide_idx + 1 < new_parent.childCount():
                self.setCurrentItem(new_parent.child(slide_idx + 1))
            self.order_changed.emit()

    def _delete_selected_slide(self) -> None:
        """Delete the selected slide."""
        selected = self.selectedItems()
        if not selected or not self._liturgy:
            return

        item = selected[0]
        item_type = item.data(0, Qt.ItemDataRole.UserRole)
        if item_type != self.ITEM_TYPE_SLIDE:
            return

        section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
        slide_id = item.data(0, Qt.ItemDataRole.UserRole + 2)

        section = self._liturgy.get_section_by_id(section_id)
        if not section:
            return

        # Find and remove the slide
        for i, slide in enumerate(section.slides):
            if slide.id == slide_id:
                section.slides.pop(i)
                break

        self._update_display()
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
