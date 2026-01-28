"""Dialog for selecting sections/slides from theme PPTX files."""

import os
from typing import List, Optional, Tuple

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QTreeWidget,
    QTreeWidgetItem,
    QLineEdit,
    QPushButton,
    QLabel,
    QDialogButtonBox,
    QGroupBox,
    QFileDialog,
    QComboBox,
    QCheckBox,
    QMessageBox,
)
from PyQt6.QtCore import Qt

from ..models import (
    LiturgySection,
    LiturgySlide,
    SectionType,
    Settings,
)
from ..services import ThemeService
from ..i18n import tr


class ThemeSectionPicker(QDialog):
    """Dialog for selecting sections from theme PPTX files."""

    ITEM_TYPE_THEME = 1
    ITEM_TYPE_SECTION = 2
    ITEM_TYPE_SLIDE = 3

    def __init__(
        self,
        settings: Settings,
        base_path: str,
        existing_sections: List[LiturgySection] = None,
        parent=None
    ):
        super().__init__(parent)
        self.settings = settings
        self.base_path = base_path
        self.theme_service = ThemeService(settings, base_path)
        self.existing_sections = existing_sections or []

        self._selected_sections: List[LiturgySection] = []
        self._add_to_existing: Optional[str] = None  # Section ID to add slides to

        self._setup_ui()
        self._populate_themes()
        self._connect_signals()

    def _setup_ui(self) -> None:
        """Setup the dialog UI."""
        self.setWindowTitle(tr("dialog.theme.title"))
        self.setMinimumSize(600, 600)
        self.resize(700, 700)

        layout = QVBoxLayout(self)

        # Instructions label
        self.instruction_label = QLabel(tr("dialog.theme.select_sections"))
        layout.addWidget(self.instruction_label)

        # Theme file selector
        theme_layout = QHBoxLayout()
        theme_label = QLabel(tr("dialog.theme.theme_file"))
        self.theme_combo = QComboBox()
        self.theme_combo.setMinimumWidth(300)
        self.browse_theme_btn = QPushButton(tr("button.browse"))
        theme_layout.addWidget(theme_label)
        theme_layout.addWidget(self.theme_combo, 1)
        theme_layout.addWidget(self.browse_theme_btn)
        layout.addLayout(theme_layout)

        # Search field
        search_layout = QHBoxLayout()
        search_label = QLabel(tr("dialog.theme.search"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(tr("dialog.theme.search"))
        self.search_input.setClearButtonEnabled(True)
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)

        # Tree widget for sections and slides
        self.tree_widget = QTreeWidget()
        self.tree_widget.setHeaderHidden(True)
        self.tree_widget.setSelectionMode(QTreeWidget.SelectionMode.ExtendedSelection)
        self.tree_widget.setAlternatingRowColors(True)
        layout.addWidget(self.tree_widget)

        # Add to options
        add_group = QGroupBox(tr("dialog.theme.add_options"))
        add_layout = QVBoxLayout(add_group)

        self.new_section_radio = QCheckBox(tr("dialog.theme.add_new_section"))
        self.new_section_radio.setChecked(True)
        add_layout.addWidget(self.new_section_radio)

        existing_layout = QHBoxLayout()
        self.existing_section_radio = QCheckBox(tr("dialog.theme.add_to_existing"))
        self.existing_combo = QComboBox()
        self.existing_combo.setEnabled(False)

        # Populate existing sections
        for section in self.existing_sections:
            self.existing_combo.addItem(section.name, section.id)

        existing_layout.addWidget(self.existing_section_radio)
        existing_layout.addWidget(self.existing_combo, 1)
        add_layout.addLayout(existing_layout)

        layout.addWidget(add_group)

        # Selection info
        self.info_label = QLabel()
        layout.addWidget(self.info_label)

        # Button box
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setText(tr("button.add"))
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText(tr("button.cancel"))
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setEnabled(False)
        layout.addWidget(self.button_box)

    def _populate_themes(self) -> None:
        """Populate the theme combo box."""
        theme_files = self.theme_service.get_theme_files()

        for theme_path in theme_files:
            name = os.path.basename(theme_path)
            self.theme_combo.addItem(name, theme_path)

        if theme_files:
            self._load_theme(theme_files[0])

    def _connect_signals(self) -> None:
        """Connect widget signals."""
        self.theme_combo.currentIndexChanged.connect(self._on_theme_changed)
        self.browse_theme_btn.clicked.connect(self._on_browse_theme)
        self.search_input.textChanged.connect(self._on_search_changed)
        self.tree_widget.itemSelectionChanged.connect(self._on_selection_changed)
        self.tree_widget.itemDoubleClicked.connect(self._on_double_click)

        self.new_section_radio.toggled.connect(self._on_add_option_changed)
        self.existing_section_radio.toggled.connect(self._on_add_option_changed)

        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

    def _on_theme_changed(self, index: int) -> None:
        """Handle theme selection change."""
        theme_path = self.theme_combo.itemData(index)
        if theme_path:
            self._load_theme(theme_path)

    def _on_browse_theme(self) -> None:
        """Browse for a theme file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("dialog.theme.browse_title"),
            self.settings.get_themes_path(self.base_path),
            "PowerPoint files (*.pptx *.ppt);;All files (*.*)"
        )

        if file_path:
            # Add to combo if not already there
            index = self.theme_combo.findData(file_path)
            if index < 0:
                name = os.path.basename(file_path)
                self.theme_combo.addItem(name, file_path)
                index = self.theme_combo.count() - 1

            self.theme_combo.setCurrentIndex(index)

    def _load_theme(self, theme_path: str) -> None:
        """Load sections from a theme file into the tree."""
        self.tree_widget.clear()

        try:
            sections = self.theme_service.get_sections_from_theme(theme_path)

            for section in sections:
                section_item = QTreeWidgetItem()
                section_item.setText(0, f"📁 {section.name}")
                section_item.setData(0, Qt.ItemDataRole.UserRole, self.ITEM_TYPE_SECTION)
                section_item.setData(0, Qt.ItemDataRole.UserRole + 1, section)
                section_item.setFlags(
                    Qt.ItemFlag.ItemIsEnabled |
                    Qt.ItemFlag.ItemIsSelectable
                )

                self.tree_widget.addTopLevelItem(section_item)

                # Add slides as children
                for slide in section.slides:
                    slide_item = QTreeWidgetItem()
                    slide_item.setText(0, f"  └─ {slide.title}")
                    slide_item.setData(0, Qt.ItemDataRole.UserRole, self.ITEM_TYPE_SLIDE)
                    slide_item.setData(0, Qt.ItemDataRole.UserRole + 1, slide)
                    slide_item.setData(0, Qt.ItemDataRole.UserRole + 2, section)  # Parent section
                    slide_item.setFlags(
                        Qt.ItemFlag.ItemIsEnabled |
                        Qt.ItemFlag.ItemIsSelectable
                    )

                    section_item.addChild(slide_item)

                section_item.setExpanded(True)

        except Exception as e:
            QMessageBox.warning(
                self,
                tr("dialog.theme.title"),
                tr("error.load_failed", error=str(e))
            )

    def _on_search_changed(self, text: str) -> None:
        """Filter tree items based on search text."""
        search_lower = text.lower()

        for i in range(self.tree_widget.topLevelItemCount()):
            section_item = self.tree_widget.topLevelItem(i)
            section_visible = search_lower in section_item.text(0).lower()

            # Check children
            child_visible = False
            for j in range(section_item.childCount()):
                child = section_item.child(j)
                if search_lower in child.text(0).lower():
                    child_visible = True
                    child.setHidden(False)
                else:
                    child.setHidden(not section_visible)

            section_item.setHidden(not (section_visible or child_visible))

    def _on_selection_changed(self) -> None:
        """Handle selection change."""
        selected = self.tree_widget.selectedItems()

        section_count = 0
        slide_count = 0

        for item in selected:
            item_type = item.data(0, Qt.ItemDataRole.UserRole)
            if item_type == self.ITEM_TYPE_SECTION:
                section_count += 1
            elif item_type == self.ITEM_TYPE_SLIDE:
                slide_count += 1

        # Update info label
        if section_count > 0 or slide_count > 0:
            parts = []
            if section_count > 0:
                parts.append(tr("dialog.theme.sections_selected", count=section_count))
            if slide_count > 0:
                parts.append(tr("dialog.theme.slides_selected", count=slide_count))
            self.info_label.setText(", ".join(parts))
        else:
            self.info_label.setText("")

        # Update OK button
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setEnabled(
            section_count > 0 or slide_count > 0
        )

    def _on_double_click(self, item: QTreeWidgetItem, column: int) -> None:
        """Handle double-click to accept selection."""
        if item.data(0, Qt.ItemDataRole.UserRole) in (self.ITEM_TYPE_SECTION, self.ITEM_TYPE_SLIDE):
            self.accept()

    def _on_add_option_changed(self, checked: bool) -> None:
        """Handle add option radio button change."""
        self.existing_combo.setEnabled(self.existing_section_radio.isChecked())

        # Ensure mutual exclusivity
        if self.sender() == self.new_section_radio and checked:
            self.existing_section_radio.setChecked(False)
        elif self.sender() == self.existing_section_radio and checked:
            self.new_section_radio.setChecked(False)

    def get_selected_sections(self) -> List[LiturgySection]:
        """Get the selected sections (new sections or with slides added)."""
        selected = self.tree_widget.selectedItems()
        result_sections = []
        slides_to_add = []  # For adding to existing section

        # First pass: collect all selected sections
        selected_section_ids = set()
        for item in selected:
            item_type = item.data(0, Qt.ItemDataRole.UserRole)
            if item_type == self.ITEM_TYPE_SECTION:
                section: LiturgySection = item.data(0, Qt.ItemDataRole.UserRole + 1)
                # Deep copy the section
                new_section = LiturgySection(
                    name=section.name,
                    section_type=section.section_type,
                    source_theme_path=section.source_theme_path,
                    slides=list(section.slides),
                )
                result_sections.append(new_section)
                selected_section_ids.add(id(section))

        # Second pass: collect individual slides not part of selected sections
        for item in selected:
            item_type = item.data(0, Qt.ItemDataRole.UserRole)
            if item_type == self.ITEM_TYPE_SLIDE:
                slide: LiturgySlide = item.data(0, Qt.ItemDataRole.UserRole + 1)
                parent_section: LiturgySection = item.data(0, Qt.ItemDataRole.UserRole + 2)

                # Skip if parent section is already selected
                if id(parent_section) in selected_section_ids:
                    continue

                if self.existing_section_radio.isChecked():
                    # Collect for adding to existing section
                    slides_to_add.append(slide)
                else:
                    # Create new section for individual slide
                    new_section = LiturgySection(
                        name=slide.title,
                        section_type=SectionType.REGULAR,
                        source_theme_path=parent_section.source_theme_path,
                        slides=[slide],
                    )
                    result_sections.append(new_section)

        return result_sections

    def get_slides_for_existing(self) -> Tuple[Optional[str], List[LiturgySlide]]:
        """Get slides to add to an existing section."""
        if not self.existing_section_radio.isChecked():
            return (None, [])

        section_id = self.existing_combo.currentData()
        if not section_id:
            return (None, [])

        selected = self.tree_widget.selectedItems()
        slides = []

        for item in selected:
            item_type = item.data(0, Qt.ItemDataRole.UserRole)
            if item_type == self.ITEM_TYPE_SLIDE:
                slide: LiturgySlide = item.data(0, Qt.ItemDataRole.UserRole + 1)
                slides.append(slide)
            elif item_type == self.ITEM_TYPE_SECTION:
                section: LiturgySection = item.data(0, Qt.ItemDataRole.UserRole + 1)
                slides.extend(section.slides)

        return (section_id, slides)

    def is_adding_to_existing(self) -> bool:
        """Check if adding to existing section."""
        return self.existing_section_radio.isChecked()
