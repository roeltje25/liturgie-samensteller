"""Dialog for editing fillable fields in slides."""

import os
from typing import Dict, List, Optional

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QTableWidget,
    QTableWidgetItem,
    QLineEdit,
    QPushButton,
    QLabel,
    QDialogButtonBox,
    QGroupBox,
    QHeaderView,
    QComboBox,
    QMessageBox,
)
from PyQt6.QtCore import Qt

from ..models import (
    Liturgy,
    LiturgySection,
    LiturgySlide,
)
from ..services import PptxService, SlideField
from ..i18n import tr


class SlideFieldEditor(QDialog):
    """Dialog for editing fields in a single slide."""

    def __init__(
        self,
        slide: LiturgySlide,
        section: LiturgySection,
        pptx_service: PptxService,
        parent=None
    ):
        super().__init__(parent)
        self.slide = slide
        self.section = section
        self.pptx_service = pptx_service

        self._fields: Dict[str, str] = dict(slide.fields)
        self._available_fields: List[SlideField] = []

        self._setup_ui()
        self._load_fields()
        self._connect_signals()

    def _setup_ui(self) -> None:
        """Setup the dialog UI."""
        self.setWindowTitle(tr("dialog.fields.slide_title", slide=self.slide.title))
        self.setMinimumSize(500, 400)
        self.resize(550, 450)

        layout = QVBoxLayout(self)

        # Info label
        info_text = f"{self.section.name} - {self.slide.title}"
        self.info_label = QLabel(info_text)
        self.info_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.info_label)

        # Fields table
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels([
            tr("dialog.fields.field_name"),
            tr("dialog.fields.value")
        ])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.table)

        # Add field button
        add_layout = QHBoxLayout()
        self.add_field_combo = QComboBox()
        self.add_field_combo.setMinimumWidth(200)
        self.add_field_btn = QPushButton(tr("button.add_field"))
        add_layout.addWidget(QLabel(tr("dialog.fields.add_field")))
        add_layout.addWidget(self.add_field_combo)
        add_layout.addWidget(self.add_field_btn)
        add_layout.addStretch()
        layout.addLayout(add_layout)

        # Button box
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setText(tr("button.save"))
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText(tr("button.cancel"))
        layout.addWidget(self.button_box)

    def _load_fields(self) -> None:
        """Load fields from the slide source."""
        # Extract available fields from the source PPTX
        if self.slide.source_path and os.path.exists(self.slide.source_path):
            self._available_fields = self.pptx_service.extract_fields(
                self.slide.source_path,
                self.slide.slide_index
            )

        # Auto-add text pattern fields (like {Bidders}) to the fields dict
        # These are the user-defined fields we want to show directly
        for field in self._available_fields:
            if field.field_type == "text_pattern" and field.name not in self._fields:
                # Add with empty value (user will fill in)
                self._fields[field.name] = ""

        # Populate the combo box with placeholder fields only (optional to add)
        # Text pattern fields are already shown in the table
        self.add_field_combo.clear()
        added_names = set(self._fields.keys())
        for field in self._available_fields:
            # Only show placeholder fields in the dropdown (not auto-added)
            if field.field_type == "placeholder" and field.name not in added_names:
                self.add_field_combo.addItem(field.name, field)

        # Add custom field option
        self.add_field_combo.addItem(tr("dialog.fields.custom"), None)

        # Populate table with current field values
        self._populate_table()

    def _populate_table(self) -> None:
        """Populate the table with current fields."""
        self.table.setRowCount(len(self._fields))

        for row, (name, value) in enumerate(self._fields.items()):
            # Field name (read-only)
            name_item = QTableWidgetItem(name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 0, name_item)

            # Field value (editable)
            value_item = QTableWidgetItem(value)
            self.table.setItem(row, 1, value_item)

    def _connect_signals(self) -> None:
        """Connect widget signals."""
        self.add_field_btn.clicked.connect(self._on_add_field)
        self.button_box.accepted.connect(self._on_save)
        self.button_box.rejected.connect(self.reject)

    def _on_add_field(self) -> None:
        """Add a new field to the table."""
        field = self.add_field_combo.currentData()

        if field is None:
            # Custom field - ask for name
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

        # Add to fields dict and table
        self._fields[field_name] = ""

        row = self.table.rowCount()
        self.table.setRowCount(row + 1)

        name_item = QTableWidgetItem(field_name)
        name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.table.setItem(row, 0, name_item)

        value_item = QTableWidgetItem("")
        self.table.setItem(row, 1, value_item)

        # Update combo box
        idx = self.add_field_combo.findText(field_name)
        if idx >= 0:
            self.add_field_combo.removeItem(idx)

    def _on_save(self) -> None:
        """Save fields and close."""
        # Update fields from table
        self._fields.clear()
        for row in range(self.table.rowCount()):
            name = self.table.item(row, 0).text()
            value = self.table.item(row, 1).text()
            if value:  # Only save non-empty values
                self._fields[name] = value

        self.accept()

    def get_fields(self) -> Dict[str, str]:
        """Get the edited fields."""
        return self._fields


class BulkFieldEditor(QDialog):
    """Dialog for editing all fields across the entire liturgy."""

    def __init__(
        self,
        liturgy: Liturgy,
        pptx_service: PptxService,
        parent=None
    ):
        super().__init__(parent)
        self.liturgy = liturgy
        self.pptx_service = pptx_service

        self._all_fields: Dict[str, str] = {}  # Shared field values
        self._field_locations: Dict[str, List[tuple]] = {}  # field_name -> [(section, slide), ...]

        self._setup_ui()
        self._scan_fields()
        self._connect_signals()

    def _setup_ui(self) -> None:
        """Setup the dialog UI."""
        self.setWindowTitle(tr("dialog.fields.bulk_title"))
        self.setMinimumSize(600, 500)
        self.resize(700, 600)

        layout = QVBoxLayout(self)

        # Instructions
        self.instruction_label = QLabel(tr("dialog.fields.bulk_instructions"))
        layout.addWidget(self.instruction_label)

        # Common fields section
        common_group = QGroupBox(tr("dialog.fields.common_fields"))
        common_layout = QVBoxLayout(common_group)

        self.common_table = QTableWidget()
        self.common_table.setColumnCount(3)
        self.common_table.setHorizontalHeaderLabels([
            tr("dialog.fields.field_name"),
            tr("dialog.fields.value"),
            tr("dialog.fields.occurrences")
        ])
        self.common_table.horizontalHeader().setStretchLastSection(True)
        self.common_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.common_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        common_layout.addWidget(self.common_table)

        layout.addWidget(common_group)

        # Quick date/info fields
        quick_group = QGroupBox(tr("dialog.fields.quick_fill"))
        quick_layout = QHBoxLayout(quick_group)

        # Date field - pre-fill with service date if available
        quick_layout.addWidget(QLabel(tr("dialog.fields.date")))
        self.date_input = QLineEdit()
        self.date_input.setPlaceholderText("26 januari 2026")
        # Pre-fill with service date formatted nicely
        if self.liturgy.service_date:
            try:
                from datetime import date
                import locale
                service_date = date.fromisoformat(self.liturgy.service_date)
                # Format as "26 januari 2026" style
                try:
                    locale.setlocale(locale.LC_TIME, '')  # Use system locale
                    formatted_date = service_date.strftime("%d %B %Y").lstrip('0')
                except Exception:
                    formatted_date = service_date.strftime("%d %B %Y").lstrip('0')
                self.date_input.setText(formatted_date)
            except Exception:
                pass
        quick_layout.addWidget(self.date_input)

        self.apply_date_btn = QPushButton(tr("dialog.fields.apply"))
        quick_layout.addWidget(self.apply_date_btn)

        quick_layout.addStretch()
        layout.addWidget(quick_group)

        # Button box
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setText(tr("button.save"))
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText(tr("button.cancel"))
        layout.addWidget(self.button_box)

    def _scan_fields(self) -> None:
        """Scan all slides for text pattern fields (like {Bidders})."""
        self._all_fields.clear()
        self._field_locations.clear()

        for section in self.liturgy.sections:
            for slide in section.slides:
                # Get fields from slide's current values (these are user-set values)
                for name, value in slide.fields.items():
                    if name not in self._all_fields:
                        self._all_fields[name] = value
                        self._field_locations[name] = []
                    self._field_locations[name].append((section, slide))

                # Also extract text pattern fields from source PPTX
                # Skip placeholder fields (TITLE, BODY, etc.) - only show {FieldName} patterns
                if slide.source_path and os.path.exists(slide.source_path):
                    try:
                        fields = self.pptx_service.extract_fields(
                            slide.source_path,
                            slide.slide_index
                        )
                        for field in fields:
                            # Only include text_pattern fields, not placeholders
                            if field.field_type == "text_pattern":
                                if field.name not in self._all_fields:
                                    self._all_fields[field.name] = ""
                                    self._field_locations[field.name] = []
                                if (section, slide) not in self._field_locations[field.name]:
                                    self._field_locations[field.name].append((section, slide))
                    except Exception:
                        pass

        self._populate_table()

    def _populate_table(self) -> None:
        """Populate the common fields table."""
        self.common_table.setRowCount(len(self._all_fields))

        for row, (name, value) in enumerate(sorted(self._all_fields.items())):
            # Field name
            name_item = QTableWidgetItem(name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.common_table.setItem(row, 0, name_item)

            # Value
            value_item = QTableWidgetItem(value)
            self.common_table.setItem(row, 1, value_item)

            # Occurrences count
            count = len(self._field_locations.get(name, []))
            count_item = QTableWidgetItem(str(count))
            count_item.setFlags(count_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.common_table.setItem(row, 2, count_item)

    def _connect_signals(self) -> None:
        """Connect widget signals."""
        self.apply_date_btn.clicked.connect(self._on_apply_date)
        self.button_box.accepted.connect(self._on_save)
        self.button_box.rejected.connect(self.reject)

    def _on_apply_date(self) -> None:
        """Apply date to common date fields."""
        date_value = self.date_input.text()
        if not date_value:
            return

        # Find date-related field names
        date_fields = ["DATUM", "DATE", "DATEUM", "DAG"]

        for row in range(self.common_table.rowCount()):
            name = self.common_table.item(row, 0).text()
            if name.upper() in date_fields:
                self.common_table.item(row, 1).setText(date_value)

    def _on_save(self) -> None:
        """Save all field values to slides."""
        # Read values from table
        new_values = {}
        for row in range(self.common_table.rowCount()):
            name = self.common_table.item(row, 0).text()
            value = self.common_table.item(row, 1).text()
            if value:
                new_values[name] = value

        # Apply to all slides
        for name, value in new_values.items():
            locations = self._field_locations.get(name, [])
            for section, slide in locations:
                slide.fields[name] = value

        self.accept()

    def get_updated_liturgy(self) -> Liturgy:
        """Get the liturgy with updated fields."""
        return self.liturgy
