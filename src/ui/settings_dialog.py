"""Settings dialog for configuring application settings."""

import os
from typing import Optional

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QLineEdit,
    QPushButton,
    QComboBox,
    QCheckBox,
    QDialogButtonBox,
    QGroupBox,
    QFileDialog,
    QLabel,
    QSpinBox,
)
from PyQt6.QtCore import Qt

from ..models import Settings
from ..i18n import tr


class SettingsDialog(QDialog):
    """Dialog for editing application settings."""

    def __init__(self, settings: Settings, base_path: str, parent=None):
        super().__init__(parent)
        self.settings = Settings(
            base_folder=settings.base_folder,
            songs_folder=settings.songs_folder,
            algemeen_folder=settings.algemeen_folder,
            output_folder=settings.output_folder,
            themes_folder=settings.themes_folder,
            collecte_filename=settings.collecte_filename,
            stub_template_filename=settings.stub_template_filename,
            output_pattern=settings.output_pattern,
            language=settings.language,
            song_cover_enabled=settings.song_cover_enabled,
            song_cover_filename=settings.song_cover_filename,
            excel_register_path=settings.excel_register_path,
            pptx_archive_folder=settings.pptx_archive_folder,
            bible_font_name=settings.bible_font_name,
            bible_font_size=settings.bible_font_size,
            bible_show_verse_numbers=settings.bible_show_verse_numbers,
            youversion_api_key=settings.youversion_api_key,
            window_width=settings.window_width,
            window_height=settings.window_height,
        )
        self.base_path = base_path

        self._setup_ui()
        self._load_settings()
        self._connect_signals()

    def _setup_ui(self) -> None:
        """Setup the dialog UI."""
        self.setWindowTitle(tr("dialog.settings.title"))
        self.setMinimumSize(500, 450)
        self.resize(600, 500)

        layout = QVBoxLayout(self)

        # Base folder group
        base_group = QGroupBox(tr("dialog.settings.base_folder_group"))
        base_layout = QFormLayout(base_group)

        self.base_folder_input = QLineEdit()
        base_browse = QPushButton(tr("button.browse"))
        base_browse.clicked.connect(self._browse_base_folder)
        base_folder_layout = QHBoxLayout()
        base_folder_layout.addWidget(self.base_folder_input)
        base_folder_layout.addWidget(base_browse)
        base_layout.addRow(tr("dialog.settings.base_folder"), base_folder_layout)

        layout.addWidget(base_group)

        # Folders group
        folders_group = QGroupBox(tr("dialog.settings.folders"))
        folders_layout = QFormLayout(folders_group)

        # Songs folder
        self.songs_folder_input = QLineEdit()
        songs_browse = QPushButton(tr("button.browse"))
        songs_browse.clicked.connect(lambda: self._browse_folder(self.songs_folder_input))
        songs_layout = QHBoxLayout()
        songs_layout.addWidget(self.songs_folder_input)
        songs_layout.addWidget(songs_browse)
        folders_layout.addRow(tr("dialog.settings.songs_folder"), songs_layout)

        # Algemeen folder
        self.algemeen_folder_input = QLineEdit()
        algemeen_browse = QPushButton(tr("button.browse"))
        algemeen_browse.clicked.connect(lambda: self._browse_folder(self.algemeen_folder_input))
        algemeen_layout = QHBoxLayout()
        algemeen_layout.addWidget(self.algemeen_folder_input)
        algemeen_layout.addWidget(algemeen_browse)
        folders_layout.addRow(tr("dialog.settings.algemeen_folder"), algemeen_layout)

        # Themes folder
        self.themes_folder_input = QLineEdit()
        themes_browse = QPushButton(tr("button.browse"))
        themes_browse.clicked.connect(lambda: self._browse_folder(self.themes_folder_input))
        themes_layout = QHBoxLayout()
        themes_layout.addWidget(self.themes_folder_input)
        themes_layout.addWidget(themes_browse)
        folders_layout.addRow(tr("dialog.settings.themes_folder"), themes_layout)

        # Output folder
        self.output_folder_input = QLineEdit()
        output_browse = QPushButton(tr("button.browse"))
        output_browse.clicked.connect(lambda: self._browse_folder(self.output_folder_input))
        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_folder_input)
        output_layout.addWidget(output_browse)
        folders_layout.addRow(tr("dialog.settings.output_folder"), output_layout)

        # PPTX archive folder
        self.pptx_archive_input = QLineEdit()
        pptx_archive_browse = QPushButton(tr("button.browse"))
        pptx_archive_browse.clicked.connect(lambda: self._browse_folder(self.pptx_archive_input))
        pptx_archive_layout = QHBoxLayout()
        pptx_archive_layout.addWidget(self.pptx_archive_input)
        pptx_archive_layout.addWidget(pptx_archive_browse)
        folders_layout.addRow(tr("dialog.settings.pptx_archive_folder"), pptx_archive_layout)

        layout.addWidget(folders_group)

        # Files group
        files_group = QGroupBox(tr("dialog.settings.files"))
        files_layout = QFormLayout(files_group)

        self.collecte_input = QLineEdit()
        files_layout.addRow(tr("dialog.settings.collecte_file"), self.collecte_input)

        self.stub_template_input = QLineEdit()
        files_layout.addRow(tr("dialog.settings.stub_template"), self.stub_template_input)

        self.output_pattern_input = QLineEdit()
        files_layout.addRow(tr("dialog.settings.output_pattern"), self.output_pattern_input)

        # Excel register file
        self.excel_register_input = QLineEdit()
        excel_browse = QPushButton(tr("button.browse_file"))
        excel_browse.clicked.connect(self._browse_excel_file)
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(self.excel_register_input)
        excel_layout.addWidget(excel_browse)
        files_layout.addRow(tr("dialog.settings.excel_register"), excel_layout)

        # Song cover slide
        self.song_cover_checkbox = QCheckBox(tr("dialog.settings.song_cover_enabled"))
        files_layout.addRow("", self.song_cover_checkbox)

        self.song_cover_input = QLineEdit()
        song_cover_browse = QPushButton(tr("button.browse_file"))
        song_cover_browse.clicked.connect(self._browse_song_cover_file)
        song_cover_layout = QHBoxLayout()
        song_cover_layout.addWidget(self.song_cover_input)
        song_cover_layout.addWidget(song_cover_browse)
        files_layout.addRow(tr("dialog.settings.song_cover_file"), song_cover_layout)

        self.song_cover_checkbox.toggled.connect(self.song_cover_input.setEnabled)
        self.song_cover_checkbox.toggled.connect(song_cover_browse.setEnabled)

        layout.addWidget(files_group)

        # Bible text group
        bible_group = QGroupBox(tr("dialog.settings.bible_group"))
        bible_layout = QFormLayout(bible_group)

        self.bible_font_name_input = QLineEdit()
        bible_layout.addRow(tr("dialog.settings.bible_font_name"), self.bible_font_name_input)

        self.bible_font_size_spin = QSpinBox()
        self.bible_font_size_spin.setRange(6, 36)
        self.bible_font_size_spin.setSuffix(" pt")
        bible_layout.addRow(tr("dialog.settings.bible_font_size"), self.bible_font_size_spin)

        self.bible_chars_per_slide_spin = QSpinBox()
        self.bible_chars_per_slide_spin.setRange(100, 2000)
        self.bible_chars_per_slide_spin.setSingleStep(50)
        self.bible_chars_per_slide_spin.setSuffix(" chars")
        bible_layout.addRow(tr("dialog.settings.bible_chars_per_slide"), self.bible_chars_per_slide_spin)

        self.bible_show_verse_numbers_check = QCheckBox(tr("dialog.settings.bible_show_verse_numbers"))
        bible_layout.addRow("", self.bible_show_verse_numbers_check)

        self.youversion_api_key_input = QLineEdit()
        self.youversion_api_key_input.setPlaceholderText(
            tr("dialog.settings.youversion_api_key_placeholder")
        )
        self.youversion_api_key_input.setEchoMode(QLineEdit.EchoMode.Password)
        bible_layout.addRow(tr("dialog.settings.youversion_api_key"), self.youversion_api_key_input)

        api_key_hint = QLabel(tr("dialog.settings.youversion_api_key_hint"))
        api_key_hint.setStyleSheet("color: grey; font-size: 11px;")
        api_key_hint.setWordWrap(True)
        bible_layout.addRow("", api_key_hint)

        layout.addWidget(bible_group)

        # Language group
        lang_layout = QHBoxLayout()
        lang_label = QLabel(tr("dialog.settings.language"))
        self.language_combo = QComboBox()
        self.language_combo.addItem(tr("language.nl"), "nl")
        self.language_combo.addItem(tr("language.en"), "en")
        lang_layout.addWidget(lang_label)
        lang_layout.addWidget(self.language_combo)
        lang_layout.addStretch()
        layout.addLayout(lang_layout)

        layout.addStretch()

        # Button box
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.button_box.button(QDialogButtonBox.StandardButton.Ok).setText(tr("button.save"))
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText(tr("button.cancel"))
        layout.addWidget(self.button_box)

    def _load_settings(self) -> None:
        """Load current settings into the form."""
        self.base_folder_input.setText(self.settings.base_folder)
        self.songs_folder_input.setText(self.settings.songs_folder)
        self.algemeen_folder_input.setText(self.settings.algemeen_folder)
        self.themes_folder_input.setText(self.settings.themes_folder)
        self.output_folder_input.setText(self.settings.output_folder)
        self.collecte_input.setText(self.settings.collecte_filename)
        self.stub_template_input.setText(self.settings.stub_template_filename)
        self.output_pattern_input.setText(self.settings.output_pattern)
        self.excel_register_input.setText(self.settings.excel_register_path)
        self.pptx_archive_input.setText(self.settings.pptx_archive_folder)

        # Song cover
        self.song_cover_checkbox.setChecked(self.settings.song_cover_enabled)
        self.song_cover_input.setText(self.settings.song_cover_filename)
        self.song_cover_input.setEnabled(self.settings.song_cover_enabled)

        # Bible text settings
        self.bible_font_name_input.setText(self.settings.bible_font_name)
        self.bible_font_size_spin.setValue(self.settings.bible_font_size)
        self.bible_chars_per_slide_spin.setValue(self.settings.bible_chars_per_slide)
        self.bible_show_verse_numbers_check.setChecked(self.settings.bible_show_verse_numbers)
        self.youversion_api_key_input.setText(self.settings.youversion_api_key)

        # Set language combo
        index = self.language_combo.findData(self.settings.language)
        if index >= 0:
            self.language_combo.setCurrentIndex(index)

    def _connect_signals(self) -> None:
        """Connect widget signals."""
        self.button_box.accepted.connect(self._on_accept)
        self.button_box.rejected.connect(self.reject)

    def _get_effective_base(self) -> str:
        """Get the effective base folder for browsing."""
        base = self.base_folder_input.text()
        if base and os.path.isdir(base):
            return base
        return self.base_path

    def _browse_base_folder(self) -> None:
        """Open folder browser for base folder."""
        current = self.base_folder_input.text()
        if not current or not os.path.isdir(current):
            current = self.base_path

        folder = QFileDialog.getExistingDirectory(
            self,
            tr("dialog.settings.base_folder"),
            current
        )

        if folder:
            self.base_folder_input.setText(folder)

    def _browse_folder(self, line_edit: QLineEdit) -> None:
        """Open folder browser dialog."""
        current = line_edit.text()
        base = self._get_effective_base()
        if current and not os.path.isabs(current):
            current = os.path.join(base, current)

        folder = QFileDialog.getExistingDirectory(
            self,
            tr("button.browse"),
            current if os.path.exists(current) else base
        )

        if folder:
            # Try to make relative path to base folder
            try:
                rel_path = os.path.relpath(folder, base)
                if not rel_path.startswith(".."):
                    folder = "./" + rel_path.replace("\\", "/")
            except ValueError:
                pass  # Different drives on Windows

            line_edit.setText(folder)

    def _browse_excel_file(self) -> None:
        """Open file browser for Excel register file."""
        current = self.excel_register_input.text()
        base = self._get_effective_base()
        if current and not os.path.isabs(current):
            current = os.path.join(base, current)

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("dialog.settings.excel_register"),
            current if current and os.path.exists(os.path.dirname(current)) else base,
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )

        if file_path:
            # Try to make relative path to base folder
            try:
                rel_path = os.path.relpath(file_path, base)
                if not rel_path.startswith(".."):
                    file_path = "./" + rel_path.replace("\\", "/")
            except ValueError:
                pass  # Different drives on Windows

            self.excel_register_input.setText(file_path)

    def _browse_song_cover_file(self) -> None:
        """Open file browser for song cover slide PPTX."""
        base = self._get_effective_base()
        algemeen = self.settings.get_algemeen_path(base)
        start_dir = algemeen if os.path.isdir(algemeen) else base

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("dialog.settings.song_cover_file"),
            start_dir,
            "PowerPoint files (*.pptx);;All files (*.*)"
        )

        if file_path:
            # Store just the filename (file lives in algemeen folder)
            self.song_cover_input.setText(os.path.basename(file_path))

    def _on_accept(self) -> None:
        """Handle dialog acceptance."""
        self.settings.base_folder = self.base_folder_input.text()
        self.settings.songs_folder = self.songs_folder_input.text()
        self.settings.algemeen_folder = self.algemeen_folder_input.text()
        self.settings.themes_folder = self.themes_folder_input.text()
        self.settings.output_folder = self.output_folder_input.text()
        self.settings.collecte_filename = self.collecte_input.text()
        self.settings.stub_template_filename = self.stub_template_input.text()
        self.settings.output_pattern = self.output_pattern_input.text()
        self.settings.excel_register_path = self.excel_register_input.text()
        self.settings.pptx_archive_folder = self.pptx_archive_input.text()
        self.settings.song_cover_enabled = self.song_cover_checkbox.isChecked()
        self.settings.song_cover_filename = self.song_cover_input.text()
        self.settings.bible_font_name = self.bible_font_name_input.text().strip() or "Calibri"
        self.settings.bible_font_size = self.bible_font_size_spin.value()
        self.settings.bible_chars_per_slide = self.bible_chars_per_slide_spin.value()
        self.settings.bible_show_verse_numbers = self.bible_show_verse_numbers_check.isChecked()
        self.settings.youversion_api_key = self.youversion_api_key_input.text().strip()
        self.settings.language = self.language_combo.currentData()

        self.accept()

    def get_settings(self) -> Settings:
        """Get the modified settings."""
        return self.settings
