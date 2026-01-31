"""Main application window."""

import os
from datetime import date
from typing import Optional

from PyQt6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QGroupBox,
    QStatusBar,
    QMenuBar,
    QMenu,
    QFileDialog,
    QMessageBox,
    QComboBox,
    QSplitter,
    QFormLayout,
    QDateEdit,
    QLineEdit,
    QCompleter,
    QDialog,
    QDialogButtonBox,
)
from PyQt6.QtCore import Qt, QSize, QDate
from PyQt6.QtGui import QAction, QKeySequence

from ..models import (
    Settings,
    Liturgy,
    LiturgySection,
    LiturgySlide,
    SectionType,
    SongLiturgyItem,
    GenericLiturgyItem,
    OfferingLiturgyItem,
)
from ..services import FolderScanner, ExportService, YouTubeService, ThemeService, PptxService
from ..i18n import tr, set_language, get_language, on_language_changed

from .liturgy_list import LiturgyListWidget
from .liturgy_tree import LiturgyTreeWidget
from .song_picker import SongPickerDialog
from .new_song_dialog import NewSongDialog
from .generic_picker import GenericPickerDialog
from .offering_picker import OfferingPickerDialog
from .settings_dialog import SettingsDialog
from .export_dialog import ExportDialog
from .youtube_dialog import YouTubeDialog
from .field_editor import SlideFieldEditor, BulkFieldEditor
from .theme_picker import ThemeSectionPicker
from .section_editor import SectionEditorDialog
from .about_dialog import AboutDialog


class MainWindow(QMainWindow):
    """Main application window."""

    def __init__(self, base_path: str = ".", skip_first_run_check: bool = False):
        super().__init__()
        self.base_path = base_path
        self.settings = Settings.load()

        # Check for first run and show setup dialog (unless handled by main.py)
        if not skip_first_run_check and self.settings.is_first_run():
            self._show_first_run_dialog()

        self.liturgy = Liturgy(name=self._generate_default_name())
        self.current_file: Optional[str] = None
        self.unsaved_changes = False

        # Services
        self.folder_scanner = FolderScanner(self.settings, base_path)
        self.export_service = ExportService(self.settings, base_path)
        self.youtube_service = YouTubeService()
        self.theme_service = ThemeService(self.settings, base_path)
        self.pptx_service = PptxService(self.settings, base_path)

        # Set language from settings
        set_language(self.settings.language)

        # Setup UI
        self._setup_ui()
        self._setup_menu()
        self._setup_statusbar()
        self._connect_signals()

        # Apply translations
        self._update_translations()

        # Check for configuration warnings
        self._check_warnings()

        # Connect language change
        on_language_changed(self._on_language_changed)

    def _show_first_run_dialog(self) -> None:
        """Show first-run dialog to select base folder."""
        from PyQt6.QtWidgets import QMessageBox, QFileDialog

        # Show welcome message
        msg = QMessageBox(self)
        msg.setWindowTitle(tr("dialog.firstrun.title"))
        msg.setText(tr("dialog.firstrun.message"))
        msg.setIcon(QMessageBox.Icon.Information)
        msg.addButton(tr("dialog.firstrun.select_folder"), QMessageBox.ButtonRole.AcceptRole)
        msg.exec()

        # Open folder selection
        folder = QFileDialog.getExistingDirectory(
            self,
            tr("dialog.settings.base_folder"),
            self.base_path
        )

        if folder:
            self.settings.base_folder = folder
            self.settings.save()

    def _generate_default_name(self) -> str:
        """Generate default liturgy name."""
        return f"Viering {date.today().strftime('%Y-%m-%d')}"

    def _set_next_sunday(self) -> None:
        """Set the service date to next Sunday."""
        from datetime import timedelta
        today = date.today()
        days_until_sunday = (6 - today.weekday()) % 7
        if days_until_sunday == 0:
            days_until_sunday = 7  # If today is Sunday, get next Sunday
        next_sunday = today + timedelta(days=days_until_sunday)
        self.service_date_edit.setDate(QDate(next_sunday.year, next_sunday.month, next_sunday.day))
        # Also set liturgy service_date (signal not connected yet during init)
        self.liturgy.service_date = next_sunday.strftime("%Y-%m-%d")

    def _setup_dienstleider_autocomplete(self) -> None:
        """Setup autocomplete for dienstleider field from Excel."""
        excel_path = self.settings.get_excel_register_path(self.base_path)
        if excel_path and os.path.exists(excel_path):
            dienstleiders = self.export_service.get_excel_dienstleiders(excel_path)
            if dienstleiders:
                completer = QCompleter(dienstleiders)
                completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
                self.dienstleider_edit.setCompleter(completer)

    def _setup_ui(self) -> None:
        """Setup the main UI layout."""
        self.setMinimumSize(QSize(900, 600))
        self.resize(self.settings.window_width, self.settings.window_height)

        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)

        # Create splitter for resizable panels
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # Left panel - Add items
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)

        self.add_items_group = QGroupBox()
        add_layout = QVBoxLayout(self.add_items_group)

        self.btn_add_song = QPushButton()
        self.btn_add_song.setMinimumHeight(40)
        self.btn_create_song = QPushButton()
        self.btn_create_song.setMinimumHeight(40)
        self.btn_add_generic = QPushButton()
        self.btn_add_generic.setMinimumHeight(40)
        self.btn_add_offering = QPushButton()
        self.btn_add_offering.setMinimumHeight(40)
        self.btn_add_from_theme = QPushButton()
        self.btn_add_from_theme.setMinimumHeight(40)
        self.btn_add_section = QPushButton()
        self.btn_add_section.setMinimumHeight(40)
        self.btn_add_pptx = QPushButton()
        self.btn_add_pptx.setMinimumHeight(40)

        add_layout.addWidget(self.btn_add_song)
        add_layout.addWidget(self.btn_create_song)
        add_layout.addWidget(self.btn_add_generic)
        add_layout.addWidget(self.btn_add_offering)
        add_layout.addWidget(self.btn_add_from_theme)
        add_layout.addWidget(self.btn_add_section)
        add_layout.addWidget(self.btn_add_pptx)

        # Separator
        add_layout.addSpacing(20)

        # Edit fields button
        self.btn_edit_fields = QPushButton()
        self.btn_edit_fields.setMinimumHeight(40)
        add_layout.addWidget(self.btn_edit_fields)

        add_layout.addStretch()

        left_layout.addWidget(self.add_items_group)
        splitter.addWidget(left_panel)

        # Right panel - Current liturgy
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)

        # Service Info Panel (Phase 2)
        self.service_info_group = QGroupBox()
        service_info_layout = QFormLayout(self.service_info_group)

        # Service date picker (default: next Sunday)
        self.service_date_label = QLabel()
        self.service_date_edit = QDateEdit()
        self.service_date_edit.setCalendarPopup(True)
        self.service_date_edit.setDisplayFormat("dd-MM-yyyy")
        self._set_next_sunday()
        service_info_layout.addRow(self.service_date_label, self.service_date_edit)

        # Dienstleider input with autocomplete
        self.dienstleider_label = QLabel()
        self.dienstleider_edit = QLineEdit()
        self.dienstleider_edit.setPlaceholderText("")  # Set in _update_translations
        self._setup_dienstleider_autocomplete()
        service_info_layout.addRow(self.dienstleider_label, self.dienstleider_edit)

        right_layout.addWidget(self.service_info_group)

        # Warning banner for missing offerings etc.
        self.warning_label = QLabel()
        self.warning_label.setStyleSheet("""
            QLabel {
                background-color: #fff3cd;
                color: #856404;
                border: 1px solid #ffc107;
                border-radius: 4px;
                padding: 8px;
                font-weight: bold;
            }
        """)
        self.warning_label.setWordWrap(True)
        self.warning_label.hide()  # Hidden by default
        right_layout.addWidget(self.warning_label)

        self.liturgy_group = QGroupBox()
        liturgy_layout = QVBoxLayout(self.liturgy_group)

        # Use tree widget for hierarchical display (v2 format)
        self.liturgy_tree = LiturgyTreeWidget(pptx_service=self.pptx_service)
        liturgy_layout.addWidget(self.liturgy_tree)

        # Keep list widget reference for backwards compatibility
        self.liturgy_list = self.liturgy_tree

        # Buttons below list
        btn_layout = QHBoxLayout()
        self.btn_delete = QPushButton()
        self.btn_edit = QPushButton()
        btn_layout.addWidget(self.btn_delete)
        btn_layout.addWidget(self.btn_edit)
        btn_layout.addStretch()
        liturgy_layout.addLayout(btn_layout)

        right_layout.addWidget(self.liturgy_group)
        splitter.addWidget(right_panel)

        # Set splitter proportions
        splitter.setSizes([250, 650])

    def _setup_menu(self) -> None:
        """Setup the menu bar."""
        menubar = self.menuBar()

        # File menu
        self.file_menu = menubar.addMenu("")
        self.action_new = self.file_menu.addAction("")
        self.action_new.setShortcut(QKeySequence.StandardKey.New)
        self.action_open = self.file_menu.addAction("")
        self.action_open.setShortcut(QKeySequence.StandardKey.Open)
        self.file_menu.addSeparator()
        self.action_save = self.file_menu.addAction("")
        self.action_save.setShortcut(QKeySequence.StandardKey.Save)
        self.action_save_as = self.file_menu.addAction("")
        self.action_save_as.setShortcut(QKeySequence.StandardKey.SaveAs)
        self.file_menu.addSeparator()
        self.action_export = self.file_menu.addAction("")
        self.action_export.setShortcut(QKeySequence("Ctrl+E"))
        self.file_menu.addSeparator()
        self.action_open_theme = self.file_menu.addAction("")
        self.action_save_as_theme = self.file_menu.addAction("")
        self.file_menu.addSeparator()
        self.action_exit = self.file_menu.addAction("")
        self.action_exit.setShortcut(QKeySequence.StandardKey.Quit)

        # Edit menu
        self.edit_menu = menubar.addMenu("")
        self.action_delete = self.edit_menu.addAction("")
        self.action_delete.setShortcut(QKeySequence.StandardKey.Delete)
        self.action_move_up = self.edit_menu.addAction("")
        self.action_move_up.setShortcut(QKeySequence("Ctrl+Up"))
        self.action_move_down = self.edit_menu.addAction("")
        self.action_move_down.setShortcut(QKeySequence("Ctrl+Down"))

        # Tools menu
        self.tools_menu = menubar.addMenu("")
        self.action_check_links = self.tools_menu.addAction("")
        self.action_edit_fields = self.tools_menu.addAction("")
        self.tools_menu.addSeparator()
        self.action_settings = self.tools_menu.addAction("")

        # Spacer for right-aligned language selector
        spacer = QWidget()
        spacer.setSizePolicy(
            spacer.sizePolicy().horizontalPolicy(),
            spacer.sizePolicy().verticalPolicy()
        )
        menubar.setCornerWidget(spacer, Qt.Corner.TopRightCorner)

        # Language selector in menu bar
        self.language_combo = QComboBox()
        self.language_combo.addItem("NL", "nl")
        self.language_combo.addItem("EN", "en")
        self.language_combo.setCurrentIndex(0 if get_language() == "nl" else 1)
        self.language_combo.setFixedWidth(60)
        menubar.setCornerWidget(self.language_combo, Qt.Corner.TopRightCorner)

        # Help menu
        self.help_menu = menubar.addMenu("")
        self.action_shortcuts = self.help_menu.addAction("")
        self.action_shortcuts.setShortcut(QKeySequence("F1"))
        self.action_workflow = self.help_menu.addAction("")
        self.help_menu.addSeparator()
        self.action_about = self.help_menu.addAction("")

    def _setup_statusbar(self) -> None:
        """Setup the status bar."""
        self.statusbar = QStatusBar()
        self.setStatusBar(self.statusbar)
        self.status_label = QLabel()
        self.statusbar.addWidget(self.status_label)

    def _connect_signals(self) -> None:
        """Connect all signals."""
        # Buttons
        self.btn_add_song.clicked.connect(self._on_add_song)
        self.btn_create_song.clicked.connect(self._on_create_song)
        self.btn_add_generic.clicked.connect(self._on_add_generic)
        self.btn_add_offering.clicked.connect(self._on_add_offering)
        self.btn_add_from_theme.clicked.connect(self._on_add_from_theme)
        self.btn_add_section.clicked.connect(self._on_add_empty_section)
        self.btn_add_pptx.clicked.connect(self._on_add_pptx)
        self.btn_edit_fields.clicked.connect(self._on_edit_fields)
        self.btn_delete.clicked.connect(self._on_delete)
        self.btn_edit.clicked.connect(self._on_edit)

        # Menu actions
        self.action_new.triggered.connect(self._on_new)
        self.action_open.triggered.connect(self._on_open)
        self.action_save.triggered.connect(self._on_save)
        self.action_save_as.triggered.connect(self._on_save_as)
        self.action_export.triggered.connect(self._on_export)
        self.action_exit.triggered.connect(self.close)

        self.action_delete.triggered.connect(self._on_delete)
        self.action_move_up.triggered.connect(self._on_move_up)
        self.action_move_down.triggered.connect(self._on_move_down)

        self.action_check_links.triggered.connect(self._on_check_links)
        self.action_edit_fields.triggered.connect(self._on_edit_fields)
        self.action_settings.triggered.connect(self._on_settings)

        self.action_open_theme.triggered.connect(self._on_open_theme)
        self.action_save_as_theme.triggered.connect(self._on_save_as_theme)

        self.action_shortcuts.triggered.connect(self._on_shortcuts)
        self.action_workflow.triggered.connect(self._on_workflow)
        self.action_about.triggered.connect(self._on_about)

        # Language combo
        self.language_combo.currentIndexChanged.connect(self._on_language_combo_changed)

        # Liturgy tree
        self.liturgy_tree.order_changed.connect(self._on_order_changed)
        self.liturgy_tree.item_double_clicked.connect(lambda: self._on_edit(0))
        self.liturgy_tree.section_selected.connect(self._on_section_selected)
        self.liturgy_tree.slide_selected.connect(self._on_slide_selected)

        # Context menu
        self.liturgy_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.liturgy_tree.customContextMenuRequested.connect(self._on_context_menu)

        # Service info panel
        self.service_date_edit.dateChanged.connect(self._on_service_date_changed)
        self.dienstleider_edit.textChanged.connect(self._on_dienstleider_changed)

    def _update_translations(self) -> None:
        """Update all UI text with current translations."""
        self.setWindowTitle(tr("app.title"))

        # Menu
        self.file_menu.setTitle(tr("menu.file"))
        self.action_new.setText(tr("menu.file.new"))
        self.action_open.setText(tr("menu.file.open"))
        self.action_save.setText(tr("menu.file.save"))
        self.action_save_as.setText(tr("menu.file.save_as"))
        self.action_export.setText(tr("menu.file.export"))
        self.action_exit.setText(tr("menu.file.exit"))
        self.action_open_theme.setText(tr("menu.file.open_theme"))
        self.action_save_as_theme.setText(tr("menu.file.save_as_theme"))

        self.edit_menu.setTitle(tr("menu.edit"))
        self.action_delete.setText(tr("menu.edit.delete"))
        self.action_move_up.setText(tr("menu.edit.move_up"))
        self.action_move_down.setText(tr("menu.edit.move_down"))

        self.tools_menu.setTitle(tr("menu.tools"))
        self.action_check_links.setText(tr("menu.tools.check_links"))
        self.action_edit_fields.setText(tr("menu.tools.edit_fields"))
        self.action_settings.setText(tr("menu.tools.settings"))

        self.help_menu.setTitle(tr("menu.help"))
        self.action_shortcuts.setText(tr("menu.help.shortcuts"))
        self.action_workflow.setText(tr("menu.help.workflow"))
        self.action_about.setText(tr("menu.help.about"))

        # Panels
        self.add_items_group.setTitle(tr("panel.actions"))
        self.service_info_group.setTitle(tr("panel.service_info.title"))
        self.liturgy_group.setTitle(tr("panel.liturgy"))

        # Service info labels
        self.service_date_label.setText(tr("panel.service_info.date"))
        self.dienstleider_label.setText(tr("panel.service_info.leader"))
        self.dienstleider_edit.setPlaceholderText(tr("panel.service_info.leader_placeholder"))

        # Buttons
        self.btn_add_song.setText(tr("button.add_song"))
        self.btn_create_song.setText(tr("button.create_song"))
        self.btn_add_generic.setText(tr("button.add_generic"))
        self.btn_add_offering.setText(tr("button.add_offering"))
        self.btn_add_from_theme.setText(tr("button.add_from_theme"))
        self.btn_add_section.setText(tr("button.add_section"))
        self.btn_add_pptx.setText(tr("button.add_pptx"))
        self.btn_edit_fields.setText(tr("button.edit_fields"))
        self.btn_delete.setText(tr("button.delete"))
        self.btn_edit.setText(tr("button.edit"))

        # Status
        self.status_label.setText(tr("status.ready"))

    def _on_language_changed(self, lang: str) -> None:
        """Handle language change."""
        self._update_translations()
        self.liturgy_tree.refresh()
        self._check_warnings()

    def _check_warnings(self) -> None:
        """Check for configuration warnings and display them."""
        warnings = []

        # Check if offering file exists
        collecte_path = self.settings.get_collecte_path(self.base_path)
        if not collecte_path or not os.path.exists(collecte_path):
            warnings.append(f"⚠ {tr('warning.no_offering_file')}")

        # Display warnings
        if warnings:
            self.warning_label.setText("\n".join(warnings))
            self.warning_label.show()
        else:
            self.warning_label.hide()

    def _on_language_combo_changed(self, index: int) -> None:
        """Handle language combo box change."""
        lang = self.language_combo.itemData(index)
        set_language(lang)
        self.settings.language = lang
        self.settings.save()

    def _on_add_song(self) -> None:
        """Add a song to the liturgy."""
        songs = self.folder_scanner.scan_songs()
        dialog = SongPickerDialog(songs, pptx_service=self.pptx_service, parent=self)

        if dialog.exec():
            item = dialog.get_selected_item()
            if item:
                self.liturgy.add_item(item)
                self.liturgy_tree.set_liturgy(self.liturgy)
                self.unsaved_changes = True
                # YouTube search is now manual - user can use Edit button or double-click

    def _on_create_song(self) -> None:
        """Create a new song from lyrics and add to liturgy."""
        dialog = NewSongDialog(self.settings, self.base_path, self)

        if dialog.exec():
            pptx_path = dialog.get_created_pptx()
            folder_path = dialog.get_created_folder()

            if pptx_path and folder_path:
                # Create a SongLiturgyItem for the new song
                from ..models import SongLiturgyItem
                import os

                # Get song title from folder name
                title = os.path.basename(folder_path)

                # Get relative path for source_path
                songs_path = self.settings.get_songs_path(self.base_path)
                try:
                    relative_path = os.path.relpath(folder_path, songs_path)
                except ValueError:
                    relative_path = folder_path

                item = SongLiturgyItem(
                    title=title,
                    source_path=relative_path,
                    pptx_path=pptx_path,
                    is_stub=False,
                )

                self.liturgy.add_item(item)
                self.liturgy_tree.set_liturgy(self.liturgy)
                self.unsaved_changes = True

                # Refresh folder scanner cache
                self.folder_scanner.refresh()

    def _on_add_generic(self) -> None:
        """Add a generic item to the liturgy."""
        items = self.folder_scanner.scan_generic()
        dialog = GenericPickerDialog(items, self.pptx_service, self)

        if dialog.exec():
            item = dialog.get_selected_item()
            if item:
                # For external files with multiple slides, create a section with all slides
                if item.pptx_path and os.path.exists(item.pptx_path) and not item.is_stub:
                    slide_count = self.pptx_service.get_slide_count(item.pptx_path)
                    if slide_count > 1:
                        # Create section with all slides
                        section = LiturgySection(
                            name=item.title,
                            section_type=SectionType.REGULAR,
                        )
                        for i in range(slide_count):
                            slide = LiturgySlide(
                                title=f"{item.title} - Slide {i + 1}" if slide_count > 1 else item.title,
                                slide_index=i,
                                source_path=item.pptx_path,
                                is_stub=False,
                            )
                            section.slides.append(slide)
                        self.liturgy.add_section(section)
                    else:
                        self.liturgy.add_item(item)
                else:
                    self.liturgy.add_item(item)
                self.liturgy_tree.set_liturgy(self.liturgy)
                self.unsaved_changes = True

    def _on_add_offering(self) -> None:
        """Add an offering item to the liturgy."""
        slides = self.folder_scanner.get_offering_slides()
        if not slides:
            QMessageBox.warning(
                self,
                tr("item.offering"),
                tr("error.file_not_found", path=self.settings.collecte_filename)
            )
            return

        dialog = OfferingPickerDialog(slides, self.settings, self.base_path, self.folder_scanner, self)

        if dialog.exec():
            item = dialog.get_selected_item()
            if item:
                self.liturgy.add_item(item)
                self.liturgy_tree.set_liturgy(self.liturgy)
                self.unsaved_changes = True

    def _on_add_from_theme(self) -> None:
        """Add sections/slides from a theme file."""
        dialog = ThemeSectionPicker(self.settings, self.base_path, self.liturgy.sections, self)
        if dialog.exec():
            sections = dialog.get_selected_sections()
            for section in sections:
                self.liturgy.add_section(section)
            self.liturgy_tree.set_liturgy(self.liturgy)
            self.unsaved_changes = True

    def _on_add_empty_section(self) -> None:
        """Add an empty section at the selected position."""
        from PyQt6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(
            self,
            tr("dialog.section.title"),
            tr("dialog.section.enter_name")
        )
        if ok and name.strip():
            section = LiturgySection(name=name.strip(), section_type=SectionType.REGULAR)

            # Get selected section index to insert after it
            selected_idx = self.liturgy_tree.get_selected_section_index()
            if selected_idx >= 0:
                # Insert after the selected section
                self.liturgy.insert_section(selected_idx + 1, section)
                new_idx = selected_idx + 1
            else:
                # No selection, add at the end
                self.liturgy.add_section(section)
                new_idx = len(self.liturgy.sections) - 1

            self.liturgy_tree.set_liturgy(self.liturgy)
            self.liturgy_tree.select_section(new_idx)
            self.unsaved_changes = True

    def _on_add_pptx(self) -> None:
        """Add a PowerPoint file directly as a new section."""
        from PyQt6.QtWidgets import QFileDialog, QInputDialog

        # Open file dialog to select PowerPoint
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("dialog.pptx.browse_title"),
            "",
            "PowerPoint Files (*.pptx *.ppt);;All Files (*)"
        )

        if not file_path:
            return

        # Get default name from filename
        default_name = os.path.splitext(os.path.basename(file_path))[0]

        # Ask for section name
        name, ok = QInputDialog.getText(
            self,
            tr("dialog.pptx.title"),
            tr("dialog.pptx.enter_name"),
            text=default_name
        )

        if not ok or not name.strip():
            return

        # Get slides info from the PowerPoint
        slides_info = self.pptx_service.get_slides_info(file_path)

        if not slides_info:
            QMessageBox.warning(
                self,
                tr("dialog.fields.error"),
                tr("dialog.pptx.no_slides")
            )
            return

        # Create section with slides
        section = LiturgySection(
            name=name.strip(),
            section_type=SectionType.REGULAR
        )

        for slide_info in slides_info:
            slide = LiturgySlide(
                title=slide_info.get("title", f"Slide {slide_info['index'] + 1}"),
                slide_index=slide_info["index"],
                source_path=file_path,
            )
            section.slides.append(slide)

        # Insert at selected position or at end
        selected_idx = self.liturgy_tree.get_selected_section_index()
        if selected_idx >= 0:
            self.liturgy.insert_section(selected_idx + 1, section)
            new_idx = selected_idx + 1
        else:
            self.liturgy.add_section(section)
            new_idx = len(self.liturgy.sections) - 1

        self.liturgy_tree.set_liturgy(self.liturgy)
        self.liturgy_tree.select_section(new_idx)
        self.unsaved_changes = True

    def _on_delete(self) -> None:
        """Delete selected item or section."""
        index = self.liturgy_tree.get_selected_index()
        if index >= 0:
            # Get the title for confirmation dialog
            if self.liturgy.sections:
                # V2 mode - get section
                if index < len(self.liturgy.sections):
                    title = self.liturgy.sections[index].name
                else:
                    return
            else:
                # V1 mode - get item
                if index < len(self.liturgy._items):
                    title = self.liturgy._items[index].title
                else:
                    return

            reply = QMessageBox.question(
                self,
                tr("dialog.confirm.delete"),
                tr("dialog.confirm.delete_text", title=title),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                if self.liturgy.sections:
                    self.liturgy.remove_section(index)
                else:
                    self.liturgy.remove_item(index)
                self.liturgy_tree.set_liturgy(self.liturgy)
                self.unsaved_changes = True

    def _on_edit(self, initial_tab: int = 0) -> None:
        """Edit selected item using unified editor dialog."""
        self._open_section_editor(initial_tab)

    def _open_section_editor(self, initial_tab: int = 0) -> None:
        """Open the unified section editor dialog."""
        # Check if a slide is selected (v2 mode)
        slide_info = self.liturgy_tree.get_selected_slide_info()
        section = None
        slide = None

        if slide_info:
            section_id, slide_id = slide_info
            result = self.liturgy.get_slide_by_id(slide_id)
            if result:
                section, slide = result
        else:
            # Section selected
            index = self.liturgy_tree.get_selected_index()
            if index >= 0 and self.liturgy.sections:
                section = self.liturgy.sections[index]
                # Use first slide if available
                if section.slides:
                    slide = section.slides[0]

        if not section:
            return

        # Open unified editor
        dialog = SectionEditorDialog(
            section=section,
            slide=slide,
            pptx_service=self.pptx_service,
            youtube_service=self.youtube_service,
            initial_tab=initial_tab,
            parent=self
        )

        if dialog.exec():
            self.unsaved_changes = True
            self.liturgy_tree.set_liturgy(self.liturgy)
            # Save YouTube links to file if applicable
            if section.is_song:
                self._save_youtube_links(section)

    def _save_youtube_links(self, section: LiturgySection) -> None:
        """Save YouTube links to youtube.txt file."""
        if section.song_source_path:
            song_folder = os.path.join(
                self.settings.get_songs_path(self.base_path),
                section.song_source_path
            )
            urls = section.youtube_links or []
            self.youtube_service.write_youtube_file(song_folder, urls)

    def _on_context_menu(self, position) -> None:
        """Show context menu for liturgy tree."""
        # Get selected item info
        slide_info = self.liturgy_tree.get_selected_slide_info()
        section = None

        if slide_info:
            section_id, slide_id = slide_info
            result = self.liturgy.get_slide_by_id(slide_id)
            if result:
                section, _ = result
        else:
            index = self.liturgy_tree.get_selected_index()
            if index >= 0 and self.liturgy.sections:
                section = self.liturgy.sections[index]

        if not section:
            return

        # Create context menu
        menu = QMenu(self)

        # Edit Fields action
        edit_fields_action = menu.addAction(tr("context.edit_fields"))
        edit_fields_action.triggered.connect(lambda: self._open_section_editor(SectionEditorDialog.TAB_FIELDS))

        # Edit YouTube action (only for songs)
        if section.is_song:
            edit_youtube_action = menu.addAction(tr("context.edit_youtube"))
            edit_youtube_action.triggered.connect(lambda: self._open_section_editor(SectionEditorDialog.TAB_YOUTUBE))

        menu.addSeparator()

        # Delete action
        delete_action = menu.addAction(tr("menu.edit.delete"))
        delete_action.triggered.connect(self._on_delete)

        # Show menu at cursor position
        menu.exec(self.liturgy_tree.mapToGlobal(position))

    def _on_move_up(self) -> None:
        """Move selected item up."""
        index = self.liturgy_tree.get_selected_index()
        if index > 0:
            self.liturgy.move_item(index, index - 1)
            self.liturgy_tree.set_liturgy(self.liturgy)
            self.liturgy_tree.select_index(index - 1)
            self.unsaved_changes = True

    def _on_move_down(self) -> None:
        """Move selected item down."""
        index = self.liturgy_tree.get_selected_index()
        if 0 <= index < len(self.liturgy.items) - 1:
            self.liturgy.move_item(index, index + 1)
            self.liturgy_tree.set_liturgy(self.liturgy)
            self.liturgy_tree.select_index(index + 1)
            self.unsaved_changes = True

    def _on_order_changed(self) -> None:
        """Handle drag-and-drop reorder."""
        # The tree widget updates the liturgy directly
        self.unsaved_changes = True

    def _on_section_selected(self, section_id: str) -> None:
        """Handle section selection."""
        # Could be used to update preview/details panel
        pass

    def _on_slide_selected(self, section_id: str, slide_id: str) -> None:
        """Handle slide selection."""
        # Could be used to update preview/details panel
        pass

    def _on_service_date_changed(self, qdate: QDate) -> None:
        """Handle service date change."""
        self.liturgy.service_date = qdate.toString("yyyy-MM-dd")
        self.unsaved_changes = True

    def _on_dienstleider_changed(self, text: str) -> None:
        """Handle dienstleider change."""
        self.liturgy.dienstleider = text.strip() if text.strip() else None
        self.unsaved_changes = True

    def _sync_service_info_from_liturgy(self) -> None:
        """Sync service info panel from liturgy values."""
        # Block signals to avoid triggering unsaved_changes
        self.service_date_edit.blockSignals(True)
        self.dienstleider_edit.blockSignals(True)

        if self.liturgy.service_date:
            qdate = QDate.fromString(self.liturgy.service_date, "yyyy-MM-dd")
            if qdate.isValid():
                self.service_date_edit.setDate(qdate)
        else:
            self._set_next_sunday()

        self.dienstleider_edit.setText(self.liturgy.dienstleider or "")

        self.service_date_edit.blockSignals(False)
        self.dienstleider_edit.blockSignals(False)

    def _on_new(self) -> None:
        """Create new liturgy."""
        if self.unsaved_changes:
            if not self._confirm_discard():
                return

        self.liturgy = Liturgy(name=self._generate_default_name())
        self.current_file = None
        self.unsaved_changes = False
        self.liturgy_tree.set_liturgy(self.liturgy)
        self._sync_service_info_from_liturgy()

    def _on_open(self) -> None:
        """Open existing liturgy."""
        if self.unsaved_changes:
            if not self._confirm_discard():
                return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("menu.file.open"),
            self.settings.get_output_path(self.base_path),
            "Liturgy files (*.json);;All files (*.*)"
        )

        if file_path:
            try:
                # Load with migration support, resolving relative paths
                effective_base = self.settings.get_effective_base_path(self.base_path)
                self.liturgy, was_migrated = Liturgy.load_with_migration(file_path, effective_base)
                self.current_file = file_path
                self.unsaved_changes = was_migrated  # Mark as changed if migrated
                self.liturgy_tree.set_liturgy(self.liturgy)
                self._sync_service_info_from_liturgy()

                # Show migration warning if format was upgraded
                if was_migrated:
                    QMessageBox.information(
                        self,
                        tr("dialog.migration.title"),
                        tr("dialog.migration.text")
                    )

                # Check links on load
                self._check_links_background()

            except Exception as e:
                QMessageBox.critical(
                    self,
                    tr("menu.file.open"),
                    tr("error.load_failed", error=str(e))
                )

    def _on_save(self) -> None:
        """Save current liturgy."""
        if self.current_file:
            self._save_to_file(self.current_file)
        else:
            self._on_save_as()

    def _on_save_as(self) -> None:
        """Save liturgy to new file."""
        default_name = f"{self.liturgy.name}.json"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            tr("menu.file.save_as"),
            os.path.join(self.settings.get_output_path(self.base_path), default_name),
            "Liturgy files (*.json);;All files (*.*)"
        )

        if file_path:
            self._save_to_file(file_path)

    def _save_to_file(self, file_path: str) -> None:
        """Save liturgy to specified file."""
        try:
            # Save with relative paths based on effective base path
            effective_base = self.settings.get_effective_base_path(self.base_path)
            self.liturgy.save(file_path, effective_base)
            self.current_file = file_path
            self.unsaved_changes = False
            self.status_label.setText(tr("status.ready"))
        except Exception as e:
            QMessageBox.critical(
                self,
                tr("menu.file.save"),
                tr("error.save_failed", error=str(e))
            )

    def _on_export(self) -> None:
        """Export liturgy to output files."""
        # Check for unfilled fields
        unfilled_slides = self.liturgy_tree.get_slides_with_unfilled_fields()
        if unfilled_slides:
            # Build warning message
            warning_lines = [tr("dialog.export.unfilled_warning")]
            for section, slide, fields in unfilled_slides[:10]:  # Limit to 10
                warning_lines.append(f"• {section.name} / {slide.title}: {', '.join(fields)}")
            if len(unfilled_slides) > 10:
                warning_lines.append(f"... and {len(unfilled_slides) - 10} more")

            reply = QMessageBox.warning(
                self,
                tr("dialog.export.title"),
                "\n".join(warning_lines),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        dialog = ExportDialog(self.liturgy, self.export_service, self.settings, self.base_path, self)
        dialog.exec()

    def _on_check_links(self) -> None:
        """Check all YouTube links (thorough mode)."""
        self.status_label.setText(tr("status.checking_links"))
        self.statusbar.repaint()

        # Collect all YouTube URLs
        urls = []
        for item in self.liturgy.items:
            if isinstance(item, SongLiturgyItem):
                urls.extend(item.youtube_links)

        if not urls:
            self.status_label.setText(tr("status.links_valid"))
            return

        # Validate (thorough)
        results = self.youtube_service.validate_links_batch(urls, thorough=True)
        invalid_count = sum(1 for _, valid, _ in results if not valid)

        if invalid_count > 0:
            self.status_label.setText(tr("status.links_invalid", count=invalid_count))
            # Could show detailed dialog here
        else:
            self.status_label.setText(tr("status.links_valid"))

    def _check_links_background(self) -> None:
        """Check links in background (fast mode)."""
        # Simplified fast check
        urls = []
        for item in self.liturgy.items:
            if isinstance(item, SongLiturgyItem):
                urls.extend(item.youtube_links)

        if urls:
            results = self.youtube_service.validate_links_batch(urls, thorough=False)
            invalid_count = sum(1 for _, valid, _ in results if not valid)
            if invalid_count > 0:
                self.status_label.setText(tr("status.links_invalid", count=invalid_count))

    def _on_settings(self) -> None:
        """Open settings dialog."""
        dialog = SettingsDialog(self.settings, self.base_path, self)
        if dialog.exec():
            self.settings = dialog.get_settings()
            self.settings.save()
            # Reinitialize services with new settings and refresh folder scanner
            self.folder_scanner = FolderScanner(self.settings, self.base_path)
            self.folder_scanner.refresh()  # Force rescan with new settings
            self.export_service = ExportService(self.settings, self.base_path)
            self.theme_service = ThemeService(self.settings, self.base_path)
            # Refresh dienstleider autocomplete
            self._setup_dienstleider_autocomplete()
            # Recheck warnings (offering file may have changed)
            self._check_warnings()

    def _on_open_theme(self) -> None:
        """Open a theme template as liturgy."""
        if self.unsaved_changes:
            if not self._confirm_discard():
                return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("menu.file.open_theme"),
            self.settings.get_themes_path(self.base_path),
            "PowerPoint files (*.pptx *.ppt);;All files (*.*)"
        )

        if file_path:
            try:
                self.liturgy = self.theme_service.load_as_liturgy(file_path)
                self.current_file = None  # Theme is not saved as JSON yet
                self.unsaved_changes = True
                self.liturgy_tree.set_liturgy(self.liturgy)
                self._sync_service_info_from_liturgy()
                self.status_label.setText(tr("status.theme_loaded"))
            except Exception as e:
                QMessageBox.critical(
                    self,
                    tr("menu.file.open_theme"),
                    tr("error.load_failed", error=str(e))
                )

    def _on_save_as_theme(self) -> None:
        """Save liturgy as theme template."""
        default_name = f"{self.liturgy.name}.pptx"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            tr("menu.file.save_as_theme"),
            os.path.join(self.settings.get_themes_path(self.base_path), default_name),
            "PowerPoint files (*.pptx);;All files (*.*)"
        )

        if file_path:
            try:
                self.theme_service.save_as_theme(self.liturgy, file_path)
                self.status_label.setText(tr("status.theme_saved"))
                QMessageBox.information(
                    self,
                    tr("menu.file.save_as_theme"),
                    tr("dialog.theme.saved", path=file_path)
                )
            except Exception as e:
                QMessageBox.critical(
                    self,
                    tr("menu.file.save_as_theme"),
                    tr("error.save_failed", error=str(e))
                )

    def _on_edit_fields(self) -> None:
        """Open field editor dialog for all fields in liturgy."""
        if not self.liturgy.sections:
            QMessageBox.information(
                self,
                tr("menu.tools.edit_fields"),
                tr("dialog.fields.no_sections")
            )
            return

        dialog = BulkFieldEditor(self.liturgy, self.pptx_service, self)
        if dialog.exec():
            self.liturgy = dialog.get_updated_liturgy()
            self.liturgy_tree.set_liturgy(self.liturgy)
            self.unsaved_changes = True

    def _on_shortcuts(self) -> None:
        """Show keyboard shortcuts dialog."""
        from PyQt6.QtWidgets import QTextBrowser
        dialog = QDialog(self)
        dialog.setWindowTitle(tr("dialog.shortcuts.title"))
        dialog.setMinimumSize(500, 400)
        layout = QVBoxLayout(dialog)

        text_browser = QTextBrowser()
        text_browser.setOpenExternalLinks(False)
        text_browser.setHtml(tr("dialog.shortcuts.content"))
        layout.addWidget(text_browser)

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)

        dialog.exec()

    def _on_workflow(self) -> None:
        """Show weekly workflow guide dialog."""
        from PyQt6.QtWidgets import QTextBrowser
        dialog = QDialog(self)
        dialog.setWindowTitle(tr("dialog.workflow.title"))
        dialog.setMinimumSize(600, 500)
        layout = QVBoxLayout(dialog)

        text_browser = QTextBrowser()
        text_browser.setOpenExternalLinks(False)
        text_browser.setHtml(tr("dialog.workflow.content"))
        layout.addWidget(text_browser)

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)

        dialog.exec()

    def _on_about(self) -> None:
        """Show about dialog."""
        dialog = AboutDialog(self)
        dialog.exec()

    def _prompt_youtube_search(self, item: SongLiturgyItem) -> None:
        """Prompt user to search for YouTube video."""
        dialog = YouTubeDialog(item.title, self.youtube_service, self)
        if dialog.exec():
            urls = dialog.get_selected_urls()
            if urls:
                item.youtube_links = urls
                # Save to youtube.txt
                if item.source_path:
                    song_folder = os.path.join(
                        self.settings.get_songs_path(self.base_path),
                        item.source_path
                    )
                    self.youtube_service.write_youtube_file(song_folder, urls)
                self.unsaved_changes = True

    def _prompt_youtube_search_v2(self, section: LiturgySection) -> None:
        """Prompt user to search for YouTube video (v2 sections)."""
        dialog = YouTubeDialog(section.name, self.youtube_service, self)
        if dialog.exec():
            urls = dialog.get_selected_urls()
            if urls:
                section.youtube_links = urls
                # Save to youtube.txt if we have a song source path
                if section.song_source_path:
                    song_folder = os.path.join(
                        self.settings.get_songs_path(self.base_path),
                        section.song_source_path
                    )
                    self.youtube_service.write_youtube_file(song_folder, urls)
                self.liturgy_tree.set_liturgy(self.liturgy)
                self.unsaved_changes = True

    def _confirm_discard(self) -> bool:
        """Ask user to confirm discarding unsaved changes."""
        reply = QMessageBox.question(
            self,
            tr("dialog.confirm.unsaved"),
            tr("dialog.confirm.unsaved_text"),
            QMessageBox.StandardButton.Save |
            QMessageBox.StandardButton.Discard |
            QMessageBox.StandardButton.Cancel
        )

        if reply == QMessageBox.StandardButton.Save:
            self._on_save()
            return not self.unsaved_changes
        elif reply == QMessageBox.StandardButton.Discard:
            return True
        return False

    def closeEvent(self, event) -> None:
        """Handle window close."""
        if self.unsaved_changes:
            if not self._confirm_discard():
                event.ignore()
                return

        # Save window size
        self.settings.window_width = self.width()
        self.settings.window_height = self.height()
        self.settings.save()

        event.accept()
