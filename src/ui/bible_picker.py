"""Bible text picker dialog.

Allows the user to:
  1. Enter a Bible reference (e.g. "John 3:16-21").
  2. Select up to 6 translations.
  3. Configure font name and size.
  4. Generate a PPTX with multi-language Bible slides.
"""

from __future__ import annotations

import os
import tempfile
from typing import Dict, List, Optional, Tuple

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from ..i18n import tr
from ..logging_config import get_logger
from ..services.bible_service import (
    BibleService,
    BibleTranslation,
    BUILTIN_TRANSLATIONS,
    parse_reference,
)
from ..services.bible_slide_service import BibleSlideConfig, BibleSlideService

logger = get_logger("bible_picker")


# ---------------------------------------------------------------------------
# Background workers
# ---------------------------------------------------------------------------

class _FetchTranslationsWorker(QThread):
    """Fetches additional translations for a language from YouVersion."""

    finished = pyqtSignal(list)    # List[BibleTranslation]
    error = pyqtSignal(str)

    def __init__(self, language_tag: str, bible_service: BibleService, parent=None):
        super().__init__(parent)
        self._lang = language_tag
        self._svc = bible_service

    def run(self) -> None:
        try:
            results = self._svc.fetch_translations_for_language(self._lang)
            self.finished.emit(results)
        except Exception as exc:
            self.error.emit(str(exc))


class _GenerateSlidesWorker(QThread):
    """Generates Bible slides in a background thread."""

    finished = pyqtSignal(str)   # path to generated PPTX
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(
        self,
        reference: str,
        translation_ids: List[int],
        config: BibleSlideConfig,
        parent=None,
    ):
        super().__init__(parent)
        self._reference = reference
        self._translation_ids = translation_ids
        self._config = config

    def run(self) -> None:
        try:
            self.progress.emit(tr("dialog.bible.status.fetching"))
            svc = BibleSlideService()
            path = svc.create_slides(
                self._reference, self._translation_ids, self._config
            )
            self.finished.emit(path)
        except Exception as exc:
            logger.error("Bible slide generation failed: %s", exc, exc_info=True)
            self.error.emit(str(exc))


# ---------------------------------------------------------------------------
# Dialog
# ---------------------------------------------------------------------------

class BiblePickerDialog(QDialog):
    """Dialog for creating multi-language Bible text slides.

    On acceptance, ``result_pptx_path`` contains the path to a temporary
    PPTX file that the caller is responsible for copying / deleting.
    ``result_section_name`` contains a human-readable name for the liturgy
    section.
    """

    def __init__(
        self,
        default_font_name: str = "Calibri",
        default_font_size: int = 12,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._bible_service = BibleService()
        self._slide_service = BibleSlideService(self._bible_service)
        self._worker: Optional[_GenerateSlidesWorker] = None
        self._fetch_worker: Optional[_FetchTranslationsWorker] = None

        self.result_pptx_path: Optional[str] = None
        self.result_section_name: str = ""

        self._default_font_name = default_font_name
        self._default_font_size = default_font_size

        # All available translations (built-in + any fetched)
        self._all_translations: List[BibleTranslation] = [
            BibleTranslation(**{k: v for k, v in t.items()})
            for t in BUILTIN_TRANSLATIONS
        ]

        self._setup_ui()
        self._connect_signals()
        self._populate_available_list()

    # ------------------------------------------------------------------
    # UI setup
    # ------------------------------------------------------------------

    def _setup_ui(self) -> None:
        self.setWindowTitle(tr("dialog.bible.title"))
        self.setMinimumSize(720, 580)
        self.resize(800, 640)

        outer = QVBoxLayout(self)

        # ---- Reference input ----
        ref_group = QGroupBox(tr("dialog.bible.reference_group"))
        ref_form = QFormLayout(ref_group)

        self.reference_edit = QLineEdit()
        self.reference_edit.setPlaceholderText(tr("dialog.bible.reference_placeholder"))
        ref_form.addRow(tr("dialog.bible.reference_label"), self.reference_edit)

        self.reference_status = QLabel("")
        self.reference_status.setStyleSheet("color: grey; font-size: 11px;")
        ref_form.addRow("", self.reference_status)

        outer.addWidget(ref_group)

        # ---- Translations ----
        trans_group = QGroupBox(tr("dialog.bible.translations_group"))
        trans_layout = QHBoxLayout(trans_group)

        # Left: available translations
        avail_panel = QVBoxLayout()
        avail_title = QLabel(tr("dialog.bible.available_translations"))
        avail_title.setStyleSheet("font-weight: bold;")
        avail_panel.addWidget(avail_title)

        # Language filter
        lang_row = QHBoxLayout()
        lang_row.addWidget(QLabel(tr("dialog.bible.filter_language")))
        self.language_combo = QComboBox()
        self.language_combo.addItem(tr("dialog.bible.all_languages"), "")
        lang_row.addWidget(self.language_combo)

        fetch_btn = QPushButton(tr("dialog.bible.fetch_more"))
        fetch_btn.setToolTip(tr("dialog.bible.fetch_more_tooltip"))
        fetch_btn.setFixedWidth(110)
        self._fetch_btn = fetch_btn
        lang_row.addWidget(fetch_btn)
        avail_panel.addLayout(lang_row)

        self.available_list = QListWidget()
        self.available_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        avail_panel.addWidget(self.available_list)
        avail_panel.addWidget(QLabel(tr("dialog.bible.available_hint")))

        # Middle: add/remove buttons
        mid_panel = QVBoxLayout()
        mid_panel.addStretch()
        add_btn = QPushButton("→")
        add_btn.setFixedWidth(40)
        add_btn.setToolTip(tr("dialog.bible.add_translation"))
        self._add_btn = add_btn
        remove_btn = QPushButton("←")
        remove_btn.setFixedWidth(40)
        remove_btn.setToolTip(tr("dialog.bible.remove_translation"))
        self._remove_btn = remove_btn
        mid_panel.addWidget(add_btn)
        mid_panel.addWidget(remove_btn)
        mid_panel.addStretch()

        # Right: selected translations
        sel_panel = QVBoxLayout()
        sel_title = QLabel(tr("dialog.bible.selected_translations"))
        sel_title.setStyleSheet("font-weight: bold;")
        sel_panel.addWidget(sel_title)
        self.selected_list = QListWidget()
        self.selected_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.selected_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        sel_panel.addWidget(self.selected_list)
        sel_panel.addWidget(QLabel(tr("dialog.bible.selected_hint")))

        trans_layout.addLayout(avail_panel, 3)
        trans_layout.addLayout(mid_panel, 0)
        trans_layout.addLayout(sel_panel, 2)

        outer.addWidget(trans_group)

        # ---- Font settings ----
        font_group = QGroupBox(tr("dialog.bible.font_group"))
        font_form = QFormLayout(font_group)

        self.font_name_edit = QLineEdit(self._default_font_name)
        font_form.addRow(tr("dialog.bible.font_name"), self.font_name_edit)

        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(6, 36)
        self.font_size_spin.setValue(self._default_font_size)
        self.font_size_spin.setSuffix(" pt")
        font_form.addRow(tr("dialog.bible.font_size"), self.font_size_spin)

        outer.addWidget(font_group)

        # ---- Progress / status ----
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)   # indeterminate
        self.progress_bar.setVisible(False)
        outer.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(self.status_label)

        # ---- Buttons ----
        button_box = QDialogButtonBox()
        self.ok_btn = button_box.addButton(
            tr("dialog.bible.generate"), QDialogButtonBox.ButtonRole.AcceptRole
        )
        self.ok_btn.setDefault(True)
        button_box.addButton(QDialogButtonBox.StandardButton.Cancel)
        button_box.rejected.connect(self.reject)
        button_box.accepted.connect(self._on_generate)
        outer.addWidget(button_box)

        self._add_language_items_to_combo()

    def _add_language_items_to_combo(self) -> None:
        """Populate the language filter combo with languages from built-in catalog."""
        seen = set()
        for t in self._all_translations:
            if t.language and t.language not in seen:
                seen.add(t.language)
                self.language_combo.addItem(
                    f"{t.language_name} ({t.language.upper()})", t.language
                )

    # ------------------------------------------------------------------
    # Signals
    # ------------------------------------------------------------------

    def _connect_signals(self) -> None:
        self.language_combo.currentIndexChanged.connect(self._on_language_filter_changed)
        self._add_btn.clicked.connect(self._on_add_translation)
        self._remove_btn.clicked.connect(self._on_remove_translation)
        self._fetch_btn.clicked.connect(self._on_fetch_more)
        self.available_list.itemDoubleClicked.connect(self._on_add_translation)
        self.selected_list.itemDoubleClicked.connect(self._on_remove_translation)
        self.reference_edit.textChanged.connect(self._on_reference_changed)

    # ------------------------------------------------------------------
    # Population
    # ------------------------------------------------------------------

    def _populate_available_list(self, filter_lang: str = "") -> None:
        """Fill the available translations list (optionally filtered by language)."""
        self.available_list.clear()

        # Collect IDs already in selected list
        selected_ids = {
            self.selected_list.item(i).data(Qt.ItemDataRole.UserRole)
            for i in range(self.selected_list.count())
        }

        shown = set()
        for t in self._all_translations:
            if filter_lang and t.language != filter_lang:
                continue
            if t.id in shown:
                continue
            if t.id in selected_ids:
                continue
            shown.add(t.id)
            item = QListWidgetItem(f"{t.abbreviation} – {t.name} ({t.language_name})")
            item.setData(Qt.ItemDataRole.UserRole, t.id)
            item.setData(Qt.ItemDataRole.UserRole + 1, t)
            self.available_list.addItem(item)

    # ------------------------------------------------------------------
    # Slots
    # ------------------------------------------------------------------

    def _on_language_filter_changed(self) -> None:
        lang = self.language_combo.currentData() or ""
        self._populate_available_list(filter_lang=lang)

    def _on_reference_changed(self, text: str) -> None:
        text = text.strip()
        if not text:
            self.reference_status.setText("")
            return
        try:
            ref = parse_reference(text)
            self.reference_status.setText(
                f"✓ {ref.display_str}"
            )
            self.reference_status.setStyleSheet("color: green; font-size: 11px;")
        except ValueError as exc:
            self.reference_status.setText(str(exc))
            self.reference_status.setStyleSheet("color: red; font-size: 11px;")

    def _on_add_translation(self) -> None:
        """Add selected translations to the selected list (max 6)."""
        selected_ids = {
            self.selected_list.item(i).data(Qt.ItemDataRole.UserRole)
            for i in range(self.selected_list.count())
        }
        for item in self.available_list.selectedItems():
            if self.selected_list.count() >= 6:
                QMessageBox.information(
                    self,
                    tr("dialog.bible.title"),
                    tr("dialog.bible.max_translations"),
                )
                break
            t_id = item.data(Qt.ItemDataRole.UserRole)
            if t_id in selected_ids:
                continue
            t: BibleTranslation = item.data(Qt.ItemDataRole.UserRole + 1)
            sel_item = QListWidgetItem(f"{t.abbreviation} – {t.name}")
            sel_item.setData(Qt.ItemDataRole.UserRole, t.id)
            sel_item.setData(Qt.ItemDataRole.UserRole + 1, t)
            self.selected_list.addItem(sel_item)
            selected_ids.add(t_id)

        self._populate_available_list(
            filter_lang=self.language_combo.currentData() or ""
        )

    def _on_remove_translation(self) -> None:
        """Remove selected translations from the selected list."""
        for item in self.selected_list.selectedItems():
            self.selected_list.takeItem(self.selected_list.row(item))
        self._populate_available_list(
            filter_lang=self.language_combo.currentData() or ""
        )

    def _on_fetch_more(self) -> None:
        """Fetch additional translations for the currently selected language."""
        lang = self.language_combo.currentData() or ""
        if not lang:
            QMessageBox.information(
                self,
                tr("dialog.bible.title"),
                tr("dialog.bible.select_language_first"),
            )
            return

        self._fetch_btn.setEnabled(False)
        self.status_label.setText(tr("dialog.bible.status.fetching_translations"))
        self.progress_bar.setVisible(True)

        self._fetch_worker = _FetchTranslationsWorker(lang, self._bible_service, self)
        self._fetch_worker.finished.connect(self._on_fetch_translations_done)
        self._fetch_worker.error.connect(self._on_fetch_translations_error)
        self._fetch_worker.start()

    def _on_fetch_translations_done(self, translations: list) -> None:
        self._fetch_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_label.setText("")

        # Merge into all_translations (deduplicate by id)
        existing_ids = {t.id for t in self._all_translations}
        added = 0
        for t in translations:
            if t.id not in existing_ids:
                self._all_translations.append(t)
                existing_ids.add(t.id)
                added += 1

        # Also add any new languages to combo
        existing_langs = {
            self.language_combo.itemData(i)
            for i in range(self.language_combo.count())
        }
        for t in translations:
            if t.language and t.language not in existing_langs:
                self.language_combo.addItem(
                    f"{t.language_name} ({t.language.upper()})", t.language
                )
                existing_langs.add(t.language)

        self._populate_available_list(
            filter_lang=self.language_combo.currentData() or ""
        )
        msg = tr("dialog.bible.fetch_done", count=added)
        self.status_label.setText(msg)

    def _on_fetch_translations_error(self, error: str) -> None:
        self._fetch_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_label.setText(tr("dialog.bible.fetch_error", error=error))

    # ------------------------------------------------------------------
    # Generate
    # ------------------------------------------------------------------

    def _on_generate(self) -> None:
        """Validate inputs and start slide generation."""
        reference = self.reference_edit.text().strip()
        if not reference:
            QMessageBox.warning(
                self,
                tr("dialog.bible.title"),
                tr("dialog.bible.error.no_reference"),
            )
            return

        try:
            parse_reference(reference)
        except ValueError as exc:
            QMessageBox.warning(self, tr("dialog.bible.title"), str(exc))
            return

        if self.selected_list.count() == 0:
            QMessageBox.warning(
                self,
                tr("dialog.bible.title"),
                tr("dialog.bible.error.no_translations"),
            )
            return

        translation_ids = [
            self.selected_list.item(i).data(Qt.ItemDataRole.UserRole)
            for i in range(self.selected_list.count())
        ]

        config = BibleSlideConfig(
            font_name=self.font_name_edit.text().strip() or "Calibri",
            font_size=self.font_size_spin.value(),
        )

        self._set_generating(True)

        self._worker = _GenerateSlidesWorker(reference, translation_ids, config, self)
        self._worker.finished.connect(self._on_generation_done)
        self._worker.error.connect(self._on_generation_error)
        self._worker.progress.connect(self._on_generation_progress)
        self._worker.start()

    def _on_generation_done(self, path: str) -> None:
        self._set_generating(False)
        self.result_pptx_path = path
        # Build a nice section name
        ref_text = self.reference_edit.text().strip()
        abbrevs = " / ".join(
            self.selected_list.item(i).data(Qt.ItemDataRole.UserRole + 1).abbreviation
            for i in range(self.selected_list.count())
        )
        self.result_section_name = f"{ref_text} ({abbrevs})"
        self.accept()

    def _on_generation_error(self, error: str) -> None:
        self._set_generating(False)
        QMessageBox.critical(
            self,
            tr("dialog.bible.title"),
            tr("dialog.bible.error.generation", error=error),
        )

    def _on_generation_progress(self, msg: str) -> None:
        self.status_label.setText(msg)

    def _set_generating(self, generating: bool) -> None:
        self.ok_btn.setEnabled(not generating)
        self.progress_bar.setVisible(generating)
        if generating:
            self.status_label.setText(tr("dialog.bible.status.generating"))
        else:
            self.status_label.setText("")
