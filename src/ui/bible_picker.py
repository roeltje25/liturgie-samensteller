"""Bible text picker dialog.

Allows the user to:
  1. Enter a Bible reference (e.g. "John 3:16-21").
  2. Select up to 6 translations, each optionally with its own reference
     override (to handle verse-numbering discrepancies between traditions).
  3. Preview fetched text for all translations side-by-side.
  4. Toggle between the original YouVersion text and a Google Translate
     rendition (translated to a configurable target language).
  5. Configure font name and size.
  6. Generate a PPTX with multi-language Bible slides.
"""

from __future__ import annotations

import os
import shutil
from typing import Dict, List, Optional, Tuple

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from ..i18n import tr
from ..logging_config import get_logger
from ..services.bible_service import (
    BibleService,
    BibleTranslation,
    BibleVerse,
    BUILTIN_TRANSLATIONS,
    parse_reference,
)
from ..services.bible_slide_service import BibleSlideConfig, BibleSlideService
from ..services.google_translate_service import (
    TRANSLATE_LANGUAGES,
    is_rtl,
    translate_batch,
)

logger = get_logger("bible_picker")


# ---------------------------------------------------------------------------
# Background workers
# ---------------------------------------------------------------------------

class _FetchTranslationsWorker(QThread):
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, language_tag: str, bible_service: BibleService, parent=None):
        super().__init__(parent)
        self._lang = language_tag
        self._svc = bible_service

    def run(self) -> None:
        try:
            self.finished.emit(self._svc.fetch_translations_for_language(self._lang))
        except Exception as exc:
            self.error.emit(str(exc))


class _FetchPreviewWorker(QThread):
    """Fetches verse texts for all selected translations."""
    # Emits (col_idx, translation_name, language, verses_as_plain_text_list)
    column_ready = pyqtSignal(int, str, str, list)
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(
        self,
        slots: List[Tuple[int, BibleTranslation, str]],  # (col_idx, trans, reference_str)
        bible_service: BibleService,
        parent=None,
    ):
        super().__init__(parent)
        self._slots = slots
        self._svc = bible_service

    def run(self) -> None:
        try:
            for col_idx, trans, ref_str in self._slots:
                try:
                    ref = parse_reference(ref_str)
                    verses = self._svc.get_verses(ref, trans.id)
                    texts = [
                        f"{v.verse_num}\u00a0{v.text}" if v.verse_num else v.text
                        for v in verses
                    ]
                except Exception as exc:
                    texts = [f"(Error: {exc})"]
                self.column_ready.emit(col_idx, trans.name, trans.language, texts)
            self.finished.emit()
        except Exception as exc:
            self.error.emit(str(exc))


class _TranslateWorker(QThread):
    """Translates verse texts via Google Translate."""
    column_ready = pyqtSignal(int, list)   # (col_idx, translated_texts)
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(
        self,
        columns: List[Tuple[int, List[str]]],  # (col_idx, texts)
        target_lang: str,
        parent=None,
    ):
        super().__init__(parent)
        self._columns = columns
        self._target_lang = target_lang

    def run(self) -> None:
        try:
            for col_idx, texts in self._columns:
                try:
                    translated = translate_batch(texts, self._target_lang)
                except Exception as exc:
                    logger.warning("Translation failed for col %s: %s", col_idx, exc)
                    translated = texts
                self.column_ready.emit(col_idx, translated)
            self.finished.emit()
        except Exception as exc:
            self.error.emit(str(exc))


class _GenerateSlidesWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(
        self,
        reference: str,
        translation_ids: List[int],
        config: BibleSlideConfig,
        reference_overrides: Dict[int, str],
        api_key: str = "",
        parent=None,
    ):
        super().__init__(parent)
        self._reference = reference
        self._translation_ids = translation_ids
        self._config = config
        self._overrides = reference_overrides
        self._api_key = api_key

    def run(self) -> None:
        try:
            self.progress.emit(tr("dialog.bible.status.fetching"))
            svc = BibleSlideService(BibleService(api_key=self._api_key))
            path = svc.create_slides(
                self._reference,
                self._translation_ids,
                self._config,
                self._overrides,
            )
            self.finished.emit(path)
        except Exception as exc:
            logger.error("Bible slide generation failed: %s", exc, exc_info=True)
            self.error.emit(str(exc))


# ---------------------------------------------------------------------------
# Main dialog
# ---------------------------------------------------------------------------

class BiblePickerDialog(QDialog):
    """Dialog for creating multi-language Bible text slides.

    On acceptance:
      - ``result_pptx_path`` → path to a temporary PPTX (caller must move/delete)
      - ``result_section_name`` → human-readable liturgy section name
    """

    def __init__(
        self,
        default_font_name: str = "Calibri",
        default_font_size: int = 12,
        api_key: str = "",
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._bible_service = BibleService(api_key=api_key)
        self._worker: Optional[_GenerateSlidesWorker] = None
        self._fetch_worker: Optional[_FetchTranslationsWorker] = None
        self._preview_worker: Optional[_FetchPreviewWorker] = None
        self._translate_worker: Optional[_TranslateWorker] = None

        self.result_pptx_path: Optional[str] = None
        self.result_section_name: str = ""

        self._default_font_name = default_font_name
        self._default_font_size = default_font_size

        # All available translations
        self._all_translations: List[BibleTranslation] = [
            BibleTranslation(**t)
            for t in BUILTIN_TRANSLATIONS
        ]

        # Cached preview data: {col_idx: (language_code, [verse_text, ...])}
        self._preview_original: Dict[int, Tuple[str, List[str]]] = {}
        self._preview_translated: Dict[int, List[str]] = {}
        self._showing_translated = False

        self._setup_ui()
        self._connect_signals()
        self._populate_available_list()

    # ------------------------------------------------------------------
    # UI setup
    # ------------------------------------------------------------------

    def _setup_ui(self) -> None:
        self.setWindowTitle(tr("dialog.bible.title"))
        self.setMinimumSize(1020, 680)
        self.resize(1100, 740)

        outer = QVBoxLayout(self)

        # ---- Top: Reference input ----
        ref_group = QGroupBox(tr("dialog.bible.reference_group"))
        ref_form = QFormLayout(ref_group)

        self.reference_edit = QLineEdit()
        self.reference_edit.setPlaceholderText(tr("dialog.bible.reference_placeholder"))
        ref_form.addRow(tr("dialog.bible.reference_label"), self.reference_edit)

        self.reference_status = QLabel("")
        self.reference_status.setStyleSheet("color: grey; font-size: 11px;")
        ref_form.addRow("", self.reference_status)

        outer.addWidget(ref_group)

        # ---- API key warning (shown when key is not configured) ----
        self._no_key_label = QLabel(tr("dialog.bible.no_api_key_warning"))
        self._no_key_label.setStyleSheet(
            "background:#f8d7da;color:#721c24;border:1px solid #f5c6cb;"
            "border-radius:3px;padding:6px;"
        )
        self._no_key_label.setWordWrap(True)
        self._no_key_label.setVisible(not self._bible_service.has_api_key())
        outer.addWidget(self._no_key_label)

        # ---- Middle: main splitter ----
        main_splitter = QSplitter(Qt.Orientation.Horizontal)

        # ======== LEFT PANEL: translation selector ========
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        trans_group = QGroupBox(tr("dialog.bible.translations_group"))
        trans_layout = QVBoxLayout(trans_group)

        # Language filter
        lang_row = QHBoxLayout()
        lang_row.addWidget(QLabel(tr("dialog.bible.filter_language")))
        self.language_combo = QComboBox()
        self.language_combo.addItem(tr("dialog.bible.all_languages"), "")
        lang_row.addWidget(self.language_combo, 1)
        self._fetch_btn = QPushButton(tr("dialog.bible.fetch_more"))
        self._fetch_btn.setFixedWidth(100)
        lang_row.addWidget(self._fetch_btn)
        trans_layout.addLayout(lang_row)

        # Available list
        avail_label = QLabel(tr("dialog.bible.available_translations"))
        avail_label.setStyleSheet("font-weight: bold;")
        trans_layout.addWidget(avail_label)
        self.available_list = QListWidget()
        self.available_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.available_list.setMaximumHeight(160)
        trans_layout.addWidget(self.available_list)
        trans_layout.addWidget(QLabel(tr("dialog.bible.available_hint")))

        # Add button
        add_row = QHBoxLayout()
        self._add_btn = QPushButton("↓ " + tr("dialog.bible.add_translation"))
        add_row.addWidget(self._add_btn)
        add_row.addStretch()
        trans_layout.addLayout(add_row)

        # Selected translations table
        sel_label = QLabel(tr("dialog.bible.selected_translations"))
        sel_label.setStyleSheet("font-weight: bold;")
        trans_layout.addWidget(sel_label)

        self.selected_table = QTableWidget(0, 3)
        self.selected_table.setHorizontalHeaderLabels([
            tr("dialog.bible.col_translation"),
            tr("dialog.bible.col_ref_override"),
            "",
        ])
        self.selected_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.selected_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        self.selected_table.horizontalHeader().resizeSection(1, 160)
        self.selected_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.selected_table.horizontalHeader().resizeSection(2, 28)
        self.selected_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.selected_table.setAlternatingRowColors(True)
        self.selected_table.verticalHeader().setVisible(False)
        self.selected_table.setMinimumHeight(140)
        trans_layout.addWidget(self.selected_table)
        trans_layout.addWidget(QLabel(tr("dialog.bible.ref_override_hint")))

        left_layout.addWidget(trans_group)

        # Font settings
        font_group = QGroupBox(tr("dialog.bible.font_group"))
        font_form = QFormLayout(font_group)
        self.font_name_edit = QLineEdit(self._default_font_name)
        font_form.addRow(tr("dialog.bible.font_name"), self.font_name_edit)
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(6, 36)
        self.font_size_spin.setValue(self._default_font_size)
        self.font_size_spin.setSuffix(" pt")
        font_form.addRow(tr("dialog.bible.font_size"), self.font_size_spin)
        left_layout.addWidget(font_group)
        left_layout.addStretch()

        main_splitter.addWidget(left_widget)

        # ======== RIGHT PANEL: preview ========
        preview_group = QGroupBox(tr("dialog.bible.preview_group"))
        preview_layout = QVBoxLayout(preview_group)

        # Preview controls
        ctrl_row = QHBoxLayout()
        self._preview_btn = QPushButton(tr("dialog.bible.fetch_preview"))
        ctrl_row.addWidget(self._preview_btn)

        self._toggle_btn = QPushButton(tr("dialog.bible.show_translated"))
        self._toggle_btn.setCheckable(True)
        self._toggle_btn.setEnabled(False)
        ctrl_row.addWidget(self._toggle_btn)

        ctrl_row.addWidget(QLabel(tr("dialog.bible.translate_to")))
        self.translate_lang_combo = QComboBox()
        for code, name in TRANSLATE_LANGUAGES:
            self.translate_lang_combo.addItem(name, code)
        ctrl_row.addWidget(self.translate_lang_combo)
        ctrl_row.addStretch()
        preview_layout.addLayout(ctrl_row)

        # Discrepancy warning
        self._discrepancy_label = QLabel("")
        self._discrepancy_label.setStyleSheet(
            "background:#fff3cd;color:#856404;border:1px solid #ffc107;"
            "border-radius:3px;padding:4px;"
        )
        self._discrepancy_label.setWordWrap(True)
        self._discrepancy_label.setVisible(False)
        preview_layout.addWidget(self._discrepancy_label)

        # Preview area: one QTextBrowser per translation column
        self._preview_scroll = QScrollArea()
        self._preview_scroll.setWidgetResizable(True)
        self._preview_container = QWidget()
        self._preview_columns_layout = QHBoxLayout(self._preview_container)
        self._preview_columns_layout.setSpacing(8)
        self._preview_scroll.setWidget(self._preview_container)
        preview_layout.addWidget(self._preview_scroll)

        self._preview_browsers: List[QTextBrowser] = []

        main_splitter.addWidget(preview_group)
        main_splitter.setSizes([420, 660])
        outer.addWidget(main_splitter, 1)

        # ---- Bottom: progress + buttons ----
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setVisible(False)
        outer.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(self.status_label)

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
        self._fetch_btn.clicked.connect(self._on_fetch_more)
        self.available_list.itemDoubleClicked.connect(self._on_add_translation)
        self.reference_edit.textChanged.connect(self._on_reference_changed)
        self._preview_btn.clicked.connect(self._on_fetch_preview)
        self._toggle_btn.clicked.connect(self._on_toggle_translation)

    # ------------------------------------------------------------------
    # Available list population
    # ------------------------------------------------------------------

    def _populate_available_list(self, filter_lang: str = "") -> None:
        self.available_list.clear()
        selected_ids = self._selected_ids()
        seen = set()
        for t in self._all_translations:
            if filter_lang and t.language != filter_lang:
                continue
            if t.id in seen or t.id in selected_ids:
                continue
            seen.add(t.id)
            item = QListWidgetItem(f"{t.abbreviation} – {t.name} ({t.language_name})")
            item.setData(Qt.ItemDataRole.UserRole, t.id)
            item.setData(Qt.ItemDataRole.UserRole + 1, t)
            self.available_list.addItem(item)

    def _selected_ids(self) -> set:
        ids = set()
        for row in range(self.selected_table.rowCount()):
            item = self.selected_table.item(row, 0)
            if item:
                ids.add(item.data(Qt.ItemDataRole.UserRole))
        return ids

    # ------------------------------------------------------------------
    # Selected table helpers
    # ------------------------------------------------------------------

    def _add_row_to_table(self, trans: BibleTranslation) -> None:
        row = self.selected_table.rowCount()
        self.selected_table.insertRow(row)

        # Col 0: translation name (read-only)
        name_item = QTableWidgetItem(f"{trans.abbreviation} – {trans.name}")
        name_item.setData(Qt.ItemDataRole.UserRole, trans.id)
        name_item.setData(Qt.ItemDataRole.UserRole + 1, trans)
        name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.selected_table.setItem(row, 0, name_item)

        # Col 1: reference override (editable)
        override_item = QTableWidgetItem("")
        override_item.setToolTip(tr("dialog.bible.ref_override_tooltip"))
        self.selected_table.setItem(row, 1, override_item)

        # Col 2: remove button
        remove_btn = QPushButton("✕")
        remove_btn.setFixedSize(26, 26)
        remove_btn.setToolTip(tr("dialog.bible.remove_translation"))
        remove_btn.clicked.connect(lambda _, r=row: self._remove_row(r))
        self.selected_table.setCellWidget(row, 2, remove_btn)

        # Fix row button references (rows shift when others are removed)
        self._rewire_remove_buttons()

    def _rewire_remove_buttons(self) -> None:
        for r in range(self.selected_table.rowCount()):
            btn = self.selected_table.cellWidget(r, 2)
            if btn:
                try:
                    btn.clicked.disconnect()
                except RuntimeError:
                    pass
                # capture r by value
                btn.clicked.connect(lambda _, row=r: self._remove_row(row))

    def _remove_row(self, row: int) -> None:
        self.selected_table.removeRow(row)
        self._rewire_remove_buttons()
        self._populate_available_list(filter_lang=self.language_combo.currentData() or "")
        self._clear_preview()

    def _get_selected_slots(self) -> List[Tuple[BibleTranslation, str]]:
        """Return list of (translation, effective_reference) in table order."""
        slots = []
        main_ref = self.reference_edit.text().strip()
        for row in range(self.selected_table.rowCount()):
            name_item = self.selected_table.item(row, 0)
            override_item = self.selected_table.item(row, 1)
            if not name_item:
                continue
            trans: BibleTranslation = name_item.data(Qt.ItemDataRole.UserRole + 1)
            override = (override_item.text().strip() if override_item else "") or main_ref
            slots.append((trans, override))
        return slots

    def _get_reference_overrides(self) -> Dict[int, str]:
        main_ref = self.reference_edit.text().strip()
        overrides: Dict[int, str] = {}
        for row in range(self.selected_table.rowCount()):
            name_item = self.selected_table.item(row, 0)
            override_item = self.selected_table.item(row, 1)
            if not name_item:
                continue
            vid = name_item.data(Qt.ItemDataRole.UserRole)
            override = (override_item.text().strip() if override_item else "")
            if override and override != main_ref:
                overrides[vid] = override
        return overrides

    # ------------------------------------------------------------------
    # Slots – translation selection
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
            self.reference_status.setText(f"✓ {ref.display_str}")
            self.reference_status.setStyleSheet("color: green; font-size: 11px;")
        except ValueError as exc:
            self.reference_status.setText(str(exc))
            self.reference_status.setStyleSheet("color: red; font-size: 11px;")

    def _on_add_translation(self) -> None:
        selected_ids = self._selected_ids()
        for item in self.available_list.selectedItems():
            if self.selected_table.rowCount() >= 6:
                QMessageBox.information(self, tr("dialog.bible.title"), tr("dialog.bible.max_translations"))
                break
            t_id = item.data(Qt.ItemDataRole.UserRole)
            if t_id in selected_ids:
                continue
            trans: BibleTranslation = item.data(Qt.ItemDataRole.UserRole + 1)
            self._add_row_to_table(trans)
            selected_ids.add(t_id)

        self._populate_available_list(filter_lang=self.language_combo.currentData() or "")
        self._clear_preview()

    def _on_fetch_more(self) -> None:
        lang = self.language_combo.currentData() or ""
        if not lang:
            QMessageBox.information(self, tr("dialog.bible.title"), tr("dialog.bible.select_language_first"))
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

        existing_ids = {t.id for t in self._all_translations}
        added = 0
        for t in translations:
            if t.id not in existing_ids:
                self._all_translations.append(t)
                existing_ids.add(t.id)
                added += 1

        existing_langs = {self.language_combo.itemData(i) for i in range(self.language_combo.count())}
        for t in translations:
            if t.language and t.language not in existing_langs:
                self.language_combo.addItem(f"{t.language_name} ({t.language.upper()})", t.language)
                existing_langs.add(t.language)

        self._populate_available_list(filter_lang=self.language_combo.currentData() or "")
        self.status_label.setText(tr("dialog.bible.fetch_done", count=added))

    def _on_fetch_translations_error(self, error: str) -> None:
        self._fetch_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_label.setText(tr("dialog.bible.fetch_error", error=error))

    # ------------------------------------------------------------------
    # Slots – preview
    # ------------------------------------------------------------------

    def _clear_preview(self) -> None:
        """Remove all preview column widgets."""
        while self._preview_columns_layout.count():
            item = self._preview_columns_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._preview_browsers.clear()
        self._preview_original.clear()
        self._preview_translated.clear()
        self._showing_translated = False
        self._toggle_btn.setChecked(False)
        self._toggle_btn.setEnabled(False)
        self._toggle_btn.setText(tr("dialog.bible.show_translated"))
        self._discrepancy_label.setVisible(False)

    def _build_preview_columns(self, n: int) -> None:
        """Create empty preview column widgets for n translations."""
        self._clear_preview()
        for col_idx in range(n):
            col_widget = QWidget()
            col_layout = QVBoxLayout(col_widget)
            col_layout.setContentsMargins(0, 0, 0, 0)
            header = QLabel("…")
            header.setStyleSheet("font-weight: bold; color: #444;")
            col_layout.addWidget(header)
            browser = QTextBrowser()
            browser.setOpenExternalLinks(False)
            browser.setReadOnly(True)
            col_layout.addWidget(browser)
            self._preview_columns_layout.addWidget(col_widget)
            self._preview_browsers.append(browser)
            # Store header ref
            browser.setProperty("_header", header)

    def _on_fetch_preview(self) -> None:
        reference = self.reference_edit.text().strip()
        if not reference:
            QMessageBox.warning(self, tr("dialog.bible.title"), tr("dialog.bible.error.no_reference"))
            return
        try:
            parse_reference(reference)
        except ValueError as exc:
            QMessageBox.warning(self, tr("dialog.bible.title"), str(exc))
            return

        slots_data = self._get_selected_slots()
        if not slots_data:
            QMessageBox.warning(self, tr("dialog.bible.title"), tr("dialog.bible.error.no_translations"))
            return

        self._build_preview_columns(len(slots_data))

        worker_slots = [
            (col_idx, trans, ref_str)
            for col_idx, (trans, ref_str) in enumerate(slots_data)
        ]

        self._preview_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.status_label.setText(tr("dialog.bible.status.fetching"))

        self._preview_worker = _FetchPreviewWorker(worker_slots, self._bible_service, self)
        self._preview_worker.column_ready.connect(self._on_preview_column_ready)
        self._preview_worker.finished.connect(self._on_preview_finished)
        self._preview_worker.error.connect(self._on_preview_error)
        self._preview_worker.start()

    def _on_preview_column_ready(
        self, col_idx: int, trans_name: str, language: str, texts: List[str]
    ) -> None:
        if col_idx >= len(self._preview_browsers):
            return
        browser = self._preview_browsers[col_idx]
        header: QLabel = browser.property("_header")
        if header:
            header.setText(trans_name)
        self._preview_original[col_idx] = (language, texts)
        self._render_column_original(col_idx, language, texts)

    def _render_column_original(self, col_idx: int, language: str, texts: List[str]) -> None:
        if col_idx >= len(self._preview_browsers):
            return
        browser = self._preview_browsers[col_idx]
        rtl = is_rtl(language)
        dir_attr = 'dir="rtl"' if rtl else ""
        align = "right" if rtl else "left"
        html_parts = []
        for t in texts:
            html_parts.append(
                f'<p {dir_attr} style="text-align:{align};margin:2px 0;">{_escape_html(t)}</p>'
            )
        browser.setHtml("".join(html_parts))

    def _render_column_translated(self, col_idx: int, texts: List[str]) -> None:
        if col_idx >= len(self._preview_browsers):
            return
        target_lang = self.translate_lang_combo.currentData() or "nl"
        rtl = is_rtl(target_lang)
        dir_attr = 'dir="rtl"' if rtl else ""
        align = "right" if rtl else "left"
        browser = self._preview_browsers[col_idx]
        html_parts = []
        for t in texts:
            html_parts.append(
                f'<p {dir_attr} style="text-align:{align};margin:2px 0;">{_escape_html(t)}</p>'
            )
        browser.setHtml("".join(html_parts))

    def _on_preview_finished(self) -> None:
        self._preview_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_label.setText("")
        self._toggle_btn.setEnabled(True)

        # Check for verse-count discrepancies
        counts = {col: len(texts) for col, (_, texts) in self._preview_original.items()}
        if len(set(counts.values())) > 1:
            detail = ", ".join(
                f"{self.selected_table.item(c, 0).text() if self.selected_table.item(c, 0) else str(c)}: {n}"
                for c, n in sorted(counts.items())
            )
            self._discrepancy_label.setText(
                tr("dialog.bible.discrepancy_warning", detail=detail)
            )
            self._discrepancy_label.setVisible(True)
        else:
            self._discrepancy_label.setVisible(False)

    def _on_preview_error(self, error: str) -> None:
        self._preview_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_label.setText(tr("dialog.bible.fetch_error", error=error))

    def _on_toggle_translation(self, checked: bool) -> None:
        if not checked:
            # Switch back to original
            self._showing_translated = False
            self._toggle_btn.setText(tr("dialog.bible.show_translated"))
            for col_idx, (lang, texts) in self._preview_original.items():
                self._render_column_original(col_idx, lang, texts)
            return

        # Switch to translated
        self._toggle_btn.setText(tr("dialog.bible.show_original"))
        target_lang = self.translate_lang_combo.currentData() or "nl"

        # Check if we already have translations for this target
        if self._preview_translated:
            self._showing_translated = True
            for col_idx, texts in self._preview_translated.items():
                self._render_column_translated(col_idx, texts)
            return

        # Need to fetch
        columns_to_translate = [
            (col_idx, texts)
            for col_idx, (_, texts) in self._preview_original.items()
        ]
        if not columns_to_translate:
            self._toggle_btn.setChecked(False)
            return

        self._toggle_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.status_label.setText(tr("dialog.bible.status.translating"))

        self._translate_worker = _TranslateWorker(columns_to_translate, target_lang, self)
        self._translate_worker.column_ready.connect(self._on_translate_column_ready)
        self._translate_worker.finished.connect(self._on_translate_finished)
        self._translate_worker.error.connect(self._on_translate_error)
        self._translate_worker.start()

    def _on_translate_column_ready(self, col_idx: int, texts: List[str]) -> None:
        self._preview_translated[col_idx] = texts
        self._render_column_translated(col_idx, texts)

    def _on_translate_finished(self) -> None:
        self._toggle_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_label.setText("")
        self._showing_translated = True

    def _on_translate_error(self, error: str) -> None:
        self._toggle_btn.setEnabled(True)
        self._toggle_btn.setChecked(False)
        self.progress_bar.setVisible(False)
        self.status_label.setText(tr("dialog.bible.fetch_error", error=error))

    # ------------------------------------------------------------------
    # Generate
    # ------------------------------------------------------------------

    def _on_generate(self) -> None:
        reference = self.reference_edit.text().strip()
        if not reference:
            QMessageBox.warning(self, tr("dialog.bible.title"), tr("dialog.bible.error.no_reference"))
            return
        try:
            parse_reference(reference)
        except ValueError as exc:
            QMessageBox.warning(self, tr("dialog.bible.title"), str(exc))
            return
        if self.selected_table.rowCount() == 0:
            QMessageBox.warning(self, tr("dialog.bible.title"), tr("dialog.bible.error.no_translations"))
            return

        translation_ids = []
        for row in range(self.selected_table.rowCount()):
            item = self.selected_table.item(row, 0)
            if item:
                translation_ids.append(item.data(Qt.ItemDataRole.UserRole))

        config = BibleSlideConfig(
            font_name=self.font_name_edit.text().strip() or "Calibri",
            font_size=self.font_size_spin.value(),
        )
        overrides = self._get_reference_overrides()

        self._set_generating(True)

        self._worker = _GenerateSlidesWorker(
            reference, translation_ids, config, overrides,
            api_key=self._bible_service._api_key, parent=self,
        )
        self._worker.finished.connect(self._on_generation_done)
        self._worker.error.connect(self._on_generation_error)
        self._worker.progress.connect(self._on_generation_progress)
        self._worker.start()

    def _on_generation_done(self, path: str) -> None:
        self._set_generating(False)
        self.result_pptx_path = path
        ref_text = self.reference_edit.text().strip()
        abbrevs = " / ".join(
            self.selected_table.item(r, 0).data(Qt.ItemDataRole.UserRole + 1).abbreviation
            for r in range(self.selected_table.rowCount())
            if self.selected_table.item(r, 0)
        )
        self.result_section_name = f"{ref_text} ({abbrevs})"
        self.accept()

    def _on_generation_error(self, error: str) -> None:
        self._set_generating(False)
        QMessageBox.critical(self, tr("dialog.bible.title"), tr("dialog.bible.error.generation", error=error))

    def _on_generation_progress(self, msg: str) -> None:
        self.status_label.setText(msg)

    def _set_generating(self, generating: bool) -> None:
        self.ok_btn.setEnabled(not generating)
        self.progress_bar.setVisible(generating)
        if not generating:
            self.status_label.setText("")
        else:
            self.status_label.setText(tr("dialog.bible.status.generating"))


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def _escape_html(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br>")
