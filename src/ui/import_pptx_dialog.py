"""Dialog for importing songs from existing PPTX presentations."""

import os
from collections import defaultdict
from datetime import date
from typing import Dict, List, Optional

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QPushButton,
    QLabel,
    QTreeWidget,
    QTreeWidgetItem,
    QProgressBar,
    QMessageBox,
    QFileDialog,
    QDialogButtonBox,
    QGroupBox,
    QListWidget,
    QMenu,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QBrush, QColor, QFont, QAction

from ..models import Settings
from ..services.pptx_scanner_service import (
    PptxScannerService,
    PptxScanResult,
    SongStatus,
    SongClassification,
)
from ..services.excel_service import ExcelService
from ..i18n import tr
from ..logging_config import get_logger

logger = get_logger("import_pptx_dialog")

# Colours used for song-status child items
_COLOR_EXCLUDED = QColor(128, 128, 128)   # grey — liturgical items
_COLOR_UNKNOWN = QColor(200, 120, 0)       # amber — not in library
_COLOR_REGISTERED = QColor(128, 128, 128)  # grey — already in register

# Custom data roles for child tree items
_ROLE_CLS = Qt.ItemDataRole.UserRole        # SongClassification
_ROLE_FILEPATH = Qt.ItemDataRole.UserRole.value + 1  # source pptx filepath


class _ScanWorker(QThread):
    """Background worker that scans PPTX files one by one."""

    progress = pyqtSignal(int, int)          # current, total
    file_scanned = pyqtSignal(object)        # PptxScanResult
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, folder_path: str, service: PptxScannerService, parent=None):
        super().__init__(parent)
        self.folder_path = folder_path
        self.service = service
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def run(self) -> None:
        try:
            if not os.path.isdir(self.folder_path):
                self.error.emit(tr("dialog.import_pptx.folder_not_found").format(
                    path=self.folder_path
                ))
                return

            filenames = sorted(
                f for f in os.listdir(self.folder_path)
                if f.lower().endswith(".pptx") and not f.startswith("~")
            )
            total = len(filenames)
            if total == 0:
                self.error.emit(tr("dialog.import_pptx.no_files"))
                return

            for i, filename in enumerate(filenames):
                if self._cancelled:
                    break
                filepath = os.path.join(self.folder_path, filename)
                result = self.service.scan_file(filepath)
                self.file_scanned.emit(result)
                self.progress.emit(i + 1, total)

            self.finished.emit()
        except Exception as exc:
            logger.error("Scan worker error: %s", exc, exc_info=True)
            self.error.emit(str(exc))


def _make_bold(item: QTreeWidgetItem, num_cols: int = 4) -> None:
    """Apply bold font to every column of *item*."""
    font = QFont()
    font.setBold(True)
    for col in range(num_cols):
        item.setFont(col, font)


class _ImportValidationDialog(QDialog):
    """Shows songs to be imported and lets the user confirm or deselect them."""

    def __init__(self, rows: List[dict], unknown_titles: List[tuple], parent=None):
        super().__init__(parent)
        self._tree: Optional[QTreeWidget] = None
        self._setup_ui(rows, unknown_titles)

    def _setup_ui(self, rows: List[dict], unknown_titles: List[tuple]) -> None:
        self.setWindowTitle(tr("dialog.import_pptx.validate_title"))
        self.setMinimumSize(700, 500)
        self.resize(820, 580)

        layout = QVBoxLayout(self)

        # Summary line
        new_rows = [r for r in rows if not r["in_register"]]
        file_count = len({r["file"] for r in new_rows})
        summary = tr("dialog.import_pptx.validate_summary").format(
            count=len(new_rows), files=file_count
        )
        layout.addWidget(QLabel(summary))

        # Tree: Song | File | Date | Status
        self._tree = QTreeWidget()
        self._tree.setHeaderLabels([
            tr("dialog.import_pptx.col_songs"),
            tr("dialog.import_pptx.col_file"),
            tr("dialog.import_pptx.col_date"),
            "Status",
        ])
        self._tree.setColumnWidth(0, 260)
        self._tree.setColumnWidth(1, 200)
        self._tree.setColumnWidth(2, 100)
        self._tree.setAlternatingRowColors(True)

        # Section headers (not checkable, bold)
        new_header = QTreeWidgetItem([tr("dialog.import_pptx.validate_new"), "", "", ""])
        new_header.setFlags(new_header.flags() & ~Qt.ItemFlag.ItemIsUserCheckable)
        _make_bold(new_header)

        reg_header = QTreeWidgetItem(
            [tr("dialog.import_pptx.validate_registered"), "", "", ""]
        )
        reg_header.setFlags(reg_header.flags() & ~Qt.ItemFlag.ItemIsUserCheckable)
        _make_bold(reg_header)

        has_new = False
        has_reg = False

        for row in rows:
            child = QTreeWidgetItem()
            child.setData(0, Qt.ItemDataRole.UserRole, row)
            child.setText(0, row["title"])
            child.setText(1, row["file"])
            child.setText(2, row["date"].isoformat() if row["date"] else "")
            status_key = f"dialog.import_pptx.status_{row['status'].value}"
            child.setText(3, tr(status_key))
            child.setFlags(child.flags() | Qt.ItemFlag.ItemIsUserCheckable)

            if not row["in_register"]:
                child.setCheckState(0, Qt.CheckState.Checked)
                new_header.addChild(child)
                has_new = True
            else:
                child.setCheckState(0, Qt.CheckState.Unchecked)
                grey = QBrush(_COLOR_REGISTERED)
                for col in range(4):
                    child.setForeground(col, grey)
                reg_header.addChild(child)
                has_reg = True

        if has_new:
            self._tree.addTopLevelItem(new_header)
            new_header.setExpanded(True)
        if has_reg:
            self._tree.addTopLevelItem(reg_header)
            reg_header.setExpanded(False)

        layout.addWidget(self._tree, 1)

        # "Not in library" group for UNKNOWN songs
        if unknown_titles:
            group = QGroupBox(tr("dialog.import_pptx.not_in_library"))
            g_layout = QVBoxLayout(group)
            lst = QListWidget()
            lst.setMaximumHeight(100)
            for title, filename in unknown_titles:
                lst.addItem(f"{title}  ({filename})")
            g_layout.addWidget(lst)
            layout.addWidget(group)

        # Buttons
        btn_box = QDialogButtonBox()
        btn_box.addButton(
            tr("dialog.import_pptx.confirm_import"),
            QDialogButtonBox.ButtonRole.AcceptRole,
        )
        btn_box.addButton(QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def result_rows(self) -> List[dict]:
        """Return the rows the user left checked (i.e. approved for import)."""
        checked: List[dict] = []
        if self._tree is None:
            return checked
        for i in range(self._tree.topLevelItemCount()):
            top = self._tree.topLevelItem(i)
            for j in range(top.childCount()):
                child = top.child(j)
                if child.checkState(0) == Qt.CheckState.Checked:
                    row = child.data(0, Qt.ItemDataRole.UserRole)
                    if row is not None:
                        checked.append(row)
        return checked


class ImportPptxDialog(QDialog):
    """Dialog that scans a folder of PPTX files and imports detected songs
    into LiederenRegister.xlsx."""

    def __init__(self, settings: Settings, base_path: str, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.base_path = base_path
        self._scanner = PptxScannerService()
        self._worker: Optional[_ScanWorker] = None
        self._results: List[PptxScanResult] = []
        self._songs_path = settings.get_songs_path(base_path)
        self._algemeen_path = settings.get_algemeen_path(base_path)

        self._setup_ui()
        self._connect_signals()

        # Pre-fill the folder from settings
        archive_path = settings.get_pptx_archive_path(base_path)
        self.folder_input.setText(archive_path)

    # ------------------------------------------------------------------
    # UI setup
    # ------------------------------------------------------------------

    def _setup_ui(self) -> None:
        self.setWindowTitle(tr("dialog.import_pptx.title"))
        self.setMinimumSize(700, 560)
        self.resize(820, 660)

        layout = QVBoxLayout(self)

        # --- Folder selection row ---
        folder_row = QHBoxLayout()
        folder_label = QLabel(tr("dialog.import_pptx.folder_label"))
        folder_row.addWidget(folder_label)

        self.folder_input = QLineEdit()
        folder_row.addWidget(self.folder_input, 1)

        self.browse_btn = QPushButton(tr("button.browse"))
        folder_row.addWidget(self.browse_btn)

        self.scan_btn = QPushButton(tr("dialog.import_pptx.scan"))
        self.scan_btn.setDefault(True)
        folder_row.addWidget(self.scan_btn)

        layout.addLayout(folder_row)

        # --- Progress bar (hidden until scanning) ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)

        # --- Results tree ---
        self.results_tree = QTreeWidget()
        self.results_tree.setHeaderLabels([
            tr("dialog.import_pptx.col_file"),
            tr("dialog.import_pptx.col_date"),
            tr("dialog.import_pptx.col_songs"),
        ])
        self.results_tree.setColumnWidth(0, 320)
        self.results_tree.setColumnWidth(1, 110)
        self.results_tree.setColumnWidth(2, 80)
        self.results_tree.setAlternatingRowColors(True)
        layout.addWidget(self.results_tree, 1)

        # --- "Not in library" panel ---
        self.not_in_library_group = QGroupBox(tr("dialog.import_pptx.not_in_library"))
        nil_layout = QVBoxLayout(self.not_in_library_group)
        self.not_in_library_tree = QTreeWidget()
        self.not_in_library_tree.setHeaderLabels([
            tr("dialog.import_pptx.col_songs"),
            tr("dialog.import_pptx.col_source"),
        ])
        self.not_in_library_tree.setColumnWidth(0, 320)
        self.not_in_library_tree.setMaximumHeight(120)
        nil_layout.addWidget(self.not_in_library_tree)
        self.not_in_library_group.hide()
        layout.addWidget(self.not_in_library_group)

        # --- Status label ---
        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

        # --- Bottom button row ---
        btn_row = QHBoxLayout()
        self.select_all_btn = QPushButton(tr("dialog.import_pptx.select_all"))
        self.deselect_all_btn = QPushButton(tr("dialog.import_pptx.deselect_all"))
        btn_row.addWidget(self.select_all_btn)
        btn_row.addWidget(self.deselect_all_btn)
        btn_row.addStretch()

        self.import_btn = QPushButton(tr("dialog.import_pptx.import_excel"))
        self.import_btn.setEnabled(False)
        btn_row.addWidget(self.import_btn)

        close_btn = QPushButton(tr("button.close"))
        close_btn.clicked.connect(self.close)
        btn_row.addWidget(close_btn)

        layout.addLayout(btn_row)

    def _connect_signals(self) -> None:
        self.browse_btn.clicked.connect(self._on_browse)
        self.scan_btn.clicked.connect(self._on_scan)
        self.select_all_btn.clicked.connect(self._on_select_all)
        self.deselect_all_btn.clicked.connect(self._on_deselect_all)
        self.import_btn.clicked.connect(self._on_import)
        self.results_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.results_tree.customContextMenuRequested.connect(self._on_context_menu)
        self.results_tree.itemChanged.connect(self._on_child_item_changed)

    # ------------------------------------------------------------------
    # Slots
    # ------------------------------------------------------------------

    def _on_browse(self) -> None:
        current = self.folder_input.text()
        if not current or not os.path.isdir(current):
            current = self.base_path

        folder = QFileDialog.getExistingDirectory(
            self,
            tr("dialog.import_pptx.folder_label"),
            current,
        )
        if folder:
            self.folder_input.setText(folder)

    def _on_scan(self) -> None:
        folder = self.folder_input.text().strip()
        if not folder:
            QMessageBox.warning(self, self.windowTitle(), tr("dialog.import_pptx.no_folder"))
            return
        if not os.path.isdir(folder):
            QMessageBox.warning(
                self,
                self.windowTitle(),
                tr("dialog.import_pptx.folder_not_found").format(path=folder),
            )
            return

        # Cancel any running scan
        if self._worker and self._worker.isRunning():
            self._worker.cancel()
            self._worker.wait()

        # Reset UI
        self.results_tree.clear()
        self._results.clear()
        self.import_btn.setEnabled(False)
        self.status_label.setText("")
        self.progress_bar.setValue(0)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.show()
        self.scan_btn.setEnabled(False)
        self.not_in_library_group.hide()
        self.not_in_library_tree.clear()

        # Refresh library paths in case settings changed
        self._songs_path = self.settings.get_songs_path(self.base_path)
        self._algemeen_path = self.settings.get_algemeen_path(self.base_path)

        self._worker = _ScanWorker(folder, self._scanner, parent=self)
        self._worker.file_scanned.connect(self._on_file_scanned)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_scan_finished)
        self._worker.error.connect(self._on_scan_error)
        self._worker.start()

    def _on_file_scanned(self, result: PptxScanResult) -> None:
        # Classify detected songs against the library
        if result.songs and not result.error:
            try:
                result.song_classifications = self._scanner.classify_songs(
                    result.songs,
                    self._songs_path,
                    self._algemeen_path,
                    user_liturgy_items=self.settings.user_liturgy_items,
                    phase3_candidates=result.phase3_candidates,
                )
            except Exception as exc:
                logger.warning(
                    "Classification failed for %s: %s", result.filename, exc
                )

        self._results.append(result)

        item = QTreeWidgetItem()
        item.setData(0, Qt.ItemDataRole.UserRole, result)

        # Checkable at the file level
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        has_songs = bool(result.songs) and not result.error
        item.setCheckState(
            0,
            Qt.CheckState.Checked if has_songs else Qt.CheckState.Unchecked,
        )

        item.setText(0, result.filename)

        if result.service_date:
            item.setText(1, result.service_date.isoformat())
        else:
            item.setText(1, tr("dialog.import_pptx.no_date"))

        if result.error:
            item.setText(2, tr("dialog.import_pptx.scan_error").format(error=result.error))
        elif result.songs:
            # Column 2: count of importable (non-excluded) songs
            if result.song_classifications:
                importable_count = sum(
                    1 for c in result.song_classifications
                    if c.status != SongStatus.EXCLUDED
                )
            else:
                importable_count = len(result.songs)

            item.setText(
                2,
                tr("dialog.import_pptx.songs_found").format(count=importable_count),
            )

            # Add colour-coded song children (each individually checkable)
            if result.song_classifications:
                self.results_tree.blockSignals(True)
                for cls in result.song_classifications:
                    child = QTreeWidgetItem(item)
                    child.setData(0, _ROLE_CLS, cls)
                    child.setData(0, _ROLE_FILEPATH, result.filepath)
                    child.setFlags(
                        child.flags()
                        | Qt.ItemFlag.ItemIsUserCheckable
                        | Qt.ItemFlag.ItemIsEnabled
                    )
                    if cls.status == SongStatus.EXCLUDED:
                        child.setCheckState(0, Qt.CheckState.Unchecked)
                        suffix = tr("dialog.import_pptx.status_excluded")
                        child.setText(0, f"  {cls.title} — {suffix}")
                        child.setForeground(0, QBrush(_COLOR_EXCLUDED))
                        font = QFont()
                        font.setItalic(True)
                        child.setFont(0, font)
                    elif cls.status == SongStatus.UNKNOWN:
                        child.setCheckState(0, Qt.CheckState.Checked)
                        suffix = tr("dialog.import_pptx.status_unknown")
                        child.setText(0, f"  {cls.title} — {suffix}")
                        child.setForeground(0, QBrush(_COLOR_UNKNOWN))
                    else:  # CONFIRMED
                        child.setCheckState(0, Qt.CheckState.Checked)
                        child.setText(0, f"  {cls.title}")
                self.results_tree.blockSignals(False)
            else:
                self.results_tree.blockSignals(True)
                for song_title in result.songs:
                    child = QTreeWidgetItem(item)
                    child.setText(0, f"  {song_title}")
                    child.setData(0, _ROLE_FILEPATH, result.filepath)
                    child.setFlags(
                        child.flags()
                        | Qt.ItemFlag.ItemIsUserCheckable
                        | Qt.ItemFlag.ItemIsEnabled
                    )
                    child.setCheckState(0, Qt.CheckState.Checked)
                self.results_tree.blockSignals(False)
        else:
            item.setText(2, tr("dialog.import_pptx.no_songs"))

        self.results_tree.addTopLevelItem(item)
        item.setExpanded(True)

    def _on_progress(self, current: int, total: int) -> None:
        self.progress_bar.setRange(0, total)
        self.progress_bar.setValue(current)

    def _on_scan_finished(self) -> None:
        self.progress_bar.hide()
        self.scan_btn.setEnabled(True)
        count = len(self._results)
        self.status_label.setText(
            tr("dialog.import_pptx.scan_complete").format(count=count)
        )
        has_any_songs = any(r.songs for r in self._results if not r.error)
        self.import_btn.setEnabled(has_any_songs)

        # Populate "not in library" panel with UNKNOWN songs
        self.not_in_library_tree.clear()
        for result in self._results:
            for cls in result.song_classifications:
                if cls.status == SongStatus.UNKNOWN:
                    row_item = QTreeWidgetItem()
                    row_item.setText(0, cls.title)
                    row_item.setText(1, result.filename)
                    self.not_in_library_tree.addTopLevelItem(row_item)

        if self.not_in_library_tree.topLevelItemCount() > 0:
            self.not_in_library_group.show()
        else:
            self.not_in_library_group.hide()

    def _on_scan_error(self, message: str) -> None:
        self.progress_bar.hide()
        self.scan_btn.setEnabled(True)
        self.status_label.setText(message)
        QMessageBox.warning(self, self.windowTitle(), message)

    def _on_context_menu(self, pos) -> None:
        item = self.results_tree.itemAt(pos)
        if item is None:
            return

        menu = QMenu(self)
        global_pos = self.results_tree.mapToGlobal(pos)

        if item.parent() is not None:
            # Child item — individual song/liturgy entry
            is_checked = item.checkState(0) == Qt.CheckState.Checked
            if is_checked:
                toggle_act = QAction(tr("dialog.import_pptx.mark_as_liturgy"), self)
            else:
                toggle_act = QAction(tr("dialog.import_pptx.mark_as_song"), self)
            menu.addAction(toggle_act)

            filepath = item.data(0, _ROLE_FILEPATH)
            open_act = None
            if filepath and os.path.exists(filepath):
                menu.addSeparator()
                open_act = QAction(tr("dialog.import_pptx.open_pptx"), self)
                menu.addAction(open_act)

            action = menu.exec(global_pos)
            if action == toggle_act:
                new_state = (
                    Qt.CheckState.Unchecked if is_checked else Qt.CheckState.Checked
                )
                item.setCheckState(0, new_state)
                # Persist the user's decision so future scans apply it automatically
                cls: Optional[SongClassification] = item.data(0, _ROLE_CLS)
                title = cls.title if cls else item.text(0).strip().split(" — ")[0].strip()
                self._update_user_liturgy_items(title, marking_as_liturgy=is_checked)
            elif open_act and action == open_act:
                self._open_file(filepath)
        else:
            # Parent (file-level) item
            result = item.data(0, Qt.ItemDataRole.UserRole)
            if result and result.filepath and os.path.exists(result.filepath):
                open_act = QAction(tr("dialog.import_pptx.open_pptx"), self)
                menu.addAction(open_act)
                action = menu.exec(global_pos)
                if action == open_act:
                    self._open_file(result.filepath)

    def _open_file(self, filepath: str) -> None:
        try:
            os.startfile(filepath)
        except Exception as exc:
            logger.warning("Could not open file %s: %s", filepath, exc)
            QMessageBox.warning(self, self.windowTitle(), str(exc))

    def _update_user_liturgy_items(self, title: str, marking_as_liturgy: bool) -> None:
        """Add or remove *title* from the persistent user liturgy items list."""
        from ..services.pptx_scanner_service import PptxScannerService
        norm = PptxScannerService._normalize(title)
        current_norms = [PptxScannerService._normalize(t) for t in self.settings.user_liturgy_items]

        if marking_as_liturgy:
            if norm not in current_norms:
                self.settings.user_liturgy_items.append(title)
                logger.debug("Added to user liturgy items: %s", title)
        else:
            self.settings.user_liturgy_items = [
                t for t, n in zip(self.settings.user_liturgy_items, current_norms)
                if n != norm
            ]
            logger.debug("Removed from user liturgy items: %s", title)

        try:
            self.settings.save()
        except Exception as exc:
            logger.warning("Could not save settings after liturgy item update: %s", exc)

    def _on_child_item_changed(self, item: QTreeWidgetItem, column: int) -> None:
        """Update child styling and parent count when a child checkbox is toggled."""
        if column != 0 or item.parent() is None:
            return

        checked = item.checkState(0) == Qt.CheckState.Checked
        cls: Optional[SongClassification] = item.data(0, _ROLE_CLS)

        self.results_tree.blockSignals(True)
        if checked:
            # Restore colour based on original auto-classification
            if cls and cls.status == SongStatus.UNKNOWN:
                item.setForeground(0, QBrush(_COLOR_UNKNOWN))
            else:
                item.setForeground(0, QBrush(QColor(0, 0, 0)))
            font = QFont()
            font.setItalic(False)
            item.setFont(0, font)
            # Restore clean title (strip any "— ..." status suffix added when excluded)
            if cls:
                title = cls.title
            else:
                raw = item.text(0).strip()
                title = raw.split(" — ")[0].strip()
            if cls and cls.status == SongStatus.UNKNOWN:
                suffix = tr("dialog.import_pptx.status_unknown")
                item.setText(0, f"  {title} — {suffix}")
            else:
                item.setText(0, f"  {title}")
        else:
            item.setForeground(0, QBrush(_COLOR_EXCLUDED))
            font = QFont()
            font.setItalic(True)
            item.setFont(0, font)
            title = cls.title if cls else item.text(0).strip().split(" — ")[0].strip()
            suffix = tr("dialog.import_pptx.status_excluded")
            item.setText(0, f"  {title} — {suffix}")
        self.results_tree.blockSignals(False)

        self._update_parent_song_count(item.parent())

    def _update_parent_song_count(self, parent: QTreeWidgetItem) -> None:
        """Recompute and display checked-child count on the parent row."""
        count = sum(
            1
            for j in range(parent.childCount())
            if parent.child(j).checkState(0) == Qt.CheckState.Checked
        )
        parent.setText(2, tr("dialog.import_pptx.songs_found").format(count=count))

    def _on_select_all(self) -> None:
        for i in range(self.results_tree.topLevelItemCount()):
            item = self.results_tree.topLevelItem(i)
            result: PptxScanResult = item.data(0, Qt.ItemDataRole.UserRole)
            if result and result.songs and not result.error:
                item.setCheckState(0, Qt.CheckState.Checked)

    def _on_deselect_all(self) -> None:
        for i in range(self.results_tree.topLevelItemCount()):
            item = self.results_tree.topLevelItem(i)
            item.setCheckState(0, Qt.CheckState.Unchecked)

    def _on_import(self) -> None:
        excel_path = self.settings.get_excel_register_path(self.base_path)
        if not excel_path:
            QMessageBox.warning(self, self.windowTitle(), tr("dialog.import_pptx.no_excel"))
            return
        if not os.path.exists(excel_path):
            QMessageBox.warning(
                self,
                self.windowTitle(),
                tr("error.file_not_found").format(path=excel_path),
            )
            return

        excel_service = ExcelService(excel_path)

        # Build song rows from checked files — respect per-child checkbox overrides
        all_rows: List[dict] = []
        for i in range(self.results_tree.topLevelItemCount()):
            item = self.results_tree.topLevelItem(i)
            if item.checkState(0) != Qt.CheckState.Checked:
                continue
            result: PptxScanResult = item.data(0, Qt.ItemDataRole.UserRole)
            if not result or not result.songs or result.error:
                continue
            if not result.service_date:
                logger.warning("Skipping %s – no service date", result.filename)
                continue

            if item.childCount() > 0:
                # Use individual child checkboxes (user may have overridden auto-classification)
                for j in range(item.childCount()):
                    child = item.child(j)
                    if child.checkState(0) != Qt.CheckState.Checked:
                        continue
                    cls: Optional[SongClassification] = child.data(0, _ROLE_CLS)
                    title = cls.title if cls else child.text(0).strip()
                    status = cls.status if cls else SongStatus.UNKNOWN
                    all_rows.append({
                        "file": result.filename,
                        "date": result.service_date,
                        "title": title,
                        "status": status,
                        "in_register": False,
                    })
            else:
                for song in result.songs:
                    all_rows.append({
                        "file": result.filename,
                        "date": result.service_date,
                        "title": song,
                        "status": SongStatus.UNKNOWN,
                        "in_register": False,
                    })

        if not all_rows:
            QMessageBox.information(
                self, self.windowTitle(), tr("dialog.import_pptx.no_songs")
            )
            return

        # Mark songs already present in the register
        try:
            registered = excel_service.get_registered_songs()
        except Exception as exc:
            logger.warning("Could not read registered songs: %s", exc)
            registered = set()

        for row in all_rows:
            row["in_register"] = row["title"].strip().lower() in registered

        # Collect UNKNOWN titles for the validation panel
        unknown_titles = [
            (r["title"], r["file"])
            for r in all_rows
            if r["status"] == SongStatus.UNKNOWN
        ]

        # Show validation dialog
        dlg = _ImportValidationDialog(all_rows, unknown_titles, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        confirmed_rows = dlg.result_rows()
        if not confirmed_rows:
            return

        # Group by (filename, service_date) and call import once per group
        groups: Dict = defaultdict(list)
        for row in confirmed_rows:
            groups[(row["file"], row["date"])].append(row["title"])

        processed = 0
        errors = []
        for (filename, svc_date), titles in groups.items():
            try:
                excel_service.import_service_songs(
                    service_date=svc_date,
                    song_titles=titles,
                )
                processed += 1
            except Exception as exc:
                logger.error(
                    "Import error for %s: %s", filename, exc, exc_info=True
                )
                errors.append(
                    tr("dialog.import_pptx.import_error").format(
                        filename=filename, error=str(exc)
                    )
                )

        msg = tr("dialog.import_pptx.import_complete").format(count=processed)
        if errors:
            msg += "\n\n" + "\n".join(errors)
            QMessageBox.warning(self, self.windowTitle(), msg)
        else:
            QMessageBox.information(self, self.windowTitle(), msg)

    def closeEvent(self, event):
        if self._worker and self._worker.isRunning():
            self._worker.cancel()
            self._worker.wait()
        super().closeEvent(event)
