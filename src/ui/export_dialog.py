"""Export dialog for exporting liturgy to various formats."""

import os
import subprocess
import sys
from typing import Optional

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QLineEdit,
    QCheckBox,
    QPushButton,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QGroupBox,
    QProgressBar,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QDesktopServices
from PyQt6.QtCore import QUrl

from ..models import Liturgy, Settings
from ..services import ExportService
from ..i18n import tr
from ..logging_config import get_logger

logger = get_logger("export_dialog")


class ExportWorker(QThread):
    """Worker thread for export operations."""

    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, export_service: ExportService, liturgy: Liturgy,
                 filename: str, export_pptx: bool, export_pdf: bool, export_txt: bool,
                 export_excel: bool = False, excel_path: Optional[str] = None):
        super().__init__()
        self.export_service = export_service
        self.liturgy = liturgy
        self.filename = filename
        self.export_pptx = export_pptx
        self.export_pdf = export_pdf
        self.export_txt = export_txt
        self.export_excel = export_excel
        self.excel_path = excel_path

    def run(self):
        try:
            results = {}
            base_filename = os.path.splitext(self.filename)[0]

            if self.export_pptx:
                self.progress.emit("Exporting PowerPoint...")
                results["pptx"] = self.export_service.export_pptx(
                    self.liturgy, f"{base_filename}.pptx"
                )

            if self.export_pdf:
                self.progress.emit("Exporting PDFs...")
                results["zip"] = self.export_service.export_pdf_zip(
                    self.liturgy, f"{base_filename}.zip"
                )

            if self.export_txt:
                self.progress.emit("Exporting links...")
                results["txt"] = self.export_service.export_links_txt(
                    self.liturgy, f"{base_filename}.txt"
                )

            if self.export_excel and self.excel_path:
                self.progress.emit("Updating Excel register...")
                results["excel"] = self.export_service.export_to_excel(
                    self.liturgy, self.excel_path
                )

            self.finished.emit(results)

        except Exception as e:
            self.error.emit(str(e))


class ExportDialog(QDialog):
    """Dialog for exporting liturgy to various formats."""

    def __init__(self, liturgy: Liturgy, export_service: ExportService,
                 settings: Optional[Settings] = None, base_path: str = ".", parent=None):
        super().__init__(parent)
        self.liturgy = liturgy
        self.export_service = export_service
        self.settings = settings
        self.base_path = base_path
        self._worker: Optional[ExportWorker] = None

        self._setup_ui()
        self._connect_signals()

    def _setup_ui(self) -> None:
        """Setup the dialog UI."""
        self.setWindowTitle(tr("dialog.export.title"))
        self.setMinimumSize(450, 300)
        self.resize(500, 350)

        layout = QVBoxLayout(self)

        # Filename input
        filename_layout = QFormLayout()
        self.filename_input = QLineEdit()
        self.filename_input.setText(self.export_service.get_default_filename(""))
        filename_layout.addRow(tr("dialog.export.filename"), self.filename_input)
        layout.addLayout(filename_layout)

        # Output folder info
        output_folder = self.export_service.get_output_folder()
        folder_label = QLabel(f"Output: {output_folder}")
        folder_label.setWordWrap(True)
        layout.addWidget(folder_label)

        # Export options
        options_group = QGroupBox("Export options")
        options_layout = QVBoxLayout(options_group)

        self.pptx_checkbox = QCheckBox(tr("dialog.export.include_pptx"))
        self.pptx_checkbox.setChecked(True)
        options_layout.addWidget(self.pptx_checkbox)

        self.pdf_checkbox = QCheckBox(tr("dialog.export.include_pdf"))
        self.pdf_checkbox.setChecked(True)
        options_layout.addWidget(self.pdf_checkbox)

        self.txt_checkbox = QCheckBox(tr("dialog.export.include_txt"))
        self.txt_checkbox.setChecked(True)
        options_layout.addWidget(self.txt_checkbox)

        # Excel export checkbox (only enabled if path is configured)
        self.excel_checkbox = QCheckBox(tr("dialog.export.include_excel"))
        self._excel_path = self._get_excel_path()
        if self._excel_path:
            self.excel_checkbox.setChecked(True)
            self.excel_checkbox.setEnabled(True)
        else:
            self.excel_checkbox.setChecked(False)
            self.excel_checkbox.setEnabled(False)
            self.excel_checkbox.setToolTip(tr("dialog.export.excel_not_configured"))
        options_layout.addWidget(self.excel_checkbox)

        # Separator line
        options_layout.addSpacing(10)

        self.open_after_checkbox = QCheckBox(tr("dialog.export.open_after"))
        self.open_after_checkbox.setChecked(True)
        options_layout.addWidget(self.open_after_checkbox)

        layout.addWidget(options_group)

        # Progress
        self.progress_label = QLabel()
        self.progress_label.setVisible(False)
        layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 0)  # Indeterminate
        layout.addWidget(self.progress_bar)

        # Results area
        self.results_label = QLabel()
        self.results_label.setWordWrap(True)
        self.results_label.setVisible(False)
        layout.addWidget(self.results_label)

        layout.addStretch()

        # Button box
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.export_button = self.button_box.button(QDialogButtonBox.StandardButton.Ok)
        self.export_button.setText(tr("button.export"))
        self.button_box.button(QDialogButtonBox.StandardButton.Cancel).setText(tr("button.cancel"))
        layout.addWidget(self.button_box)

    def _connect_signals(self) -> None:
        """Connect widget signals."""
        self.button_box.accepted.connect(self._on_export)
        self.button_box.rejected.connect(self.reject)

        # Update export button based on checkbox state
        self.pptx_checkbox.toggled.connect(self._update_export_button)
        self.pdf_checkbox.toggled.connect(self._update_export_button)
        self.txt_checkbox.toggled.connect(self._update_export_button)
        self.excel_checkbox.toggled.connect(self._update_export_button)

    def _get_excel_path(self) -> Optional[str]:
        """Get Excel register path from settings if configured and file exists."""
        if not self.settings:
            return None
        path = self.settings.get_excel_register_path(self.base_path)
        if path and os.path.exists(path):
            return path
        return None

    def _update_export_button(self) -> None:
        """Update export button enabled state."""
        any_selected = (
            self.pptx_checkbox.isChecked() or
            self.pdf_checkbox.isChecked() or
            self.txt_checkbox.isChecked() or
            self.excel_checkbox.isChecked()
        )
        self.export_button.setEnabled(any_selected)

    def _on_export(self) -> None:
        """Start the export process."""
        filename = self.filename_input.text().strip()
        if not filename:
            filename = self.export_service.get_default_filename("")

        # Check Excel export requirements
        logger.debug(f"Excel checkbox checked: {self.excel_checkbox.isChecked()}")
        logger.debug(f"Excel path: {self._excel_path}")
        logger.debug(f"Liturgy service_date: {self.liturgy.service_date}")
        logger.debug(f"Liturgy dienstleider: {self.liturgy.dienstleider}")

        export_excel = self.excel_checkbox.isChecked() and self._excel_path
        if export_excel:
            # Check if service_date is set
            if not self.liturgy.service_date:
                logger.warning("Excel export skipped: no service_date set")
                QMessageBox.warning(
                    self,
                    tr("dialog.export.title"),
                    tr("dialog.export.excel_warning_no_date")
                )
                export_excel = False
            # Check if dienstleider is set (warn but allow to continue)
            elif not self.liturgy.dienstleider:
                reply = QMessageBox.question(
                    self,
                    tr("dialog.export.title"),
                    tr("dialog.export.excel_warning_no_leader"),
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )
                if reply != QMessageBox.StandardButton.Yes:
                    logger.debug("Excel export skipped: user declined without dienstleider")
                    export_excel = False

        logger.debug(f"Final export_excel decision: {export_excel}")

        # Show progress
        self.progress_label.setVisible(True)
        self.progress_bar.setVisible(True)
        self.export_button.setEnabled(False)

        # Start worker
        self._worker = ExportWorker(
            self.export_service,
            self.liturgy,
            filename,
            self.pptx_checkbox.isChecked(),
            self.pdf_checkbox.isChecked(),
            self.txt_checkbox.isChecked(),
            export_excel,
            self._excel_path,
        )
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_export_finished)
        self._worker.error.connect(self._on_export_error)
        self._worker.start()

    def _on_progress(self, message: str) -> None:
        """Update progress display."""
        self.progress_label.setText(message)

    def _on_export_finished(self, results: dict) -> None:
        """Handle export completion."""
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)

        # Show results
        files_created = []
        for key, path in results.items():
            if path:
                if key == "excel":
                    files_created.append(f"{os.path.basename(path)} ({tr('dialog.export.excel_updated')})")
                else:
                    files_created.append(os.path.basename(path))

        if files_created:
            self.results_label.setText(
                f"{tr('dialog.export.success')}\n\n"
                f"{tr('dialog.export.files_created')}\n" +
                "\n".join(f"  - {f}" for f in files_created)
            )
            self.results_label.setVisible(True)

        # Open file if checkbox is checked
        if self.open_after_checkbox.isChecked():
            # Prefer opening PPTX, then ZIP, then TXT
            file_to_open = results.get("pptx") or results.get("zip") or results.get("txt")
            if file_to_open and os.path.exists(file_to_open):
                self._open_file(file_to_open)

        # Change button to Close
        self.export_button.setText(tr("button.close"))
        self.export_button.setEnabled(True)
        self.button_box.accepted.disconnect()
        self.button_box.accepted.connect(self.accept)

    def _on_export_error(self, error: str) -> None:
        """Handle export error."""
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)
        self.export_button.setEnabled(True)

        QMessageBox.critical(
            self,
            tr("dialog.export.title"),
            tr("error.export_failed", error=error)
        )

    def _open_file(self, file_path: str) -> None:
        """Open a file with the system's default application."""
        try:
            if sys.platform == "win32":
                os.startfile(file_path)
            elif sys.platform == "darwin":
                subprocess.run(["open", file_path], check=True)
            else:
                subprocess.run(["xdg-open", file_path], check=True)
        except Exception:
            # Fallback to Qt's desktop services
            QDesktopServices.openUrl(QUrl.fromLocalFile(file_path))
