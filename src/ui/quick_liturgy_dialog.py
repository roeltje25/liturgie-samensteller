"""Quick Fill Songs dialog for bulk song assignment to liturgy sections."""

from typing import List, Optional, Dict, Tuple

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QSplitter,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QComboBox,
    QWidget,
    QMessageBox,
    QHeaderView,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

from ..models import Song, LiturgySection, LiturgySlide
from ..services.song_matcher import find_best_matches
from ..i18n import tr
from ..logging_config import get_logger

logger = get_logger("quick_liturgy_dialog")

# Score thresholds
CONFIRMED_THRESHOLD = 0.75
AMBIGUOUS_THRESHOLD = 0.40


def _compute_default_assignments(
    song_count: int, song_sections: List[LiturgySection]
) -> List[Optional[str]]:
    """Compute default section ID for each song index.

    Returns a list of section IDs (or None for "new section") of length song_count.

    Assignment algorithm:
    - n = number of *empty* SONG sections, S = song_count
    - n=0, sections exist  → all songs go to the first SONG section
    - n=0, no sections      → new section per song
    - S ≤ n                 → one song per empty section
    - S > n                 → section[0] gets songs 0..S-(n-1), sections 1..n-1 get one each
    """
    if song_count == 0:
        return []

    empty_sections = [s for s in song_sections if not s.slides]
    n = len(empty_sections)
    S = song_count
    assignments: List[Optional[str]] = []

    if n == 0:
        if song_sections:
            # Append all to first SONG section
            first_id = song_sections[0].id
            assignments = [first_id] * S
        else:
            # Create new section per song
            assignments = [None] * S
    elif S <= n:
        # One song per empty section
        for i in range(S):
            assignments.append(empty_sections[i].id)
    else:
        # S > n: overflow goes to first empty section
        overflow = S - (n - 1)
        for _ in range(overflow):
            assignments.append(empty_sections[0].id)
        for i in range(1, n):
            assignments.append(empty_sections[i].id)

    return assignments


class QuickLiturgyDialog(QDialog):
    """Dialog for pasting a song list and assigning songs to liturgy sections in bulk."""

    def __init__(
        self,
        songs: List[Song],
        liturgy_sections: List[LiturgySection],
        parent=None,
    ):
        super().__init__(parent)
        self.songs = songs
        self.liturgy_sections = liturgy_sections  # All SONG sections in liturgy order
        self._rows: List[Dict] = []  # Internal per-row state

        self._setup_ui()
        self._connect_signals()

    # ------------------------------------------------------------------
    # UI setup
    # ------------------------------------------------------------------

    def _setup_ui(self) -> None:
        self.setWindowTitle(tr("dialog.quick_fill.title"))
        self.setMinimumSize(820, 480)
        self.resize(950, 560)

        layout = QVBoxLayout(self)

        # Splitter: input on left, results table on right
        splitter = QSplitter(Qt.Orientation.Horizontal)
        layout.addWidget(splitter)

        # --- Left: paste area ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 4, 0)

        input_label = QLabel(tr("dialog.quick_fill.input_label"))
        left_layout.addWidget(input_label)

        self.input_text = QPlainTextEdit()
        left_layout.addWidget(self.input_text)

        self.match_btn = QPushButton(tr("dialog.quick_fill.match_btn"))
        left_layout.addWidget(self.match_btn)

        splitter.addWidget(left_widget)

        # --- Right: results table ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(4, 0, 0, 0)

        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels([
            tr("dialog.quick_fill.col_typed"),
            tr("dialog.quick_fill.col_matched"),
            tr("dialog.quick_fill.col_section"),
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        right_layout.addWidget(self.table)

        splitter.addWidget(right_widget)
        splitter.setSizes([300, 550])

        # --- Bottom buttons ---
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.add_btn = QPushButton(tr("dialog.quick_fill.add_btn"))
        self.add_btn.setEnabled(False)
        btn_layout.addWidget(self.add_btn)

        cancel_btn = QPushButton(tr("button.cancel"))
        btn_layout.addWidget(cancel_btn)

        layout.addLayout(btn_layout)

        self.add_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)

    def _connect_signals(self) -> None:
        self.match_btn.clicked.connect(self._on_match)

    # ------------------------------------------------------------------
    # Matching
    # ------------------------------------------------------------------

    def _on_match(self) -> None:
        """Parse the pasted text and run fuzzy matching for each line."""
        text = self.input_text.toPlainText()
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

        if not lines:
            QMessageBox.information(
                self,
                tr("dialog.quick_fill.title"),
                tr("dialog.quick_fill.no_input"),
            )
            return

        assignments = _compute_default_assignments(len(lines), self.liturgy_sections)

        self._rows = []
        for i, line in enumerate(lines):
            top_matches = find_best_matches(line, self.songs, limit=3)

            if top_matches and top_matches[0][1] >= CONFIRMED_THRESHOLD:
                state = "confirmed"
                best_song = top_matches[0][0]
            elif top_matches and top_matches[0][1] >= AMBIGUOUS_THRESHOLD:
                state = "ambiguous"
                best_song = top_matches[0][0]
            else:
                state = "stub"
                best_song = None

            target_id = assignments[i] if i < len(assignments) else None

            self._rows.append({
                "typed": line,
                "state": state,
                "song": best_song,
                "top_matches": top_matches,
                "target_section_id": target_id,
            })

        self._populate_table()
        self.add_btn.setEnabled(True)

    # ------------------------------------------------------------------
    # Table population
    # ------------------------------------------------------------------

    def _section_options(self) -> List[Tuple[str, Optional[str]]]:
        """Return (display_name, section_id) pairs for the section dropdown."""
        options = []
        for sec in self.liturgy_sections:
            label = sec.name or f"Sectie {sec.id[:6]}"
            options.append((label, sec.id))
        options.append((tr("dialog.quick_fill.new_section_option"), None))
        return options

    def _populate_table(self) -> None:
        """Rebuild the table rows from self._rows."""
        self.table.setRowCount(0)
        section_options = self._section_options()

        for row_idx, row in enumerate(self._rows):
            self.table.insertRow(row_idx)
            state = row["state"]

            # --- Column 0: Typed text ---
            typed_item = QTableWidgetItem(row["typed"])
            typed_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            if state == "stub":
                typed_item.setBackground(QColor("#ffebee"))
            elif state == "ambiguous":
                typed_item.setBackground(QColor("#fff8e1"))
            self.table.setItem(row_idx, 0, typed_item)

            # --- Column 1: Match result ---
            if state == "confirmed":
                matched_item = QTableWidgetItem("✓ " + row["song"].display_title)
                matched_item.setForeground(QColor("#2e7d32"))
                matched_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                self.table.setItem(row_idx, 1, matched_item)

            elif state == "ambiguous":
                combo = QComboBox()
                for song, score in row["top_matches"]:
                    combo.addItem(f"{song.display_title} ({int(score * 100)}%)", song)
                combo.currentIndexChanged.connect(
                    lambda idx, r=row_idx: self._on_match_combo_changed(r, idx)
                )
                self.table.setCellWidget(row_idx, 1, combo)

            else:
                # Stub row: label + "Find…" button
                stub_widget = QWidget()
                stub_layout = QHBoxLayout(stub_widget)
                stub_layout.setContentsMargins(4, 2, 4, 2)
                stub_label = QLabel("✗ " + tr("dialog.quick_fill.stub_label"))
                stub_label.setStyleSheet("color: #c62828;")
                find_btn = QPushButton(tr("dialog.quick_fill.find_btn"))
                find_btn.setMaximumHeight(26)
                find_btn.clicked.connect(
                    lambda checked, r=row_idx: self._on_find_song(r)
                )
                stub_layout.addWidget(stub_label)
                stub_layout.addWidget(find_btn)
                stub_layout.addStretch()
                self.table.setCellWidget(row_idx, 1, stub_widget)

            # --- Column 2: Section dropdown ---
            sec_combo = QComboBox()
            for name, sid in section_options:
                sec_combo.addItem(name, sid)

            # Select the default
            target_id = row["target_section_id"]
            for idx, (_, sid) in enumerate(section_options):
                if sid == target_id:
                    sec_combo.setCurrentIndex(idx)
                    break

            sec_combo.currentIndexChanged.connect(
                lambda idx, r=row_idx: self._on_section_combo_changed(r, idx)
            )
            self.table.setCellWidget(row_idx, 2, sec_combo)

        self.table.resizeRowsToContents()

    # ------------------------------------------------------------------
    # Signal handlers
    # ------------------------------------------------------------------

    def _on_match_combo_changed(self, row_idx: int, combo_idx: int) -> None:
        combo = self.table.cellWidget(row_idx, 1)
        if combo:
            self._rows[row_idx]["song"] = combo.itemData(combo_idx)

    def _on_section_combo_changed(self, row_idx: int, combo_idx: int) -> None:
        sec_combo = self.table.cellWidget(row_idx, 2)
        if sec_combo:
            self._rows[row_idx]["target_section_id"] = sec_combo.itemData(combo_idx)

    def _on_find_song(self, row_idx: int) -> None:
        """Open SongPickerDialog so the user can manually resolve a stub row."""
        from .song_picker import SongPickerDialog
        dlg = SongPickerDialog(self.songs, parent=self)
        if dlg.exec():
            item = dlg.get_selected_item()
            if item and item.pptx_path:
                song = next(
                    (s for s in self.songs if s.pptx_path == item.pptx_path),
                    None,
                )
                if song:
                    self._rows[row_idx]["song"] = song
                    self._rows[row_idx]["state"] = "confirmed"
                    self._populate_table()

    # ------------------------------------------------------------------
    # Result
    # ------------------------------------------------------------------

    def result_rows(self) -> List[Dict]:
        """Return list of {slide: LiturgySlide, target_section_id: str|None} dicts."""
        results = []
        for row in self._rows:
            song: Optional[Song] = row.get("song")
            is_stub = row["state"] == "stub" or song is None
            title = song.display_title if song else row["typed"]

            slide = LiturgySlide(
                title=title,
                slide_index=0,
                source_path=song.pptx_path if song else None,
                is_stub=is_stub,
                pdf_path=song.pdf_path if song else None,
                youtube_links=list(song.youtube_links) if song else [],
            )
            results.append({
                "slide": slide,
                "target_section_id": row["target_section_id"],
            })
        return results
