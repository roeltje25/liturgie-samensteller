# Development Guidelines for Liturgie Samensteller

## Project Overview

Liturgie Samensteller (PowerPoint Mixer) is a PyQt6 application for building church service PowerPoint presentations. It manages songs, generic slides, offerings, and exports to various formats.

## Project Structure

```
src/
├── __init__.py          # Version info (__version__, __revision__, __build_date__)
├── main.py              # Application entry point
├── logging_config.py    # Logging setup
├── i18n/                # Internationalization
│   ├── __init__.py      # tr() function, set_language(), get_language()
│   ├── en.json          # English translations
│   └── nl.json          # Dutch translations (primary)
├── models/              # Data models
│   ├── liturgy.py       # Liturgy, LiturgySection, LiturgySlide, v1 item classes
│   └── settings.py      # Application settings
├── services/            # Business logic
│   ├── pptx_service.py  # PowerPoint operations (VBScript/python-pptx)
│   ├── export_service.py # Export orchestration
│   ├── excel_service.py # Excel register operations
│   ├── folder_scanner.py # Scan songs/generic folders
│   └── ...
└── ui/                  # PyQt6 UI components
    ├── main_window.py   # Main application window
    ├── liturgy_tree.py  # Tree widget for sections/slides
    ├── *_dialog.py      # Various dialogs
    └── *_picker.py      # Picker dialogs for songs, offerings, etc.
```

## Coding Conventions

### Naming
- **Classes**: `CamelCase` (e.g., `LiturgySection`, `ExportService`)
- **Functions/Methods**: `snake_case` (e.g., `get_slide_count`, `_on_add_song`)
- **Private methods**: Prefix with `_` (e.g., `_setup_ui`, `_connect_signals`)
- **Constants**: `UPPER_SNAKE_CASE` (e.g., `ITEM_TYPE_SECTION`)

### Imports
```python
# Standard library
import os
from typing import List, Optional, Dict

# Third-party
from PyQt6.QtWidgets import QDialog, QVBoxLayout
from PyQt6.QtCore import Qt, pyqtSignal

# Local imports (relative)
from ..models import Liturgy, LiturgySection
from ..services import PptxService
from ..i18n import tr
from ..logging_config import get_logger

logger = get_logger("module_name")
```

## Logging

Always use the centralized logging configuration:

```python
from ..logging_config import get_logger

logger = get_logger("module_name")

# Usage
logger.debug(f"Processing item: {item.title}")
logger.info(f"Exported to: {output_path}")
logger.warning(f"File not found: {path}")
logger.error(f"Export failed: {e}", exc_info=True)
```

Log file location: `{project_root}/liturgy_builder.log`

### What to Log
- **DEBUG**: Variable values, flow tracing, path conversions
- **INFO**: Major operations (save, load, export)
- **WARNING**: Non-critical issues (missing optional files)
- **ERROR**: Failures with `exc_info=True` for stack traces

## Internationalization (i18n)

All user-facing text must use the `tr()` function:

```python
from ..i18n import tr

# Simple key
label.setText(tr("button.save"))

# With parameters
message = tr("error.file_not_found", path=file_path)
title = tr("dialog.fields.slide_title", slide=slide.title)
```

### Translation Files
- Located in `src/i18n/en.json` and `src/i18n/nl.json`
- Keys are hierarchical: `"category.subcategory.key"`
- Parameters use `{param_name}` syntax

```json
{
  "button.save": "Opslaan",
  "error.file_not_found": "Bestand niet gevonden: {path}",
  "dialog.fields.slide_title": "Velden voor {slide}"
}
```

## Data Models

### Dataclasses with Serialization

```python
from dataclasses import dataclass, field
from typing import Optional, Dict, List

@dataclass
class LiturgySlide:
    id: str = field(default_factory=generate_uuid)
    title: str = ""
    source_path: Optional[str] = None
    fields: Dict[str, str] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            "id": self.id,
            "title": self.title,
            "source_path": self.source_path,
            "fields": self.fields,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "LiturgySlide":
        """Create instance from dictionary."""
        return cls(
            id=data.get("id", generate_uuid()),
            title=data.get("title", ""),
            source_path=data.get("source_path"),
            fields=data.get("fields", {}),
        )
```

### Path Handling
- Store paths as relative in JSON (for portability)
- Convert to absolute when loading using `base_path`
- Use `settings.get_*_path(base_path)` methods for configured paths

```python
# Getting configured paths
songs_path = self.settings.get_songs_path(self.base_path)
collecte_path = self.settings.get_collecte_path(self.base_path)

# Path conversion in Liturgy save/load
Liturgy._convert_paths_in_dict(data, base_path, to_relative=True)  # Save
Liturgy._convert_paths_in_dict(data, base_path, to_relative=False) # Load
```

## UI Patterns

### Dialog Structure

```python
class MyDialog(QDialog):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        self.data = data

        self._setup_ui()
        self._connect_signals()

    def _setup_ui(self) -> None:
        """Setup the dialog UI."""
        self.setWindowTitle(tr("dialog.my.title"))
        self.setMinimumSize(500, 400)

        layout = QVBoxLayout(self)
        # ... add widgets ...

        # Button box
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        layout.addWidget(self.button_box)

    def _connect_signals(self) -> None:
        """Connect widget signals."""
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
```

### Main Window Add/Insert Pattern

When adding items to the liturgy, insert at cursor position:

```python
def _get_insertion_index(self) -> int:
    """Get index where new sections should be inserted."""
    selected_idx = self.liturgy_tree.get_selected_section_index()
    if selected_idx >= 0:
        return selected_idx + 1  # Insert after selected
    return -1  # Add at end

def _insert_section_at_cursor(self, section: LiturgySection) -> int:
    """Insert section at cursor position, return new index."""
    insert_idx = self._get_insertion_index()
    if insert_idx >= 0:
        self.liturgy.insert_section(insert_idx, section)
        return insert_idx
    else:
        self.liturgy.add_section(section)
        return len(self.liturgy.sections) - 1
```

### Tree Widget Patterns

```python
# Store data in tree items
item.setData(0, Qt.ItemDataRole.UserRole, self.ITEM_TYPE_SECTION)
item.setData(0, Qt.ItemDataRole.UserRole + 1, section.id)

# Retrieve data
item_type = item.data(0, Qt.ItemDataRole.UserRole)
section_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
```

## Services

### Service Class Pattern

```python
class MyService:
    def __init__(self, settings: Settings, base_path: str):
        self.settings = settings
        self.base_path = base_path

    def do_operation(self, input_data) -> str:
        """
        Perform operation.

        Args:
            input_data: The input data.

        Returns:
            Path to result file.

        Raises:
            FileNotFoundError: If required file is missing.
        """
        logger.info(f"Starting operation with: {input_data}")
        # ... implementation ...
        return result_path
```

### PowerPoint Operations

- **Prefer VBScript** on Windows for full fidelity (themes, animations)
- **Fallback to python-pptx** when VBScript unavailable
- VBScript is generated dynamically and executed via `cscript.exe`

```python
def _merge_with_vbscript(self, slides_to_copy, section_info=None) -> str:
    """Use VBScript for native PowerPoint handling."""
    # Build VBScript
    vbs_lines = [...]

    # Write to temp file
    vbs_path = tempfile.mktemp(suffix='.vbs')
    with open(vbs_path, 'w', encoding='utf-8') as f:
        f.write('\r\n'.join(vbs_lines))

    # Execute
    result = subprocess.run(
        ['cscript', '//nologo', vbs_path],
        capture_output=True, text=True, timeout=120
    )
```

## Version Management

Update version in `src/__init__.py` for each change:

```python
__version__ = "1.1.29"
```

Version format: `MAJOR.MINOR.PATCH`
- Bump PATCH for bug fixes and small features
- Bump MINOR for significant features
- Bump MAJOR for breaking changes

## Error Handling

```python
try:
    result = self.service.do_operation(data)
except FileNotFoundError as e:
    logger.error(f"File not found: {e}")
    QMessageBox.warning(self, tr("error.title"), tr("error.file_not_found", path=str(e)))
except Exception as e:
    logger.error(f"Operation failed: {e}", exc_info=True)
    QMessageBox.critical(self, tr("error.title"), tr("error.operation_failed", error=str(e)))
```

## Testing Checklist

Before committing changes:

1. **Functionality**: Does the feature work as expected?
2. **Edge cases**: Empty data, missing files, special characters in names
3. **UI**: Are all strings translated? Do dialogs resize properly?
4. **Logging**: Is there enough logging to diagnose issues?
5. **Error handling**: Are errors caught and reported to user?

## Common Patterns

### Checking File Existence

```python
has_pptx = slide.source_path and os.path.exists(slide.source_path)
pptx_missing = not slide.is_stub and not has_pptx
```

### Signal/Slot Connection

```python
# In _connect_signals()
self.button.clicked.connect(self._on_button_clicked)
self.tree.itemSelectionChanged.connect(self._on_selection_changed)

# Disconnect if needed
self.button.clicked.disconnect()
```

### Background Operations

```python
class Worker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def run(self):
        try:
            self.progress.emit("Working...")
            result = do_work()
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))
```

## Notes

- **Dutch is primary language**: UI defaults to Dutch, English is secondary
- **Windows-first**: VBScript/COM automation assumes Windows + PowerPoint
- **Liturgy v2 format**: Uses sections containing slides (v1 used flat items list)
- **Excel register**: Optional feature for tracking song usage across services
