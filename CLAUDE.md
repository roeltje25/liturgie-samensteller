# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Liturgie Samensteller (PowerPoint Mixer) — a Windows PyQt6 desktop app for building church service PowerPoint presentations. Combines songs, offerings, and general liturgical items into a merged PPTX. Dutch is the primary language; English is secondary.

## Running & Building

```bash
# Run the app
python run.py

# Build executable (interactive menu: single exe, folder, or Nuitka)
build.bat

# Build single exe directly
python -m PyInstaller build_onefile.spec --clean --noconfirm

# Clean build artifacts
clean.bat
```

**No automated tests exist.** Testing is manual.

## Architecture

**Entry point:** `run.py` → `src/main.py:main()` (splash screen, load settings, launch MainWindow)

**Three-layer architecture:**
- **Models** (`src/models/`) — dataclasses with `to_dict()`/`from_dict()` JSON serialization
- **Services** (`src/services/`) — business logic, no UI dependencies
- **UI** (`src/ui/`) — PyQt6 widgets, dialogs, and pickers

### Key Services
- `pptx_service.py` — Merges slides. **Prefers VBScript** (via `cscript.exe`) for full fidelity; falls back to python-pptx. Generates VBS dynamically, writes to temp file, executes with subprocess.
- `export_service.py` — Orchestrates export (PPTX merge, PDF archive, links overview, Excel update)
- `excel_service.py` — Tracks song usage in LiederenRegister.xlsx (openpyxl)
- `folder_scanner.py` — Scans Songs/Algemeen/Themes folders with caching
- `youtube_service.py` — yt-dlp subprocess for YouTube search/validation

### Data Model (Liturgy v2)
Hierarchical: `Liturgy` → `LiturgySection[]` → `LiturgySlide[]`

- `LiturgySection` has a `type` (REGULAR or SONG), optional song metadata (PDF path, YouTube links)
- `LiturgySlide` has `title`, `source_path`, `fields` dict, `is_stub` flag
- Paths stored as relative in JSON, converted to absolute on load via `base_path`
- Each object has a UUID `id` field
- Legacy v1 classes (`LiturgyItem`, `SongLiturgyItem`, etc.) still exist but v2 is primary

### UI Patterns
- Dialogs follow `__init__` → `_setup_ui()` → `_connect_signals()` pattern
- Tree widget stores item type in `UserRole` and IDs in `UserRole+1`
- Background work uses `QThread` with `finished`/`error`/`progress` signals
- All user-facing strings use `tr("key.path")` from `src/i18n/`

## Conventions

- **i18n:** All user-facing text must use `tr()`. Translation files: `src/i18n/{nl,en}.json`
- **Logging:** Use `get_logger("module_name")` from `src/logging_config`. Log to `liturgy_builder.log` (5MB rotating). Use `exc_info=True` on errors.
- **Imports:** stdlib → third-party → relative local (`from ..models import ...`)
- **Version:** Bump `__version__` in `src/__init__.py` (MAJOR.MINOR.PATCH). Bump PATCH for fixes/small features.
- **Paths:** Store relative in serialized data, resolve to absolute at runtime via settings helpers (`settings.get_songs_path()`, etc.)

## Dependencies

Runtime: PyQt6, python-pptx, Pillow, yt-dlp, requests, lxml, openpyxl, pywin32
Build: PyInstaller (primary), optionally Nuitka

## Key File Locations

- Settings persist to `AppData/LiturgieSamensteller/settings.json`
- User data folder structure: `Liederen/` (songs), `Algemeen/` (general items + Collecte.pptx), `Themas/` (templates), `Vieringen/` (output)
- Songs live in subfolders containing `.pptx`, optional `.pdf`, optional `youtube.txt`
