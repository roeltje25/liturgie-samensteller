# Liturgie Samensteller

A Windows application for building church service (liturgy) PowerPoint presentations. Combines songs, offerings, and general liturgical elements into a single presentation.

## Features

- **Song Library**: Browse and select songs from your collection
- **Offering Slides**: Select offering destinations with slide previews
- **Generic Items**: Add general liturgical elements (prayers, readings, etc.)
- **Theme Templates**: Save and reuse common service structures
- **YouTube Integration**: Search and link YouTube videos for songs
- **PDF Support**: Link PDF sheet music to songs
- **Excel Register**: Track song usage in an Excel spreadsheet
- **Export Options**: Generate PowerPoint, PDF bundle, and links overview
- **Bilingual**: Dutch and English interface

## Requirements

- Windows 10 or later
- Microsoft PowerPoint (required for slide merging and thumbnails)

## Installation

See [INSTALL.md](INSTALL.md) for detailed installation instructions.

### Quick Start (End Users)

1. Download `LiturgieSamensteller.exe` from releases
2. Run the executable
3. On first run, select your base folder containing liturgy files

### Development Setup

```bash
# Clone the repository
git clone <repository-url>
cd powerpoint-mixer

# Create virtual environment
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python run.py
```

## Building

```bash
# Install build dependencies
pip install -r requirements-dev.txt

# Build executable (interactive - choose single exe or folder)
build.bat

# Or build single executable directly
python -m PyInstaller build_onefile.spec --clean --noconfirm

# Clean build artifacts
clean.bat
```

## Project Structure

```
powerpoint-mixer/
├── src/
│   ├── i18n/          # Translation files (nl.json, en.json)
│   ├── models/        # Data models (Settings, Liturgy, etc.)
│   ├── services/      # Business logic (PptxService, FolderScanner, etc.)
│   └── ui/            # PyQt6 user interface
├── build.bat          # Build script
├── build.spec         # PyInstaller config (folder mode)
├── build_onefile.spec # PyInstaller config (single exe)
├── clean.bat          # Clean build artifacts
├── requirements.txt   # Runtime dependencies
├── requirements-dev.txt # Development dependencies
└── run.py             # Application entry point
```

## Folder Structure (User Data)

The application expects this folder structure in your base folder:

```
Base Folder/
├── Liederen/          # Song folders (each with .pptx, .pdf, youtube.txt)
├── Algemeen/          # General items (.pptx files)
│   ├── Collecte.pptx  # Offering slides
│   └── StubTemplate.pptx  # Template for stub slides
├── Themas/            # Theme templates (.liturgy files)
└── Vieringen/         # Output folder for generated files
```

## License

[Add your license here]
