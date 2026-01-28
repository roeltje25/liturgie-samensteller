# Installation Guide - Liturgie Samensteller

## For End Users

### Requirements
- Windows 10 or later
- Microsoft PowerPoint (required for slide merging and thumbnails)

### Installation Options

**Option 1: Single Executable (Simplest)**
1. Download `LiturgieSamensteller.exe` (70 MB)
2. Place it anywhere you like
3. Double-click to run
4. Note: First startup may take a few seconds as it extracts files

**Option 2: Installer (Recommended)**
1. Download `LiturgieSamensteller_Setup_x.x.x.exe`
2. Run the installer
3. Follow the installation wizard
4. Launch from Start Menu or Desktop shortcut

**Option 3: Portable (Folder)**
1. Download the ZIP file
2. Extract to a folder of your choice
3. Run `LiturgieSamensteller.exe` from the extracted folder

### First-time Setup
1. Go to **Extra > Instellingen** (Tools > Settings)
2. Configure your folder paths:
   - **Liederen map**: Folder containing your song PowerPoints
   - **Algemeen map**: Folder for general/liturgical items
   - **Uitvoer map**: Where exported files will be saved
   - **Collecte bestand**: Path to the offerings PowerPoint file

---

## For Developers

### Requirements
- Python 3.10 or later
- Microsoft PowerPoint (for full functionality)

### Setup Development Environment

```bash
# Clone the repository
git clone <repository-url>
cd powerpoint-mixer

# Create virtual environment (recommended)
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python run.py
```

### Building the Executable

```bash
# Install build dependencies
pip install -r requirements-dev.txt

# Build (creates dist\LiturgieSamensteller\)
build.bat

# Or manually:
pyinstaller build.spec --clean
```

The built application will be in `dist\LiturgieSamensteller\`.

### Creating a Distribution Package

After building, create a ZIP file of the `dist\LiturgieSamensteller` folder:

```bash
cd dist
powershell Compress-Archive -Path LiturgieSamensteller -DestinationPath LiturgieSamensteller-v1.0.0.zip
```

---

## Troubleshooting

### "PowerPoint not found" errors
- Ensure Microsoft PowerPoint is installed
- The app uses PowerPoint for merging slides and generating thumbnails

### Thumbnails not loading
- Thumbnails require PowerPoint to be installed
- First load may be slow as thumbnails are generated and cached

### Application won't start
- Check that all files from the distribution are present
- Try running from command line to see error messages:
  ```
  LiturgieSamensteller.exe
  ```
