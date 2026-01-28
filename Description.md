# PowerPoint Mixer - Liturgie Samensteller

## Overview
A bilingual (Dutch/English) desktop application for preparing church liturgy presentations.
The tool combines PowerPoint slides from various sources into a single presentation for worship services.

## Technology Stack
- **Language**: Python 3.13+
- **UI Framework**: PyQt6
- **PowerPoint Library**: python-pptx
- **YouTube Integration**: yt-dlp
- **Packaging**: PyInstaller (single .exe distribution)

## Folder Structure (Default)
```
./
├── Songs/                    # Hierarchical folder structure
│   ├── Category1/
│   │   ├── SongName/         # Leaf = song folder
│   │   │   ├── song.pptx     # Optional - song slides
│   │   │   ├── song.pdf      # Optional - sheet music
│   │   │   └── youtube.txt   # Optional - one URL per line
│   │   └── AnotherSong/
│   └── Category2/
│       └── ...
├── Algemeen/                  # General/generic slides (flat structure)
│   ├── Collecte.pptx         # Special: offerings with multiple slides
│   ├── StubTemplate.pptx     # Optional: template for missing songs
│   ├── Welkom.pptx
│   ├── Gebed.pptx
│   └── ...
├── Thema/                     # Theme templates
│   └── StandardService.pptx  # Reusable service structures
├── Vieringen/                 # Output folder
│   └── yyyymmdd_viering.pptx
└── LiederenRegister.xlsx     # Optional: song usage tracking
```

## Core Features

### Building a Liturgy
- **Section-based structure**: Organize content into logical sections with slides
- **Multiple content sources**:
  - Songs from hierarchical folder structure
  - General items (prayers, readings, etc.)
  - Offerings with smart title extraction
  - Theme templates for consistent service structure
  - Direct PowerPoint file import for custom content (sermons, announcements)
- **Drag-and-drop reordering** of sections and slides
- **Save/load** liturgy to/from JSON files

### Content Types

1. **Songs (Liederen)**
   - Searchable selection with fuzzy matching
   - Visual indicators for available resources (📺 YouTube, 📄 PDF)
   - Stub support for songs not yet in library
   - External file import option

2. **General Items (Algemeen)**
   - Select from pre-made slides (Welcome, Prayer, Blessing, etc.)
   - Imports all slides from selected file
   - Preserves source formatting

3. **Offerings (Collecte)**
   - Smart title extraction from slide content
   - Pattern matching for "Collecte: [purpose]" in text and notes
   - Fallback to slide notes when main content lacks title
   - Preview before selection

4. **Theme Templates**
   - Load entire service structures
   - Select specific sections or slides to add
   - Save current liturgy as reusable template

5. **Custom PowerPoint Files**
   - Import any .pptx directly as a new section
   - Useful for sermons, special announcements, guest content

### Field System
- **Dynamic field detection**: Finds placeholders like {DATE}, {LEADER}, etc.
- **Multiline text support**: Fields can contain multiple lines
- **Bulk editing**: Edit common fields across all slides at once
- **Visual indicators**: Warning icons for unfilled required fields

### YouTube Integration
- **Link storage**: `youtube.txt` in each song folder
- **Search functionality**: Find videos by song title
- **Audio preview**: Play button to preview videos before selecting
- **Link validation**: Check if links are still valid
- **Automatic save**: Links saved to song folder for future use

### Service Metadata
- **Service date**: Date picker with smart default (next Sunday)
- **Service leader**: Text field with autocomplete from Excel register
- Used for filename generation and Excel tracking

### Excel Song Register
- **Automatic tracking**: Records which songs are used on which date
- **Year columns**: Automatically adds new year columns as needed
- **Song list**: Maintains master list of all songs
- **Service leader tracking**: Records who led each service
- **Named ranges**: For formula references

### Localization
- Full bilingual support: Dutch (Nederlands) and English
- Seamless language toggle without restart
- All UI elements translated

### Export Options
1. **Combined PowerPoint** (.pptx)
   - All selected content merged into single presentation
   - Source formatting preserved
   - Configurable filename pattern

2. **Music PDF Archive** (.zip)
   - All available sheet music PDFs bundled
   - For musicians/worship team

3. **Links Overview** (.txt)
   - Song list with YouTube links
   - For sharing with team

4. **Excel Register Update**
   - Updates song usage tracking
   - Records service metadata

## Settings
Configurable via Settings dialog:
- Folder paths (Songs, General, Themes, Output)
- Collecte PowerPoint filename
- Stub template filename
- Output filename pattern (supports {date} placeholder)
- Excel register path
- Language preference

Settings persisted to `settings.json`.

---

## Recommendations for Church Success

### Quick Wins (Easy to Implement)

1. **Keyboard Shortcuts Reference**
   - Add Help > Keyboard Shortcuts menu
   - Common actions: Ctrl+N (new), Ctrl+S (save), Ctrl+E (export), Delete, Ctrl+Up/Down (reorder)
   - Print a quick reference card for operators

2. **Duplicate Previous Service**
   - "Open recent" with option to duplicate
   - Makes week-to-week preparation faster when services follow similar patterns

3. **Slide Thumbnails**
   - Show small preview images in the song/item pickers
   - Helps identify correct slides quickly

4. **Auto-backup**
   - Automatically save backup copies of liturgy files
   - Prevent accidental loss of work

5. **Confirmation on Export**
   - Show summary of what will be exported before proceeding
   - List any warnings (missing files, unfilled fields)

### Medium-Term Improvements

6. **Print Order of Service**
   - Generate printable bulletin/order of service
   - Include song titles, readings, and service elements
   - Configurable template

7. **Song Usage Statistics Dashboard**
   - View most/least used songs
   - See when songs were last used
   - Helps with song rotation planning

8. **Service Templates Library**
   - Pre-made templates for different service types
   - Normal Sunday, Communion, Christmas, Easter, Baptism, etc.
   - Quick start for special occasions

9. **QR Codes in Export**
   - Option to add QR codes linking to YouTube videos
   - Congregants can follow along on phones
   - Include in PDF/print output

10. **Undo/Redo Support**
    - Track changes within session
    - Recover from accidental deletions

### Long-Term Vision

11. **Cloud Sync**
    - Sync song library and liturgies across devices
    - Multiple volunteers can prepare services
    - Central repository of approved content

12. **Mobile Companion App**
    - Musicians can view upcoming songs and YouTube links
    - Practice mode with playlist generation
    - Push notifications when new liturgy is published

13. **Calendar Integration**
    - Link to church calendar
    - Auto-populate service dates
    - See liturgy assignments per date

14. **Collaborative Editing**
    - Multiple people can work on same liturgy
    - Comments and suggestions
    - Approval workflow

15. **Import from Other Systems**
    - Import song databases from other church software
    - CCLI integration for licensing compliance
    - OpenLP/EasyWorship compatibility

### Operational Best Practices

1. **Folder Organization**
   - Maintain consistent naming in Songs folder
   - Use categories that make sense for your church (by book, by theme, by language)
   - Keep PDFs alongside PPTXs for easy access

2. **Weekly Workflow**
   - Prepare liturgy early in the week
   - Share exported links file with musicians by Wednesday
   - Final review Friday, export Saturday

3. **Quality Control**
   - Periodically check YouTube link validity (Tools > Check YouTube links)
   - Review Excel register for song rotation
   - Update stub songs when real slides become available

4. **Backup Strategy**
   - Keep liturgy JSON files in version control or cloud storage
   - Backup entire Songs folder monthly
   - Test restores periodically

5. **Training**
   - Train 2-3 people on the software
   - Document any church-specific conventions
   - Keep this Description.md updated with local practices

### Technical Notes

- **Performance**: Large song libraries (1000+) may slow initial load; consider organizing into subcategories
- **PowerPoint Compatibility**: Best results with .pptx format; .ppt files may have formatting issues
- **YouTube**: Requires yt-dlp package; audio preview uses streaming (no download)
- **Excel**: Requires existing LiederenRegister.xlsx with proper structure; creates year columns automatically

---

## Version History

- **v2.0**: Section-based architecture, theme templates, field system, Excel integration
- **v1.0**: Basic liturgy building with songs, general items, and offerings
