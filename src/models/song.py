"""Song model representing a song in the Songs folder."""

import os
from dataclasses import dataclass, field
from typing import List, Optional, Dict


def _parse_properties(filepath: str) -> Dict[str, str]:
    """Parse Java properties file (key=value format)."""
    props = {}
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, _, value = line.partition('=')
                    props[key.strip()] = value.strip()
    except (IOError, UnicodeDecodeError):
        pass
    return props


@dataclass
class Song:
    """Represents a song with its associated files."""

    name: str
    folder_path: str
    relative_path: str  # Relative path from Songs root for display
    title: Optional[str] = None  # Title from song.properties, falls back to name
    pptx_path: Optional[str] = None
    pdf_path: Optional[str] = None
    youtube_links: List[str] = field(default_factory=list)

    @property
    def display_title(self) -> str:
        """Get display title (from song.properties or folder name)."""
        return self.title if self.title else self.name

    @property
    def has_pptx(self) -> bool:
        """Check if song has a PowerPoint file."""
        return self.pptx_path is not None and os.path.exists(self.pptx_path)

    @property
    def has_pdf(self) -> bool:
        """Check if song has a PDF file."""
        return self.pdf_path is not None and os.path.exists(self.pdf_path)

    @property
    def has_youtube(self) -> bool:
        """Check if song has YouTube links."""
        return len(self.youtube_links) > 0

    def load_youtube_links(self) -> List[str]:
        """Load YouTube links from youtube.txt in the song folder."""
        youtube_file = os.path.join(self.folder_path, "youtube.txt")
        if os.path.exists(youtube_file):
            try:
                with open(youtube_file, "r", encoding="utf-8") as f:
                    links = [
                        line.strip() for line in f.readlines() if line.strip()
                    ]
                    self.youtube_links = links
                    return links
            except IOError:
                pass
        return []

    def save_youtube_links(self, links: List[str]) -> None:
        """Save YouTube links to youtube.txt in the song folder."""
        self.youtube_links = links
        youtube_file = os.path.join(self.folder_path, "youtube.txt")
        with open(youtube_file, "w", encoding="utf-8") as f:
            for link in links:
                f.write(f"{link}\n")

    @classmethod
    def from_folder(cls, folder_path: str, songs_root: str) -> Optional["Song"]:
        """Create a Song from a folder path if it contains song files."""
        if not os.path.isdir(folder_path):
            return None

        # Get folder name as song name
        name = os.path.basename(folder_path)

        # Calculate relative path for display
        relative_path = os.path.relpath(folder_path, songs_root)

        # Find pptx file, pdf file, and song.properties
        pptx_path = None
        pdf_path = None
        title = None

        for file in os.listdir(folder_path):
            lower_file = file.lower()
            full_path = os.path.join(folder_path, file)

            if lower_file.endswith((".pptx", ".ppt")):
                pptx_path = full_path
            elif lower_file.endswith(".pdf"):
                pdf_path = full_path
            elif lower_file == "song.properties":
                # Parse song.properties for title
                props = _parse_properties(full_path)
                title = props.get("title")

        # Create song instance
        song = cls(
            name=name,
            folder_path=folder_path,
            relative_path=relative_path,
            title=title,
            pptx_path=pptx_path,
            pdf_path=pdf_path,
        )

        # Load YouTube links
        song.load_youtube_links()

        return song


@dataclass
class GenericItem:
    """Represents an item from the Generic (Algemeen) folder."""

    name: str
    pptx_path: str

    @property
    def display_name(self) -> str:
        """Get display name without extension."""
        return os.path.splitext(self.name)[0]


# Backwards compatibility alias
AlgemeenItem = GenericItem


@dataclass
class OfferingSlide:
    """Represents a single slide from the Offering (Collecte) PowerPoint."""

    index: int
    title: str


# Backwards compatibility alias
CollecteSlide = OfferingSlide
