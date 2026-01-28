"""Liturgy and item models with hierarchical section support."""

import json
import os
import uuid
from dataclasses import dataclass, field, asdict
from datetime import date
from typing import List, Optional, Dict, Any, Union
from enum import Enum


class ItemType(Enum):
    """Types of liturgy items (v1 compatibility)."""

    SONG = "song"
    GENERIC = "generic"
    OFFERING = "offering"


class SectionType(Enum):
    """Types of liturgy sections (v2)."""

    REGULAR = "regular"
    SONG = "song"


# Backwards compatibility aliases
ElementType = ItemType
ElementType.ALGEMEEN = ItemType.GENERIC
ElementType.COLLECTE = ItemType.OFFERING


def generate_uuid() -> str:
    """Generate a new UUID string."""
    return str(uuid.uuid4())


@dataclass
class LiturgySlide:
    """A single slide within a liturgy section."""

    id: str = field(default_factory=generate_uuid)
    title: str = ""
    slide_index: int = 0
    source_path: Optional[str] = None
    fields: Dict[str, str] = field(default_factory=dict)
    is_stub: bool = False

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            "id": self.id,
            "title": self.title,
            "slide_index": self.slide_index,
            "source_path": self.source_path,
            "fields": self.fields,
            "is_stub": self.is_stub,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "LiturgySlide":
        """Create slide from dictionary."""
        return cls(
            id=data.get("id", generate_uuid()),
            title=data.get("title", ""),
            slide_index=data.get("slide_index", 0),
            source_path=data.get("source_path"),
            fields=data.get("fields", {}),
            is_stub=data.get("is_stub", False),
        )


@dataclass
class LiturgySection:
    """A section containing slides, optionally with song metadata."""

    id: str = field(default_factory=generate_uuid)
    name: str = ""
    section_type: SectionType = SectionType.REGULAR
    source_theme_path: Optional[str] = None
    slides: List[LiturgySlide] = field(default_factory=list)

    # Song-specific fields (only used when section_type == SONG)
    pdf_path: Optional[str] = None
    youtube_links: List[str] = field(default_factory=list)
    song_source_path: Optional[str] = None  # relative path in Songs folder

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        d = {
            "id": self.id,
            "name": self.name,
            "section_type": self.section_type.value,
            "source_theme_path": self.source_theme_path,
            "slides": [s.to_dict() for s in self.slides],
        }

        # Include song-specific fields only for song sections
        if self.section_type == SectionType.SONG:
            d["pdf_path"] = self.pdf_path
            d["youtube_links"] = self.youtube_links
            d["song_source_path"] = self.song_source_path

        return d

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "LiturgySection":
        """Create section from dictionary."""
        section_type_str = data.get("section_type", "regular")
        section_type = SectionType(section_type_str)

        section = cls(
            id=data.get("id", generate_uuid()),
            name=data.get("name", ""),
            section_type=section_type,
            source_theme_path=data.get("source_theme_path"),
        )

        # Load slides
        for slide_data in data.get("slides", []):
            section.slides.append(LiturgySlide.from_dict(slide_data))

        # Load song-specific fields
        if section_type == SectionType.SONG:
            section.pdf_path = data.get("pdf_path")
            section.youtube_links = data.get("youtube_links", [])
            section.song_source_path = data.get("song_source_path")

        return section

    @property
    def is_song(self) -> bool:
        """Check if this is a song section."""
        return self.section_type == SectionType.SONG

    @property
    def has_youtube(self) -> bool:
        """Check if section has YouTube links."""
        return bool(self.youtube_links)

    @property
    def has_pdf(self) -> bool:
        """Check if section has a PDF."""
        return bool(self.pdf_path)

    @property
    def has_pptx(self) -> bool:
        """Check if section has PowerPoint slides."""
        return any(slide.source_path for slide in self.slides)


# ============================================================================
# V1 Legacy Classes (for backwards compatibility)
# ============================================================================


@dataclass
class LiturgyItem:
    """Base class for liturgy items (v1 format)."""

    item_type: ItemType
    title: str
    source_path: Optional[str] = None
    is_stub: bool = False

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            "type": self.item_type.value,
            "title": self.title,
            "source_path": self.source_path,
            "is_stub": self.is_stub,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "LiturgyItem":
        """Create item from dictionary."""
        # Handle both old and new type values
        type_value = data["type"]
        if type_value == "algemeen":
            type_value = "generic"
        elif type_value == "collecte":
            type_value = "offering"

        item_type = ItemType(type_value)

        if item_type == ItemType.SONG:
            return SongLiturgyItem.from_dict(data)
        elif item_type == ItemType.OFFERING:
            return OfferingLiturgyItem.from_dict(data)
        else:
            return GenericLiturgyItem.from_dict(data)


# Backwards compatibility alias
LiturgyElement = LiturgyItem


@dataclass
class SongLiturgyItem(LiturgyItem):
    """A song item in the liturgy (v1 format)."""

    item_type: ItemType = field(default=ItemType.SONG, init=False)
    pptx_path: Optional[str] = None
    pdf_path: Optional[str] = None
    youtube_links: List[str] = field(default_factory=list)
    is_stub: bool = False

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        d = super().to_dict()
        d.update(
            {
                "pptx_path": self.pptx_path,
                "pdf_path": self.pdf_path,
                "youtube_links": self.youtube_links,
                "is_stub": self.is_stub,
            }
        )
        return d

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "SongLiturgyItem":
        """Create song item from dictionary."""
        return cls(
            title=data["title"],
            source_path=data.get("source_path"),
            pptx_path=data.get("pptx_path"),
            pdf_path=data.get("pdf_path"),
            youtube_links=data.get("youtube_links", []),
            is_stub=data.get("is_stub", False),
        )


# Backwards compatibility alias
SongElement = SongLiturgyItem


@dataclass
class GenericLiturgyItem(LiturgyItem):
    """A generic item from the Generic (Algemeen) folder (v1 format)."""

    item_type: ItemType = field(default=ItemType.GENERIC, init=False)
    pptx_path: Optional[str] = None
    is_stub: bool = False

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        d = super().to_dict()
        d["pptx_path"] = self.pptx_path
        d["is_stub"] = self.is_stub
        return d

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "GenericLiturgyItem":
        """Create generic item from dictionary."""
        return cls(
            title=data["title"],
            source_path=data.get("source_path"),
            pptx_path=data.get("pptx_path"),
            is_stub=data.get("is_stub", False),
        )


# Backwards compatibility alias
AlgemeenElement = GenericLiturgyItem


@dataclass
class OfferingLiturgyItem(LiturgyItem):
    """An offering item with a specific slide selection (v1 format)."""

    item_type: ItemType = field(default=ItemType.OFFERING, init=False)
    slide_index: int = 0
    slide_title: str = ""
    pptx_path: Optional[str] = None  # Optional: path to custom offerings file
    is_stub: bool = False

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        d = super().to_dict()
        d.update(
            {
                "slide_index": self.slide_index,
                "slide_title": self.slide_title,
                "pptx_path": self.pptx_path,
                "is_stub": self.is_stub,
            }
        )
        return d

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "OfferingLiturgyItem":
        """Create offering item from dictionary."""
        return cls(
            title=data["title"],
            source_path=data.get("source_path"),
            slide_index=data.get("slide_index", 0),
            slide_title=data.get("slide_title", ""),
            pptx_path=data.get("pptx_path"),
            is_stub=data.get("is_stub", False),
        )


# Backwards compatibility alias
CollecteElement = OfferingLiturgyItem


# ============================================================================
# Liturgy Class (supports both v1 items and v2 sections)
# ============================================================================


@dataclass
class Liturgy:
    """A complete liturgy with sections and slides (v2) or items (v1)."""

    name: str
    format_version: int = 2
    theme_source_path: Optional[str] = None
    sections: List[LiturgySection] = field(default_factory=list)
    created_date: str = field(default_factory=lambda: date.today().isoformat())

    # Service metadata for Excel export
    service_date: Optional[str] = None  # ISO date string (YYYY-MM-DD)
    dienstleider: Optional[str] = None  # Service leader name

    # V1 compatibility: items list (deprecated, use sections)
    _items: List[LiturgyItem] = field(default_factory=list, repr=False)

    @property
    def items(self) -> List[LiturgyItem]:
        """Get items list (v1 compatibility). Converts from sections if needed."""
        if self._items:
            return self._items
        # Convert sections to items for backwards compatibility
        return self._sections_to_items()

    @items.setter
    def items(self, value: List[LiturgyItem]) -> None:
        """Set items list (v1 compatibility)."""
        self._items = value

    # Backwards compatibility property
    @property
    def elements(self) -> List[LiturgyItem]:
        """Backwards compatibility alias for items."""
        return self.items

    @elements.setter
    def elements(self, value: List[LiturgyItem]) -> None:
        """Backwards compatibility alias for items."""
        self.items = value

    def _sections_to_items(self) -> List[LiturgyItem]:
        """Convert v2 sections to v1 items list."""
        items = []
        for section in self.sections:
            if section.is_song:
                # Song section becomes a SongLiturgyItem
                pptx_path = None
                if section.slides:
                    pptx_path = section.slides[0].source_path
                items.append(SongLiturgyItem(
                    title=section.name,
                    source_path=section.song_source_path,
                    pptx_path=pptx_path,
                    pdf_path=section.pdf_path,
                    youtube_links=section.youtube_links,
                    is_stub=section.slides[0].is_stub if section.slides else False,
                ))
            else:
                # Regular section - each slide could be a separate item
                for slide in section.slides:
                    items.append(GenericLiturgyItem(
                        title=slide.title or section.name,
                        source_path=slide.source_path,
                        pptx_path=slide.source_path,
                        is_stub=slide.is_stub,
                    ))
        return items

    def add_element(self, element: LiturgyItem) -> None:
        """Add an item to the liturgy. (Backwards compatibility alias)"""
        self.add_item(element)

    def add_item(self, item: LiturgyItem) -> None:
        """Add an item to the liturgy. Converts to section if using v2 format."""
        if self.format_version >= 2 or self.sections:
            # Convert to section and add to sections (v2 mode)
            section = self._item_to_section(item)
            self.sections.append(section)
        else:
            # Add to items list (v1 mode)
            self._items.append(item)

    def add_section(self, section: LiturgySection) -> None:
        """Add a section to the liturgy (v2 mode)."""
        self.sections.append(section)

    def insert_section(self, index: int, section: LiturgySection) -> None:
        """Insert a section at a specific position (v2 mode)."""
        index = max(0, min(index, len(self.sections)))
        self.sections.insert(index, section)

    def remove_element(self, index: int) -> None:
        """Remove an item at the given index. (Backwards compatibility alias)"""
        self.remove_item(index)

    def remove_item(self, index: int) -> None:
        """Remove an item at the given index (v1 mode)."""
        if 0 <= index < len(self._items):
            self._items.pop(index)

    def remove_section(self, index: int) -> None:
        """Remove a section at the given index (v2 mode)."""
        if 0 <= index < len(self.sections):
            self.sections.pop(index)

    def move_element(self, from_index: int, to_index: int) -> None:
        """Move an item from one position to another. (Backwards compatibility alias)"""
        self.move_item(from_index, to_index)

    def move_item(self, from_index: int, to_index: int) -> None:
        """Move an item from one position to another (v1 mode)."""
        if 0 <= from_index < len(self._items) and 0 <= to_index < len(self._items):
            item = self._items.pop(from_index)
            self._items.insert(to_index, item)

    def move_section(self, from_index: int, to_index: int) -> None:
        """Move a section from one position to another (v2 mode)."""
        if 0 <= from_index < len(self.sections) and 0 <= to_index < len(self.sections):
            section = self.sections.pop(from_index)
            self.sections.insert(to_index, section)

    def move_slide_within_section(
        self, section_index: int, from_slide_idx: int, to_slide_idx: int
    ) -> None:
        """Move a slide within a section."""
        if 0 <= section_index < len(self.sections):
            section = self.sections[section_index]
            if 0 <= from_slide_idx < len(section.slides) and 0 <= to_slide_idx < len(section.slides):
                slide = section.slides.pop(from_slide_idx)
                section.slides.insert(to_slide_idx, slide)

    def move_slide_to_section(
        self,
        from_section_idx: int,
        from_slide_idx: int,
        to_section_idx: int,
        to_slide_idx: int
    ) -> None:
        """Move a slide from one section to another."""
        if not (0 <= from_section_idx < len(self.sections)):
            return
        if not (0 <= to_section_idx < len(self.sections)):
            return

        from_section = self.sections[from_section_idx]
        to_section = self.sections[to_section_idx]

        if not (0 <= from_slide_idx < len(from_section.slides)):
            return

        # Remove from source section
        slide = from_section.slides.pop(from_slide_idx)

        # Clamp target index
        to_slide_idx = max(0, min(to_slide_idx, len(to_section.slides)))

        # Insert into target section
        to_section.slides.insert(to_slide_idx, slide)

    def get_section_by_id(self, section_id: str) -> Optional[LiturgySection]:
        """Find a section by its ID."""
        for section in self.sections:
            if section.id == section_id:
                return section
        return None

    def get_slide_by_id(self, slide_id: str) -> Optional[tuple]:
        """Find a slide by its ID. Returns (section, slide) or None."""
        for section in self.sections:
            for slide in section.slides:
                if slide.id == slide_id:
                    return (section, slide)
        return None

    def is_v2(self) -> bool:
        """Check if this liturgy uses v2 format (sections)."""
        return bool(self.sections) and not self._items

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        # Always save in v2 format
        if self.sections:
            d = {
                "format_version": 2,
                "name": self.name,
                "created_date": self.created_date,
                "theme_source_path": self.theme_source_path,
                "sections": [s.to_dict() for s in self.sections],
            }
            # Include service metadata if set
            if self.service_date:
                d["service_date"] = self.service_date
            if self.dienstleider:
                d["dienstleider"] = self.dienstleider
            return d
        else:
            # V1 format (items)
            return {
                "name": self.name,
                "created_date": self.created_date,
                "elements": [e.to_dict() for e in self._items],
            }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Liturgy":
        """Create liturgy from dictionary."""
        format_version = data.get("format_version", 1)

        liturgy = cls(
            name=data["name"],
            created_date=data.get("created_date", date.today().isoformat()),
            format_version=format_version,
        )

        # Load service metadata
        liturgy.service_date = data.get("service_date")
        liturgy.dienstleider = data.get("dienstleider")

        if format_version >= 2:
            # V2 format - load sections
            liturgy.theme_source_path = data.get("theme_source_path")
            for section_data in data.get("sections", []):
                liturgy.sections.append(LiturgySection.from_dict(section_data))
        else:
            # V1 format - load items
            for element_data in data.get("elements", []):
                liturgy._items.append(LiturgyItem.from_dict(element_data))

        return liturgy

    @classmethod
    def migrate_v1_to_v2(cls, v1_liturgy: "Liturgy") -> "Liturgy":
        """Migrate a v1 liturgy to v2 format."""
        v2_liturgy = cls(
            name=v1_liturgy.name,
            created_date=v1_liturgy.created_date,
            format_version=2,
        )

        for item in v1_liturgy._items:
            section = cls._item_to_section(item)
            v2_liturgy.sections.append(section)

        return v2_liturgy

    @staticmethod
    def _item_to_section(item: LiturgyItem) -> LiturgySection:
        """Convert a v1 item to a v2 section."""
        if isinstance(item, SongLiturgyItem):
            # Song becomes a song section
            section = LiturgySection(
                name=item.title,
                section_type=SectionType.SONG,
                pdf_path=item.pdf_path,
                youtube_links=item.youtube_links,
                song_source_path=item.source_path,
            )

            # Create slide entries for each slide in the song PPTX
            # For now, create a single slide entry representing all slides
            slide = LiturgySlide(
                title=item.title,
                slide_index=0,
                source_path=item.pptx_path,
                is_stub=item.is_stub,
            )
            section.slides.append(slide)

        elif isinstance(item, OfferingLiturgyItem):
            # Offering becomes a regular section with one slide
            section = LiturgySection(
                name=item.slide_title or item.title,
                section_type=SectionType.REGULAR,
            )

            slide = LiturgySlide(
                title=item.slide_title or item.title,
                slide_index=item.slide_index,
                source_path=item.pptx_path,
                is_stub=item.is_stub,
            )
            section.slides.append(slide)

        else:
            # Generic item becomes a regular section
            generic_item: GenericLiturgyItem = item
            section = LiturgySection(
                name=item.title,
                section_type=SectionType.REGULAR,
            )

            slide = LiturgySlide(
                title=item.title,
                slide_index=0,
                source_path=generic_item.pptx_path,
                is_stub=item.is_stub,
            )
            section.slides.append(slide)

        return section

    def save(self, path: str) -> None:
        """Save liturgy to JSON file."""
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)

    @classmethod
    def load(cls, path: str) -> "Liturgy":
        """Load liturgy from JSON file."""
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls.from_dict(data)

    @classmethod
    def load_with_migration(cls, path: str) -> tuple:
        """
        Load liturgy from JSON file, migrating if needed.
        Returns (liturgy, was_migrated) tuple.
        """
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        format_version = data.get("format_version", 1)
        liturgy = cls.from_dict(data)

        if format_version < 2:
            # Migrate to v2
            liturgy = cls.migrate_v1_to_v2(liturgy)
            return (liturgy, True)

        return (liturgy, False)
