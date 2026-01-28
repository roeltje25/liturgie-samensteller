"""Models package."""

from .settings import Settings
from .liturgy import (
    # V2 classes (new)
    SectionType,
    LiturgySlide,
    LiturgySection,
    generate_uuid,
    # V1 classes (still used)
    ItemType,
    LiturgyItem,
    SongLiturgyItem,
    GenericLiturgyItem,
    OfferingLiturgyItem,
    Liturgy,
    # Backwards compatibility aliases
    ElementType,
    LiturgyElement,
    SongElement,
    AlgemeenElement,
    CollecteElement,
)
from .song import (
    Song,
    # New names
    GenericItem,
    OfferingSlide,
    # Backwards compatibility aliases
    AlgemeenItem,
    CollecteSlide,
)

__all__ = [
    "Settings",
    # V2 classes (new)
    "SectionType",
    "LiturgySlide",
    "LiturgySection",
    "generate_uuid",
    # V1 classes (still used)
    "ItemType",
    "LiturgyItem",
    "SongLiturgyItem",
    "GenericLiturgyItem",
    "OfferingLiturgyItem",
    "Liturgy",
    "Song",
    "GenericItem",
    "OfferingSlide",
    # Backwards compatibility aliases
    "ElementType",
    "LiturgyElement",
    "SongElement",
    "AlgemeenElement",
    "CollecteElement",
    "AlgemeenItem",
    "CollecteSlide",
]
