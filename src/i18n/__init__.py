"""Internationalization package."""

from .translations import (
    TranslationManager,
    tr,
    set_language,
    get_language,
    on_language_changed,
)

__all__ = [
    "TranslationManager",
    "tr",
    "set_language",
    "get_language",
    "on_language_changed",
]
