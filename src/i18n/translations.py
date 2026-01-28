"""Translation system for bilingual support."""

import json
import os
from typing import Dict, Optional, Callable, List

from PyQt6.QtCore import QObject, pyqtSignal


class TranslationManager(QObject):
    """Manages translations and language switching."""

    language_changed = pyqtSignal(str)

    _instance: Optional["TranslationManager"] = None

    def __init__(self):
        super().__init__()
        self._current_language = "nl"
        self._translations: Dict[str, Dict[str, str]] = {}
        self._load_translations()

    @classmethod
    def instance(cls) -> "TranslationManager":
        """Get or create singleton instance."""
        if cls._instance is None:
            cls._instance = TranslationManager()
        return cls._instance

    def _load_translations(self) -> None:
        """Load all translation files."""
        i18n_dir = os.path.dirname(__file__)

        for lang in ["nl", "en"]:
            file_path = os.path.join(i18n_dir, f"{lang}.json")
            if os.path.exists(file_path):
                try:
                    with open(file_path, "r", encoding="utf-8") as f:
                        self._translations[lang] = json.load(f)
                except (json.JSONDecodeError, IOError):
                    self._translations[lang] = {}
            else:
                self._translations[lang] = {}

    def get_current_language(self) -> str:
        """Get current language code."""
        return self._current_language

    def set_language(self, lang: str) -> None:
        """Set the current language and emit signal."""
        if lang in self._translations and lang != self._current_language:
            self._current_language = lang
            self.language_changed.emit(lang)

    def get_available_languages(self) -> List[str]:
        """Get list of available language codes."""
        return list(self._translations.keys())

    def tr(self, key: str, **kwargs) -> str:
        """
        Get translated string for key.
        Supports format placeholders: tr("hello", name="World") -> "Hello, World!"
        """
        translations = self._translations.get(self._current_language, {})
        text = translations.get(key, key)

        # Apply format arguments if any
        if kwargs:
            try:
                text = text.format(**kwargs)
            except KeyError:
                pass

        return text


# Global shortcut function
def tr(key: str, **kwargs) -> str:
    """Translate a key using the current language."""
    return TranslationManager.instance().tr(key, **kwargs)


def set_language(lang: str) -> None:
    """Set the current language."""
    TranslationManager.instance().set_language(lang)


def get_language() -> str:
    """Get the current language."""
    return TranslationManager.instance().get_current_language()


def on_language_changed(callback: Callable[[str], None]) -> None:
    """Register a callback for language changes."""
    TranslationManager.instance().language_changed.connect(callback)
