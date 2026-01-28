"""Settings model for the application."""

import json
import os
from dataclasses import dataclass, field, asdict
from typing import Optional


@dataclass
class Settings:
    """Application settings."""

    # Base folder - all relative paths are relative to this
    base_folder: str = ""  # Empty means not configured (first run)

    # Subfolders (relative to base_folder, or absolute)
    songs_folder: str = "./Liederen"
    algemeen_folder: str = "./Algemeen"
    output_folder: str = "./Vieringen"
    themes_folder: str = "./Themas"
    collecte_filename: str = "Collecte.pptx"
    stub_template_filename: str = "StubTemplate.pptx"
    output_pattern: str = "{date}_viering-generated.pptx"
    language: str = "nl"

    # Excel registration
    excel_register_path: str = ""  # Path to LiederenRegister.xlsx

    # Window state
    window_width: int = 1200
    window_height: int = 800

    def is_first_run(self) -> bool:
        """Check if this is the first run (base folder not configured)."""
        return not self.base_folder

    def get_effective_base_path(self, fallback_path: str = ".") -> str:
        """Get the effective base path (base_folder if set, otherwise fallback)."""
        if self.base_folder and os.path.isdir(self.base_folder):
            return self.base_folder
        return fallback_path

    @classmethod
    def load(cls, path: str = "settings.json") -> "Settings":
        """Load settings from JSON file."""
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                return cls(**data)
            except (json.JSONDecodeError, TypeError):
                pass
        return cls()

    def save(self, path: str = "settings.json") -> None:
        """Save settings to JSON file."""
        with open(path, "w", encoding="utf-8") as f:
            json.dump(asdict(self), f, indent=2, ensure_ascii=False)

    def _resolve_base(self, fallback_path: str = ".") -> str:
        """Resolve the base path to use for relative paths."""
        return self.get_effective_base_path(fallback_path)

    def get_songs_path(self, base_path: str = ".") -> str:
        """Get absolute path to songs folder."""
        if os.path.isabs(self.songs_folder):
            return self.songs_folder
        return os.path.normpath(os.path.join(self._resolve_base(base_path), self.songs_folder))

    def get_algemeen_path(self, base_path: str = ".") -> str:
        """Get absolute path to algemeen folder."""
        if os.path.isabs(self.algemeen_folder):
            return self.algemeen_folder
        return os.path.normpath(os.path.join(self._resolve_base(base_path), self.algemeen_folder))

    def get_output_path(self, base_path: str = ".") -> str:
        """Get absolute path to output folder."""
        if os.path.isabs(self.output_folder):
            return self.output_folder
        return os.path.normpath(os.path.join(self._resolve_base(base_path), self.output_folder))

    def get_collecte_path(self, base_path: str = ".") -> str:
        """Get absolute path to collecte PowerPoint."""
        algemeen = self.get_algemeen_path(base_path)
        return os.path.join(algemeen, self.collecte_filename)

    def get_stub_template_path(self, base_path: str = ".") -> Optional[str]:
        """Get absolute path to stub template if it exists."""
        algemeen = self.get_algemeen_path(base_path)
        path = os.path.join(algemeen, self.stub_template_filename)
        if os.path.exists(path):
            return path
        return None

    def get_themes_path(self, base_path: str = ".") -> str:
        """Get absolute path to themes folder."""
        if os.path.isabs(self.themes_folder):
            return self.themes_folder
        return os.path.normpath(os.path.join(self._resolve_base(base_path), self.themes_folder))

    def get_excel_register_path(self, base_path: str = ".") -> Optional[str]:
        """Get absolute path to Excel register file, or None if not set."""
        if not self.excel_register_path:
            return None
        if os.path.isabs(self.excel_register_path):
            return self.excel_register_path
        return os.path.normpath(os.path.join(self._resolve_base(base_path), self.excel_register_path))
