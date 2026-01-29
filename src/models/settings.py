"""Settings model for the application."""

import json
import os
import sys
from dataclasses import dataclass, field, asdict
from typing import Optional

from ..logging_config import get_logger

logger = get_logger("settings")

# Application name for config directory
APP_NAME = "LiturgieSamensteller"


def get_config_dir() -> str:
    """Get the appropriate config directory for the current platform.

    Windows: %APPDATA%/LiturgieSamensteller
    macOS: ~/Library/Application Support/LiturgieSamensteller
    Linux: ~/.config/LiturgieSamensteller
    """
    if sys.platform == "win32":
        # Windows: use APPDATA
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
    elif sys.platform == "darwin":
        # macOS: use Application Support
        base = os.path.expanduser("~/Library/Application Support")
    else:
        # Linux/Unix: use XDG_CONFIG_HOME or ~/.config
        base = os.environ.get("XDG_CONFIG_HOME", os.path.expanduser("~/.config"))

    config_dir = os.path.join(base, APP_NAME)

    # Create directory if it doesn't exist
    if not os.path.exists(config_dir):
        try:
            os.makedirs(config_dir, exist_ok=True)
            logger.debug(f"Created config directory: {config_dir}")
        except Exception as e:
            logger.error(f"Failed to create config directory {config_dir}: {e}")
            # Fall back to current directory
            return "."

    return config_dir


def get_settings_path() -> str:
    """Get the full path to the settings file."""
    return os.path.join(get_config_dir(), "settings.json")


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
    excel_register_path: str = "./LiederenRegister.xlsx"

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
    def load(cls, path: Optional[str] = None) -> "Settings":
        """Load settings from JSON file.

        Args:
            path: Path to settings file. If None, uses default config location.
        """
        if path is None:
            path = get_settings_path()

        if os.path.exists(path):
            try:
                logger.debug(f"Loading settings from: {path}")
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                # Filter out unknown keys to handle version differences
                valid_fields = {f.name for f in cls.__dataclass_fields__.values()}
                filtered_data = {k: v for k, v in data.items() if k in valid_fields}
                if len(filtered_data) != len(data):
                    ignored = set(data.keys()) - valid_fields
                    logger.debug(f"Ignored unknown settings keys: {ignored}")
                return cls(**filtered_data)
            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON in settings file {path}: {e}")
            except TypeError as e:
                logger.error(f"Invalid settings data in {path}: {e}")
            except Exception as e:
                logger.error(f"Unexpected error loading settings from {path}: {e}", exc_info=True)
        else:
            logger.debug(f"Settings file not found, using defaults: {path}")
        return cls()

    def save(self, path: Optional[str] = None) -> None:
        """Save settings to JSON file.

        Args:
            path: Path to settings file. If None, uses default config location.
        """
        if path is None:
            path = get_settings_path()

        # Ensure directory exists
        dir_path = os.path.dirname(path)
        if dir_path and not os.path.exists(dir_path):
            os.makedirs(dir_path, exist_ok=True)

        with open(path, "w", encoding="utf-8") as f:
            json.dump(asdict(self), f, indent=2, ensure_ascii=False)
        logger.debug(f"Settings saved to: {path}")

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
