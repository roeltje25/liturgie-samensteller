"""UI components package."""

from .main_window import MainWindow
from .liturgy_list import LiturgyListWidget
from .liturgy_tree import LiturgyTreeWidget
from .song_picker import SongPickerDialog
from .generic_picker import GenericPickerDialog, AlgemeenPickerDialog
from .offering_picker import OfferingPickerDialog, CollectePickerDialog
from .settings_dialog import SettingsDialog
from .export_dialog import ExportDialog
from .youtube_dialog import YouTubeDialog
from .theme_picker import ThemeSectionPicker
from .field_editor import SlideFieldEditor, BulkFieldEditor
from .section_editor import SectionEditorDialog
from .about_dialog import AboutDialog
from .splash_screen import create_splash, show_splash

__all__ = [
    "MainWindow",
    "LiturgyListWidget",
    "LiturgyTreeWidget",
    "SongPickerDialog",
    # New names
    "GenericPickerDialog",
    "OfferingPickerDialog",
    # Backwards compatibility aliases
    "AlgemeenPickerDialog",
    "CollectePickerDialog",
    "SettingsDialog",
    "ExportDialog",
    "YouTubeDialog",
    "ThemeSectionPicker",
    "SlideFieldEditor",
    "BulkFieldEditor",
    "SectionEditorDialog",
    "AboutDialog",
    "create_splash",
    "show_splash",
]
