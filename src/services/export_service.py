"""Service for exporting liturgy to various formats."""

import os
import zipfile
from datetime import date
from typing import List, Optional

from ..models import Settings, Liturgy, SongLiturgyItem, ItemType
from .pptx_service import PptxService
from .excel_service import ExcelService


class ExportService:
    """Service for exporting liturgy outputs."""

    def __init__(self, settings: Settings, base_path: str = "."):
        self.settings = settings
        self.base_path = base_path
        self.pptx_service = PptxService(settings, base_path)

    def get_default_filename(self, extension: str = ".pptx", service_date: Optional[str] = None) -> str:
        """Generate default filename based on pattern.

        Args:
            extension: File extension to use.
            service_date: ISO date string (YYYY-MM-DD) for the filename. Falls back to today.
        """
        pattern = self.settings.output_pattern

        if service_date:
            d = date.fromisoformat(service_date)
        else:
            d = date.today()

        filename = pattern.format(
            date=d.strftime("%Y%m%d"),
            year=d.strftime("%Y"),
            month=d.strftime("%m"),
            day=d.strftime("%d"),
        )

        # Ensure correct extension
        base, _ = os.path.splitext(filename)
        return base + extension

    def get_output_folder(self) -> str:
        """Get the output folder path, creating it if necessary."""
        output_path = self.settings.get_output_path(self.base_path)
        os.makedirs(output_path, exist_ok=True)
        return output_path

    def export_pptx(self, liturgy: Liturgy, filename: Optional[str] = None) -> str:
        """
        Export liturgy to a PowerPoint file.
        Returns the path to the created file.
        """
        if filename is None:
            filename = self.get_default_filename(".pptx")

        output_folder = self.get_output_folder()
        output_path = os.path.join(output_folder, filename)

        # Merge all slides (returns path to temp file)
        # Use v2 merge if liturgy has sections, v1 merge for legacy items
        if liturgy.sections:
            temp_path = self.pptx_service.merge_liturgy_v2(liturgy)
        else:
            temp_path = self.pptx_service.merge_liturgy(liturgy)

        # Save to final location
        self.pptx_service.save_presentation(temp_path, output_path)

        return output_path

    def export_pdf_zip(self, liturgy: Liturgy, filename: Optional[str] = None) -> str:
        """
        Export all available PDFs to a zip file.
        Returns the path to the created zip file.
        """
        if filename is None:
            filename = self.get_default_filename(".zip")

        output_folder = self.get_output_folder()
        output_path = os.path.join(output_folder, filename)

        # Collect all PDFs from song items
        pdf_files = []
        for item in liturgy.items:
            if item.item_type == ItemType.SONG:
                song_item: SongLiturgyItem = item
                if song_item.pdf_path and os.path.exists(song_item.pdf_path):
                    pdf_files.append((song_item.pdf_path, song_item.title))

        # Create zip file
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for pdf_path, title in pdf_files:
                # Use sanitized title as filename in zip
                safe_title = self._sanitize_filename(title)
                ext = os.path.splitext(pdf_path)[1]
                arcname = f"{safe_title}{ext}"

                # Handle duplicates
                counter = 1
                original_arcname = arcname
                while arcname in zf.namelist():
                    base, ext = os.path.splitext(original_arcname)
                    arcname = f"{base}_{counter}{ext}"
                    counter += 1

                zf.write(pdf_path, arcname)

        return output_path

    def export_links_txt(self, liturgy: Liturgy, filename: Optional[str] = None) -> str:
        """
        Export liturgy with YouTube links to a text file.
        Returns the path to the created file.
        """
        if filename is None:
            filename = self.get_default_filename(".txt")

        output_folder = self.get_output_folder()
        output_path = os.path.join(output_folder, filename)

        lines = []
        lines.append(f"Liturgie: {liturgy.name}")
        lines.append(f"Datum: {liturgy.created_date}")
        lines.append("")
        lines.append("=" * 50)
        lines.append("")

        for i, item in enumerate(liturgy.items, 1):
            lines.append(f"{i}. {item.title}")

            if item.item_type == ItemType.SONG:
                song_item: SongLiturgyItem = item
                if song_item.youtube_links:
                    for link in song_item.youtube_links:
                        lines.append(f"   YouTube: {link}")
                else:
                    lines.append("   (geen YouTube link)")

                if song_item.pdf_path:
                    lines.append(f"   PDF: {os.path.basename(song_item.pdf_path)}")

            lines.append("")

        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

        return output_path

    def export_all(
        self, liturgy: Liturgy, base_filename: Optional[str] = None
    ) -> dict:
        """
        Export all formats (pptx, zip, txt).
        Returns dict with paths to all created files.
        """
        if base_filename is None:
            base_filename = self.get_default_filename("")
        else:
            base_filename, _ = os.path.splitext(base_filename)

        results = {}

        # Export PowerPoint
        results["pptx"] = self.export_pptx(liturgy, f"{base_filename}.pptx")

        # Export PDF zip (only if there are PDFs)
        has_pdfs = any(
            item.item_type == ItemType.SONG
            and item.pdf_path
            and os.path.exists(item.pdf_path)
            for item in liturgy.items
        )
        if has_pdfs:
            results["zip"] = self.export_pdf_zip(liturgy, f"{base_filename}.zip")

        # Export links text
        results["txt"] = self.export_links_txt(liturgy, f"{base_filename}.txt")

        return results

    def _sanitize_filename(self, filename: str) -> str:
        """Remove or replace invalid filename characters."""
        invalid_chars = '<>:"/\\|?*'
        result = filename
        for char in invalid_chars:
            result = result.replace(char, "_")
        return result.strip()

    def export_to_excel(self, liturgy: Liturgy, excel_path: str) -> str:
        """Export liturgy to Excel registration file.

        Args:
            liturgy: The liturgy to export.
            excel_path: Path to the Excel file.

        Returns:
            Path to the updated Excel file.
        """
        excel_service = ExcelService(excel_path)
        return excel_service.export_liturgy(liturgy)

    def get_excel_dienstleiders(self, excel_path: str) -> List[str]:
        """Get list of unique Dienstleiders for autocomplete.

        Args:
            excel_path: Path to the Excel file.

        Returns:
            List of unique dienstleider names.
        """
        if not excel_path or not os.path.exists(excel_path):
            return []
        excel_service = ExcelService(excel_path)
        return excel_service.get_dienstleiders()
