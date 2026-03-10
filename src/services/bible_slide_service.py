"""Bible slide generation service.

Creates PowerPoint slides containing Bible text in multiple translations,
laid out in a grid (one column per translation).  Each slide shows a subset
of the requested verse range (determined by a text-fitting algorithm) and
includes a QR code linking to the full passage on YouVersion.
"""

import math
import os
import tempfile
from dataclasses import dataclass, field
from io import BytesIO
from typing import Dict, List, Optional, Tuple

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt, Emu

from ..logging_config import get_logger
from .bible_service import (
    BibleReference,
    BibleService,
    BibleTranslation,
    BibleVerse,
)

logger = get_logger("bible_slide_service")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

@dataclass
class BibleSlideConfig:
    """Configuration for Bible slide generation."""

    font_name: str = "Calibri"
    font_size: int = 12           # points
    slide_width: float = 10.0     # inches (standard 4:3)
    slide_height: float = 7.5     # inches
    show_verse_numbers: bool = True
    title_font_size: int = 18     # points
    header_font_size: int = 11    # points (translation name row)
    title_bold: bool = True
    bg_color: Optional[Tuple[int, int, int]] = None   # None = white
    title_color: Tuple[int, int, int] = (0, 0, 0)     # black
    text_color: Tuple[int, int, int] = (0, 0, 0)      # black
    header_color: Tuple[int, int, int] = (80, 80, 80) # dark grey
    qr_size: float = 1.0          # inches (square)
    column_gap: float = 0.1       # inches between columns
    margin_left: float = 0.3      # inches
    margin_right: float = 0.3     # inches
    margin_top: float = 0.25      # inches
    margin_bottom: float = 0.25   # inches


# Vertical layout constants (as fractions of slide height, resolved at runtime)
_TITLE_HEIGHT_FRAC = 0.11   # 11 % of slide height for title bar
_HEADER_HEIGHT_FRAC = 0.07  # 7 % for translation-name header row

# Safety factor: fill text boxes to at most this fraction of their height
# to avoid accidental overflow due to estimation inaccuracies.
_FILL_SAFETY = 0.85

# Approximate average character width as a fraction of font size (Calibri-like)
_AVG_CHAR_WIDTH_RATIO = 0.50


# ---------------------------------------------------------------------------
# BibleSlideService
# ---------------------------------------------------------------------------

class BibleSlideService:
    """Creates PowerPoint presentations from Bible text fetched via BibleService."""

    def __init__(self, bible_service: Optional[BibleService] = None) -> None:
        self._bible = bible_service or BibleService()

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def create_slides(
        self,
        reference_str: str,
        translation_ids: List[int],
        config: Optional[BibleSlideConfig] = None,
    ) -> str:
        """Generate a temporary PPTX file with Bible text slides.

        Each slide contains a grid of columns (one per translation) for a
        subset of the verse range.  A QR code linking to the full passage on
        YouVersion is placed in the bottom-right corner.

        Args:
            reference_str: Human reference string, e.g. "John 3:16-21".
            translation_ids: List of YouVersion version IDs (max 6).
            config: Slide configuration.  Defaults used if not provided.

        Returns:
            Absolute path to the generated temporary PPTX file.
            Caller is responsible for moving / deleting it.

        Raises:
            ValueError: On parse errors or no matching verses.
            requests.RequestException: On network failures.
        """
        from .bible_service import parse_reference

        if config is None:
            config = BibleSlideConfig()

        if not translation_ids:
            raise ValueError("At least one translation ID is required.")

        translation_ids = translation_ids[:6]  # cap at 6

        # 1. Parse reference
        ref = parse_reference(reference_str)
        logger.info("Creating Bible slides for: %s, translations: %s", ref.display_str, translation_ids)

        # 2. Fetch verses for each translation
        translations_info: List[BibleTranslation] = self._bible.get_builtin_translations()
        id_to_trans: Dict[int, BibleTranslation] = {t.id: t for t in translations_info}

        per_translation: List[Tuple[BibleTranslation, List[BibleVerse]]] = []
        for vid in translation_ids:
            trans = id_to_trans.get(vid, BibleTranslation(
                id=vid, abbreviation=str(vid), name=str(vid), language="", language_name=""
            ))
            try:
                verses = self._bible.get_verses(ref, vid)
                if not verses:
                    logger.warning("No verses returned for version %s", vid)
                    verses = [BibleVerse(verse_num=0, text="(No text available)")]
            except Exception as exc:
                logger.error("Failed to fetch verses for version %s: %s", vid, exc, exc_info=True)
                verses = [BibleVerse(verse_num=0, text=f"(Error: {exc})")]
            per_translation.append((trans, verses))

        # 3. Determine QR URL (use first translation)
        qr_url = self._bible.get_youversion_url(ref, translation_ids[0])

        # 4. Build slides
        return self._build_presentation(ref, per_translation, qr_url, config)

    # ------------------------------------------------------------------
    # Private – presentation builder
    # ------------------------------------------------------------------

    def _build_presentation(
        self,
        ref: BibleReference,
        per_translation: List[Tuple[BibleTranslation, List[BibleVerse]]],
        qr_url: str,
        config: BibleSlideConfig,
    ) -> str:
        """Build the PPTX and return the temp file path."""

        prs = Presentation()
        prs.slide_width = Inches(config.slide_width)
        prs.slide_height = Inches(config.slide_height)

        # Pre-compute layout geometry
        geom = _SlideGeometry(config)

        # Normalise verse lists: all translations must cover the same verse numbers.
        # Use union of all verse numbers so every column shows the same rows.
        all_verse_nums = sorted(
            {v.verse_num for _, verses in per_translation for v in verses}
        )
        if not all_verse_nums:
            raise ValueError("No verse data retrieved for any translation.")

        # Build verse_text lookup per (translation_index, verse_num)
        verse_lookup: Dict[Tuple[int, int], str] = {}
        for t_idx, (_, verses) in enumerate(per_translation):
            for v in verses:
                verse_lookup[(t_idx, v.verse_num)] = v.text

        n_cols = len(per_translation)

        # Determine how many verse rows fit on a single slide
        verse_groups = self._group_verses_for_slides(
            all_verse_nums=all_verse_nums,
            verse_lookup=verse_lookup,
            n_translations=n_cols,
            geom=geom,
            config=config,
        )

        # Generate QR code image once
        qr_image: Optional[BytesIO] = _generate_qr(qr_url)

        # Blank slide layout (index 6 in most built-in themes is blank)
        blank_layout = prs.slide_layouts[6]

        slide_label = ref.display_str

        for group_idx, group_verse_nums in enumerate(verse_groups):
            slide = prs.slides.add_slide(blank_layout)

            # Optional background fill
            if config.bg_color:
                _fill_slide_bg(slide, config.bg_color, prs)

            # --- Title bar ---
            title_text = f"{slide_label}"
            if len(verse_groups) > 1:
                # Show verse sub-range on this slide
                v_first = group_verse_nums[0]
                v_last = group_verse_nums[-1]
                title_text += f"  ({ref.chapter}:{v_first}–{v_last})"
            _add_text_box(
                slide,
                x=Inches(config.margin_left),
                y=Inches(config.margin_top),
                w=Inches(config.slide_width - config.margin_left - config.margin_right),
                h=Inches(geom.title_height_in),
                text=title_text,
                font_name=config.font_name,
                font_size=config.title_font_size,
                bold=config.title_bold,
                color=config.title_color,
                align=PP_ALIGN.LEFT,
            )

            # --- Translation header row ---
            for col_idx, (trans, _) in enumerate(per_translation):
                col_x = Inches(geom.column_x(col_idx, n_cols))
                col_w = Inches(geom.column_width(n_cols))
                header_y = Inches(
                    config.margin_top + geom.title_height_in + geom.header_gap_in
                )
                header_h = Inches(geom.header_height_in)
                header_text = f"{trans.abbreviation} – {trans.name}"
                _add_text_box(
                    slide,
                    x=col_x,
                    y=header_y,
                    w=col_w,
                    h=header_h,
                    text=header_text,
                    font_name=config.font_name,
                    font_size=config.header_font_size,
                    bold=False,
                    color=config.header_color,
                    align=PP_ALIGN.LEFT,
                )

            # --- Verse text columns ---
            text_top = (
                config.margin_top
                + geom.title_height_in
                + geom.header_gap_in
                + geom.header_height_in
                + geom.text_top_gap_in
            )

            for col_idx, (trans, _) in enumerate(per_translation):
                col_x = Inches(geom.column_x(col_idx, n_cols))
                col_w = Inches(geom.column_width(n_cols))
                text_h = Inches(geom.text_area_height_in(config))

                # Build text for this column / verse group
                lines = []
                for vnum in group_verse_nums:
                    verse_text = verse_lookup.get((col_idx, vnum), "")
                    if config.show_verse_numbers and vnum != 0:
                        lines.append(f"{vnum} {verse_text}")
                    else:
                        lines.append(verse_text)
                full_text = "\n".join(lines)

                _add_text_box(
                    slide,
                    x=col_x,
                    y=Inches(text_top),
                    w=col_w,
                    h=text_h,
                    text=full_text,
                    font_name=config.font_name,
                    font_size=config.font_size,
                    bold=False,
                    color=config.text_color,
                    align=PP_ALIGN.LEFT,
                    word_wrap=True,
                )

            # --- QR code ---
            if qr_image:
                qr_x = Inches(
                    config.slide_width - config.margin_right - config.qr_size
                )
                qr_y = Inches(
                    config.slide_height - config.margin_bottom - config.qr_size
                )
                qr_image.seek(0)
                slide.shapes.add_picture(
                    qr_image,
                    left=qr_x,
                    top=qr_y,
                    width=Inches(config.qr_size),
                    height=Inches(config.qr_size),
                )

        # Save to temp file
        fd, tmp_path = tempfile.mkstemp(suffix=".pptx", prefix="bible_slides_")
        os.close(fd)
        prs.save(tmp_path)
        logger.info("Bible slides saved to: %s (%d slides)", tmp_path, len(verse_groups))
        return tmp_path

    # ------------------------------------------------------------------
    # Private – text fitting
    # ------------------------------------------------------------------

    def _group_verses_for_slides(
        self,
        all_verse_nums: List[int],
        verse_lookup: Dict[Tuple[int, int], str],
        n_translations: int,
        geom: "_SlideGeometry",
        config: BibleSlideConfig,
    ) -> List[List[int]]:
        """Group verse numbers into slides so text fits within each cell.

        Uses a character-count approximation to estimate the rendered height
        of each verse block.

        Returns a list of groups; each group is a list of verse numbers
        to put on one slide.
        """
        col_width_in = geom.column_width(n_translations)
        col_width_pt = col_width_in * 72.0
        text_area_h_pt = geom.text_area_height_in(config) * 72.0

        # Characters per line (approximation)
        avg_char_w_pt = config.font_size * _AVG_CHAR_WIDTH_RATIO
        chars_per_line = max(1, col_width_pt / avg_char_w_pt)

        # Height per text line (including line spacing)
        line_height_pt = config.font_size * 1.35
        # Extra spacing between paragraphs (verses)
        para_spacing_pt = config.font_size * 0.4

        available_h_pt = text_area_h_pt * _FILL_SAFETY

        groups: List[List[int]] = []
        current_group: List[int] = []
        current_h_pt: float = 0.0

        for vnum in all_verse_nums:
            # Find the maximum height needed across all translations for this verse
            verse_h_pt = 0.0
            for t_idx in range(n_translations):
                text = verse_lookup.get((t_idx, vnum), "")
                if config.show_verse_numbers and vnum != 0:
                    text = f"{vnum} {text}"
                h = _estimate_text_height_pt(
                    text, chars_per_line, line_height_pt, para_spacing_pt
                )
                verse_h_pt = max(verse_h_pt, h)

            if current_group and (current_h_pt + verse_h_pt > available_h_pt):
                # Flush current group, start a new one
                groups.append(current_group)
                current_group = [vnum]
                current_h_pt = verse_h_pt
            else:
                current_group.append(vnum)
                current_h_pt += verse_h_pt

        if current_group:
            groups.append(current_group)

        if not groups:
            groups = [[]]

        return groups


# ---------------------------------------------------------------------------
# Slide geometry helper
# ---------------------------------------------------------------------------

class _SlideGeometry:
    """Pre-computes layout metrics for a given config."""

    def __init__(self, config: BibleSlideConfig) -> None:
        self.c = config
        self.title_height_in: float = config.slide_height * _TITLE_HEIGHT_FRAC
        self.header_gap_in: float = 0.05
        self.header_height_in: float = config.slide_height * _HEADER_HEIGHT_FRAC
        self.text_top_gap_in: float = 0.05

    def column_width(self, n_cols: int) -> float:
        """Width of one column (inches)."""
        c = self.c
        total_gap = c.column_gap * (n_cols - 1) if n_cols > 1 else 0
        usable_w = c.slide_width - c.margin_left - c.margin_right - total_gap
        return max(0.5, usable_w / n_cols)

    def column_x(self, col_idx: int, n_cols: int) -> float:
        """Left edge of a column (inches)."""
        c = self.c
        return c.margin_left + col_idx * (self.column_width(n_cols) + c.column_gap)

    def text_area_height_in(self, config: BibleSlideConfig) -> float:
        """Height available for verse text, below the header row (inches)."""
        used = (
            config.margin_top
            + self.title_height_in
            + self.header_gap_in
            + self.header_height_in
            + self.text_top_gap_in
            + config.margin_bottom
            + config.qr_size  # leave room for QR code in last column area
        )
        return max(0.5, config.slide_height - used)


# ---------------------------------------------------------------------------
# Text height estimation
# ---------------------------------------------------------------------------

def _estimate_text_height_pt(
    text: str,
    chars_per_line: float,
    line_height_pt: float,
    para_spacing_pt: float,
) -> float:
    """Estimate the height (in points) a block of text will occupy."""
    if not text:
        return 0.0
    paragraphs = text.split("\n")
    total_h = 0.0
    for para in paragraphs:
        n_chars = max(1, len(para))
        n_lines = math.ceil(n_chars / chars_per_line)
        total_h += n_lines * line_height_pt + para_spacing_pt
    return total_h


# ---------------------------------------------------------------------------
# python-pptx helpers
# ---------------------------------------------------------------------------

def _add_text_box(
    slide,
    x: Emu,
    y: Emu,
    w: Emu,
    h: Emu,
    text: str,
    font_name: str,
    font_size: int,
    bold: bool,
    color: Tuple[int, int, int],
    align: PP_ALIGN = PP_ALIGN.LEFT,
    word_wrap: bool = True,
) -> None:
    """Add a text box to a slide."""
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap

    # Split text into paragraphs
    paragraphs = text.split("\n")
    for p_idx, para_text in enumerate(paragraphs):
        if p_idx == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = para_text
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(*color)


def _fill_slide_bg(slide, color: Tuple[int, int, int], prs: Presentation) -> None:
    """Fill slide background with a solid color."""
    from pptx.util import Pt
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*color)


# ---------------------------------------------------------------------------
# QR code generation
# ---------------------------------------------------------------------------

def _generate_qr(url: str) -> Optional[BytesIO]:
    """Generate a QR code PNG image in memory.  Returns None if qrcode not available."""
    try:
        import qrcode  # type: ignore
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=10,
            border=2,
        )
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        buf = BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return buf
    except ImportError:
        logger.warning("qrcode library not available; QR codes will not be generated.")
        return None
    except Exception as exc:
        logger.error("Failed to generate QR code: %s", exc, exc_info=True)
        return None
