"""Bible text fetching service using the YouVersion (bible.com) API.

Uses the informal YouVersion web API to fetch Bible texts in multiple translations.
The API is not officially documented but is widely used by community projects.
"""

import math
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple
from urllib.parse import quote

import requests

from ..logging_config import get_logger

logger = get_logger("bible_service")

# YouVersion API base URLs
YOUVERSION_API_BASE = "https://nodejs.bible.com/api/bible"
YOUVERSION_VERSIONS_URL = "https://www.bible.com/api/bible/versions"
YOUVERSION_SHARE_BASE = "https://www.bible.com/bible"

# Request timeout in seconds
REQUEST_TIMEOUT = 15

# Default request headers to mimic a browser
_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.bible.com/",
}

# ---------------------------------------------------------------------------
# Built-in translation catalog
# These are YouVersion version IDs for commonly used Bible translations.
# The IDs correspond to the numeric version identifiers used by bible.com.
# ---------------------------------------------------------------------------
BUILTIN_TRANSLATIONS = [
    # English
    {"id": 1,    "abbreviation": "KJV",       "name": "King James Version",              "language": "en", "language_name": "English"},
    {"id": 59,   "abbreviation": "ESV",       "name": "English Standard Version",        "language": "en", "language_name": "English"},
    {"id": 111,  "abbreviation": "NIV",       "name": "New International Version",       "language": "en", "language_name": "English"},
    {"id": 116,  "abbreviation": "NLT",       "name": "New Living Translation",          "language": "en", "language_name": "English"},
    {"id": 100,  "abbreviation": "NASB",      "name": "New American Standard Bible",     "language": "en", "language_name": "English"},
    {"id": 206,  "abbreviation": "WEB",       "name": "World English Bible",             "language": "en", "language_name": "English"},
    {"id": 2692, "abbreviation": "NIV11",     "name": "New International Version 2011",  "language": "en", "language_name": "English"},
    {"id": 97,   "abbreviation": "MSG",       "name": "The Message",                     "language": "en", "language_name": "English"},
    {"id": 37,   "abbreviation": "HCSB",      "name": "Holman Christian Standard Bible", "language": "en", "language_name": "English"},
    # Dutch
    {"id": 48,   "abbreviation": "NBG51",     "name": "NBG-vertaling 1951",              "language": "nl", "language_name": "Nederlands"},
    {"id": 328,  "abbreviation": "NBV",       "name": "Nieuwe Bijbelvertaling",           "language": "nl", "language_name": "Nederlands"},
    {"id": 1816, "abbreviation": "HSV",       "name": "Herziene Statenvertaling",        "language": "nl", "language_name": "Nederlands"},
    {"id": 524,  "abbreviation": "BGT",       "name": "Bijbel in Gewone Taal",           "language": "nl", "language_name": "Nederlands"},
    {"id": 278,  "abbreviation": "SV",        "name": "Statenvertaling",                 "language": "nl", "language_name": "Nederlands"},
    # German
    {"id": 51,   "abbreviation": "LUTH1545",  "name": "Luther Bibel 1545",               "language": "de", "language_name": "Deutsch"},
    {"id": 157,  "abbreviation": "ELB",       "name": "Elberfelder Bibel",               "language": "de", "language_name": "Deutsch"},
    {"id": 1588, "abbreviation": "NGU-DE",    "name": "Neue Genfer Übersetzung",         "language": "de", "language_name": "Deutsch"},
    {"id": 70,   "abbreviation": "SCH2000",   "name": "Schlachter 2000",                 "language": "de", "language_name": "Deutsch"},
    # French
    {"id": 93,   "abbreviation": "LSG",       "name": "Louis Segond 1910",               "language": "fr", "language_name": "Français"},
    {"id": 1462, "abbreviation": "BDS",       "name": "Bible du Semeur",                 "language": "fr", "language_name": "Français"},
    {"id": 4096, "abbreviation": "NBS",       "name": "Nouvelle Bible Segond",           "language": "fr", "language_name": "Français"},
    # Spanish
    {"id": 128,  "abbreviation": "RVR60",     "name": "Reina-Valera 1960",               "language": "es", "language_name": "Español"},
    {"id": 149,  "abbreviation": "RVA",       "name": "Reina-Valera Actualizada",        "language": "es", "language_name": "Español"},
    {"id": 503,  "abbreviation": "NVI",       "name": "Nueva Versión Internacional",     "language": "es", "language_name": "Español"},
    # Portuguese
    {"id": 212,  "abbreviation": "ARC",       "name": "Almeida Revisada Corrigida",      "language": "pt", "language_name": "Português"},
    {"id": 1608, "abbreviation": "NVT",       "name": "Nova Versão Transformadora",      "language": "pt", "language_name": "Português"},
    # Italian
    {"id": 121,  "abbreviation": "CEI",       "name": "Conferenza Episcopale Italiana",  "language": "it", "language_name": "Italiano"},
    {"id": 431,  "abbreviation": "NR2006",    "name": "Nuova Riveduta 2006",             "language": "it", "language_name": "Italiano"},
    # Russian
    {"id": 400,  "abbreviation": "SYNO",      "name": "Синодальный перевод",             "language": "ru", "language_name": "Русский"},
    # Chinese Simplified
    {"id": 1268, "abbreviation": "CNVS",      "name": "Chinese New Version (Simplified)","language": "zh", "language_name": "中文"},
    # Greek
    {"id": 2097, "abbreviation": "SBLG",      "name": "SBL Greek New Testament",        "language": "el", "language_name": "Ελληνικά"},
    # Hebrew
    {"id": 2310, "abbreviation": "WLC",       "name": "Westminster Leningrad Codex",     "language": "he", "language_name": "עברית"},
    # Arabic
    {"id": 3,    "abbreviation": "AVDB",      "name": "Arabic Bible",                   "language": "ar", "language_name": "العربية"},
]


# ---------------------------------------------------------------------------
# USFM book code mappings
# ---------------------------------------------------------------------------
# Maps lowercased common names / abbreviations to USFM codes
BOOK_NAME_TO_USFM: Dict[str, str] = {
    # Genesis
    "gen": "GEN", "genesis": "GEN", "gene": "GEN",
    # Exodus
    "exo": "EXO", "exodus": "EXO", "ex": "EXO",
    # Leviticus
    "lev": "LEV", "leviticus": "LEV",
    # Numbers
    "num": "NUM", "numbers": "NUM", "numeri": "NUM",
    # Deuteronomy
    "deu": "DEU", "deuteronomy": "DEU", "deut": "DEU",
    # Joshua
    "jos": "JOS", "joshua": "JOS", "joz": "JOS",
    # Judges
    "jdg": "JDG", "judges": "JDG", "jud": "JDG", "richteren": "JDG",
    # Ruth
    "rut": "RUT", "ruth": "RUT",
    # 1 Samuel
    "1sa": "1SA", "1sam": "1SA", "1samuel": "1SA", "1samuël": "1SA",
    # 2 Samuel
    "2sa": "2SA", "2sam": "2SA", "2samuel": "2SA", "2samuël": "2SA",
    # 1 Kings
    "1ki": "1KI", "1kings": "1KI", "1kon": "1KI", "1koningen": "1KI",
    # 2 Kings
    "2ki": "2KI", "2kings": "2KI", "2kon": "2KI", "2koningen": "2KI",
    # 1 Chronicles
    "1ch": "1CH", "1chr": "1CH", "1chronicles": "1CH", "1kron": "1CH", "1kronieken": "1CH",
    # 2 Chronicles
    "2ch": "2CH", "2chr": "2CH", "2chronicles": "2CH", "2kron": "2CH", "2kronieken": "2CH",
    # Ezra
    "ezr": "EZR", "ezra": "EZR",
    # Nehemiah
    "neh": "NEH", "nehemiah": "NEH", "nehemia": "NEH",
    # Esther
    "est": "EST", "esther": "EST", "ester": "EST",
    # Job
    "job": "JOB",
    # Psalms
    "psa": "PSA", "psalms": "PSA", "psalm": "PSA", "ps": "PSA", "pss": "PSA",
    # Proverbs
    "pro": "PRO", "proverbs": "PRO", "prov": "PRO", "spreuken": "PRO",
    # Ecclesiastes
    "ecc": "ECC", "ecclesiastes": "ECC", "eccl": "ECC", "pred": "ECC", "prediker": "ECC",
    # Song of Solomon
    "son": "SNG", "sng": "SNG", "song": "SNG", "songofsolomon": "SNG",
    "hld": "SNG", "hooglied": "SNG",
    # Isaiah
    "isa": "ISA", "isaiah": "ISA", "jes": "ISA", "jesaja": "ISA",
    # Jeremiah
    "jer": "JER", "jeremiah": "JER",
    # Lamentations
    "lam": "LAM", "lamentations": "LAM", "klaagl": "LAM", "klaagliederen": "LAM",
    # Ezekiel
    "eze": "EZK", "ezk": "EZK", "ezekiel": "EZK", "ez": "EZK",
    "ezech": "EZK", "ezechiël": "EZK",
    # Daniel
    "dan": "DAN", "daniel": "DAN",
    # Hosea
    "hos": "HOS", "hosea": "HOS",
    # Joel
    "joe": "JOL", "jol": "JOL", "joel": "JOL",
    # Amos
    "amo": "AMO", "amos": "AMO",
    # Obadiah
    "oba": "OBA", "obadiah": "OBA", "obadja": "OBA",
    # Jonah
    "jon": "JON", "jonah": "JON", "jona": "JON",
    # Micah
    "mic": "MIC", "micah": "MIC", "micha": "MIC",
    # Nahum
    "nah": "NAH", "nahum": "NAH",
    # Habakkuk
    "hab": "HAB", "habakkuk": "HAB", "habakuk": "HAB",
    # Zephaniah
    "zep": "ZEP", "zephaniah": "ZEP", "zef": "ZEP", "zefanja": "ZEP",
    # Haggai
    "hag": "HAG", "haggai": "HAG", "haggaï": "HAG",
    # Zechariah
    "zec": "ZEC", "zechariah": "ZEC", "zach": "ZEC", "zacharia": "ZEC",
    # Malachi
    "mal": "MAL", "malachi": "MAL", "maleachi": "MAL",
    # Matthew
    "mat": "MAT", "matthew": "MAT", "matt": "MAT", "mt": "MAT",
    "matteüs": "MAT", "mattheus": "MAT",
    # Mark
    "mar": "MRK", "mrk": "MRK", "mark": "MRK", "mk": "MRK",
    "marc": "MRK", "marcus": "MRK",
    # Luke
    "luk": "LUK", "luke": "LUK", "lc": "LUK",
    "lukas": "LUK",
    # John
    "joh": "JHN", "jhn": "JHN", "john": "JHN", "jn": "JHN",
    "johannes": "JHN",
    # Acts
    "act": "ACT", "acts": "ACT", "hand": "ACT", "handelingen": "ACT",
    # Romans
    "rom": "ROM", "romans": "ROM",
    # 1 Corinthians
    "1co": "1CO", "1cor": "1CO", "1corinthians": "1CO",
    "1kor": "1CO", "1korintiërs": "1CO",
    # 2 Corinthians
    "2co": "2CO", "2cor": "2CO", "2corinthians": "2CO",
    "2kor": "2CO", "2korintiërs": "2CO",
    # Galatians
    "gal": "GAL", "galatians": "GAL", "galaten": "GAL",
    # Ephesians
    "eph": "EPH", "ephesians": "EPH", "ef": "EPH", "efeziërs": "EPH",
    # Philippians
    "php": "PHP", "phil": "PHP", "philippians": "PHP", "fil": "PHP", "filippenzen": "PHP",
    # Colossians
    "col": "COL", "colossians": "COL", "kol": "COL", "kolossenzen": "COL",
    # 1 Thessalonians
    "1th": "1TH", "1thess": "1TH", "1thessalonians": "1TH",
    "1tes": "1TH", "1tessalonicenzen": "1TH",
    # 2 Thessalonians
    "2th": "2TH", "2thess": "2TH", "2thessalonians": "2TH",
    "2tes": "2TH", "2tessalonicenzen": "2TH",
    # 1 Timothy
    "1ti": "1TI", "1tim": "1TI", "1timothy": "1TI",
    # 2 Timothy
    "2ti": "2TI", "2tim": "2TI", "2timothy": "2TI",
    # Titus
    "tit": "TIT", "titus": "TIT",
    # Philemon
    "phm": "PHM", "philemon": "PHM", "filem": "PHM",
    # Hebrews
    "heb": "HEB", "hebrews": "HEB",
    # James
    "jas": "JAS", "james": "JAS", "jak": "JAS", "jakobus": "JAS",
    # 1 Peter
    "1pe": "1PE", "1pet": "1PE", "1peter": "1PE", "1ptr": "1PE",
    "1petr": "1PE",
    # 2 Peter
    "2pe": "2PE", "2pet": "2PE", "2peter": "2PE", "2ptr": "2PE",
    "2petr": "2PE",
    # 1 John
    "1jo": "1JN", "1jn": "1JN", "1john": "1JN", "1joh": "1JN",
    # 2 John
    "2jo": "2JN", "2jn": "2JN", "2john": "2JN", "2joh": "2JN",
    # 3 John
    "3jo": "3JN", "3jn": "3JN", "3john": "3JN", "3joh": "3JN",
    # Jude
    "jude": "JUD", "judas": "JUD",
    # Revelation
    "rev": "REV", "revelation": "REV", "revelations": "REV",
    "opb": "REV", "openbaring": "REV",
}

# Human-readable book names (USFM code → English name)
USFM_TO_BOOK_NAME: Dict[str, str] = {
    "GEN": "Genesis", "EXO": "Exodus", "LEV": "Leviticus", "NUM": "Numbers",
    "DEU": "Deuteronomy", "JOS": "Joshua", "JDG": "Judges", "RUT": "Ruth",
    "1SA": "1 Samuel", "2SA": "2 Samuel", "1KI": "1 Kings", "2KI": "2 Kings",
    "1CH": "1 Chronicles", "2CH": "2 Chronicles", "EZR": "Ezra", "NEH": "Nehemiah",
    "EST": "Esther", "JOB": "Job", "PSA": "Psalms", "PRO": "Proverbs",
    "ECC": "Ecclesiastes", "SNG": "Song of Solomon", "ISA": "Isaiah",
    "JER": "Jeremiah", "LAM": "Lamentations", "EZK": "Ezekiel", "DAN": "Daniel",
    "HOS": "Hosea", "JOL": "Joel", "AMO": "Amos", "OBA": "Obadiah",
    "JON": "Jonah", "MIC": "Micah", "NAH": "Nahum", "HAB": "Habakkuk",
    "ZEP": "Zephaniah", "HAG": "Haggai", "ZEC": "Zechariah", "MAL": "Malachi",
    "MAT": "Matthew", "MRK": "Mark", "LUK": "Luke", "JHN": "John",
    "ACT": "Acts", "ROM": "Romans", "1CO": "1 Corinthians", "2CO": "2 Corinthians",
    "GAL": "Galatians", "EPH": "Ephesians", "PHP": "Philippians", "COL": "Colossians",
    "1TH": "1 Thessalonians", "2TH": "2 Thessalonians", "1TI": "1 Timothy",
    "2TI": "2 Timothy", "TIT": "Titus", "PHM": "Philemon", "HEB": "Hebrews",
    "JAS": "James", "1PE": "1 Peter", "2PE": "2 Peter", "1JN": "1 John",
    "2JN": "2 John", "3JN": "3 John", "JUD": "Jude", "REV": "Revelation",
}


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class BibleTranslation:
    """A Bible translation / version."""
    id: int
    abbreviation: str
    name: str
    language: str        # ISO 639-1 language code (e.g. "en", "nl")
    language_name: str   # Human-readable language name


@dataclass
class BibleVerse:
    """A single Bible verse."""
    verse_num: int
    text: str


@dataclass
class BibleReference:
    """A parsed Bible reference."""
    book_usfm: str      # USFM book code (e.g. "JHN")
    chapter: int
    verse_start: int
    verse_end: Optional[int] = None  # None = single verse

    @property
    def book_name(self) -> str:
        return USFM_TO_BOOK_NAME.get(self.book_usfm, self.book_usfm)

    @property
    def usfm_chapter(self) -> str:
        """USFM reference for the chapter (e.g. 'JHN.3')."""
        return f"{self.book_usfm}.{self.chapter}"

    @property
    def display_str(self) -> str:
        """Human-readable reference string."""
        if self.verse_end and self.verse_end != self.verse_start:
            return f"{self.book_name} {self.chapter}:{self.verse_start}-{self.verse_end}"
        return f"{self.book_name} {self.chapter}:{self.verse_start}"


# ---------------------------------------------------------------------------
# Reference parser
# ---------------------------------------------------------------------------

def parse_reference(reference_str: str) -> BibleReference:
    """Parse a Bible reference string into a BibleReference.

    Supports formats like:
      - "John 3:16"
      - "John 3:16-21"
      - "Psalm 23:1-6"
      - "1 Cor 13:1-13"
      - "Johannes 3:16"    (Dutch)
      - "Ps 23:1-6"
      - "Rev 22:1"

    Raises ValueError if the reference cannot be parsed.
    """
    s = reference_str.strip()
    if not s:
        raise ValueError("Empty reference string")

    # Pattern: optional digit prefix + book name + chapter:verse[-verse]
    # Allows: "1 Corinthians", "2 Tim", "Song of Solomon", "1Co", etc.
    pattern = re.compile(
        r"^(\d\s*)?([A-Za-zÀ-öø-ÿëïüäöé]+(?:\s+[A-Za-zÀ-öø-ÿëïüäöé]+)*)"
        r"\s+(\d+)\s*[:.]\s*(\d+)(?:\s*[-–—]\s*(\d+))?$",
        re.IGNORECASE | re.UNICODE,
    )

    m = pattern.match(s)
    if not m:
        raise ValueError(
            f"Cannot parse reference '{reference_str}'. "
            "Expected format: 'Book Chapter:VerseStart[-VerseEnd]', e.g. 'John 3:16-21'"
        )

    digit_prefix = (m.group(1) or "").strip()
    book_raw = m.group(2).strip()
    chapter = int(m.group(3))
    verse_start = int(m.group(4))
    verse_end_str = m.group(5)
    verse_end = int(verse_end_str) if verse_end_str else None

    # Build lookup key
    if digit_prefix:
        # e.g. "1 Corinthians" → "1corinthians", "2 Tim" → "2tim"
        lookup_key = (digit_prefix + book_raw).replace(" ", "").lower()
    else:
        lookup_key = book_raw.replace(" ", "").lower()

    # Try full key first, then progressively shorter prefixes
    usfm = None
    for key_candidate in _generate_lookup_candidates(lookup_key):
        if key_candidate in BOOK_NAME_TO_USFM:
            usfm = BOOK_NAME_TO_USFM[key_candidate]
            break

    if usfm is None:
        raise ValueError(
            f"Unknown book name: '{digit_prefix}{book_raw}'. "
            "Please use a standard book name or abbreviation."
        )

    if verse_end is not None and verse_end < verse_start:
        raise ValueError(
            f"End verse ({verse_end}) cannot be less than start verse ({verse_start})."
        )

    return BibleReference(
        book_usfm=usfm,
        chapter=chapter,
        verse_start=verse_start,
        verse_end=verse_end,
    )


def _generate_lookup_candidates(key: str) -> List[str]:
    """Generate lookup key candidates from most specific to least specific."""
    candidates = [key]
    # Try without trailing 's' (e.g. "psalms" → "psalm")
    if key.endswith("s") and len(key) > 3:
        candidates.append(key[:-1])
    # Try first 4 chars if key is long enough
    if len(key) >= 4:
        candidates.append(key[:4])
    if len(key) >= 3:
        candidates.append(key[:3])
    return candidates


# ---------------------------------------------------------------------------
# BibleService
# ---------------------------------------------------------------------------

class BibleService:
    """Fetches Bible texts from YouVersion (bible.com) API.

    Uses the unofficial nodejs.bible.com chapter API.
    Caches fetched chapters to avoid redundant network requests.
    """

    def __init__(self) -> None:
        self._chapter_cache: Dict[str, List[BibleVerse]] = {}
        self._translations_cache: Optional[List[BibleTranslation]] = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_builtin_translations(self) -> List[BibleTranslation]:
        """Return the built-in translation catalog (no network required)."""
        return [
            BibleTranslation(
                id=t["id"],
                abbreviation=t["abbreviation"],
                name=t["name"],
                language=t["language"],
                language_name=t["language_name"],
            )
            for t in BUILTIN_TRANSLATIONS
        ]

    def fetch_translations_for_language(
        self, language_tag: str
    ) -> List[BibleTranslation]:
        """Fetch available translations for a language tag from YouVersion.

        Falls back to built-in catalog on error.

        Args:
            language_tag: ISO 639-1 language code (e.g. "en", "nl").

        Returns:
            List of BibleTranslation objects.
        """
        try:
            resp = requests.get(
                YOUVERSION_VERSIONS_URL,
                params={"language_tag": language_tag, "type": "all"},
                headers=_HEADERS,
                timeout=REQUEST_TIMEOUT,
            )
            resp.raise_for_status()
            data = resp.json()

            versions = data if isinstance(data, list) else data.get("versions", [])
            results = []
            for v in versions:
                if not isinstance(v, dict):
                    continue
                results.append(
                    BibleTranslation(
                        id=int(v.get("id", 0)),
                        abbreviation=v.get("abbreviation", ""),
                        name=v.get("title", v.get("local_title", "")),
                        language=language_tag,
                        language_name=v.get("language_tag", language_tag),
                    )
                )
            return results
        except Exception as exc:
            logger.warning(
                "Failed to fetch translations from YouVersion for language '%s': %s. "
                "Using built-in catalog.",
                language_tag,
                exc,
            )
            return [
                t for t in self.get_builtin_translations()
                if t.language == language_tag
            ]

    def get_verses(
        self,
        reference: BibleReference,
        version_id: int,
    ) -> List[BibleVerse]:
        """Fetch verses for a Bible reference and translation.

        Retrieves the whole chapter and returns only the requested verse range.
        Results are cached by (version_id, usfm_chapter).

        Args:
            reference: Parsed BibleReference.
            version_id: YouVersion numeric version ID.

        Returns:
            List of BibleVerse objects (verse_num, text).

        Raises:
            requests.RequestException: On network failure.
            ValueError: If the requested verse range is not in the response.
        """
        cache_key = f"{version_id}:{reference.usfm_chapter}"
        if cache_key not in self._chapter_cache:
            self._chapter_cache[cache_key] = self._fetch_chapter(
                version_id, reference.usfm_chapter
            )

        all_verses = self._chapter_cache[cache_key]
        return self._filter_verses(all_verses, reference.verse_start, reference.verse_end)

    def get_youversion_url(
        self,
        reference: BibleReference,
        version_id: int,
    ) -> str:
        """Return the YouVersion share URL for the full reference.

        Example: https://www.bible.com/bible/111/JHN.3.16-21.NIV
        """
        usfm = reference.usfm_chapter
        usfm_with_verses = f"{usfm}.{reference.verse_start}"
        if reference.verse_end and reference.verse_end != reference.verse_start:
            usfm_with_verses += f"-{reference.verse_end}"
        return f"{YOUVERSION_SHARE_BASE}/{version_id}/{usfm_with_verses}"

    def clear_cache(self) -> None:
        """Clear the verse cache."""
        self._chapter_cache.clear()

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _fetch_chapter(self, version_id: int, usfm_chapter: str) -> List[BibleVerse]:
        """Fetch all verses for a chapter from the YouVersion API.

        The nodejs.bible.com API returns JSON with a 'items' or 'verses' array.
        """
        url = f"{YOUVERSION_API_BASE}/chapter/{version_id}/{usfm_chapter}.json"
        logger.info("Fetching Bible chapter: %s (version %s)", usfm_chapter, version_id)

        resp = requests.get(url, headers=_HEADERS, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()

        data = resp.json()
        return self._parse_chapter_response(data, usfm_chapter)

    def _parse_chapter_response(
        self, data: dict, usfm_chapter: str
    ) -> List[BibleVerse]:
        """Parse a YouVersion chapter API response into a list of BibleVerse."""
        verses: List[BibleVerse] = []

        # Format A: {"items": [{"type": "verse", "verseId": "JHN.3.1", "content": "..."}]}
        if "items" in data:
            for item in data["items"]:
                if not isinstance(item, dict):
                    continue
                if item.get("type") != "verse":
                    continue
                verse_id = item.get("verseId", item.get("verse_id", ""))
                verse_num = self._extract_verse_num(verse_id)
                if verse_num is None:
                    # Try 'human' label
                    try:
                        verse_num = int(item.get("human", item.get("label", "0")))
                    except (ValueError, TypeError):
                        continue
                text = _strip_html(_clean_verse_text(item.get("content", "")))
                if text:
                    verses.append(BibleVerse(verse_num=verse_num, text=text))

        # Format B: {"verses": [{"id": "JHN.3.1", "verse": 1, "text": "..."}]}
        elif "verses" in data:
            for item in data["verses"]:
                if not isinstance(item, dict):
                    continue
                verse_num = item.get("verse", item.get("verse_num"))
                if verse_num is None:
                    verse_id = item.get("id", item.get("verseId", ""))
                    verse_num = self._extract_verse_num(verse_id)
                if verse_num is None:
                    continue
                try:
                    verse_num = int(verse_num)
                except (ValueError, TypeError):
                    continue
                text = _strip_html(_clean_verse_text(item.get("text", item.get("content", ""))))
                if text:
                    verses.append(BibleVerse(verse_num=verse_num, text=text))

        # Format C: raw list of verse objects
        elif isinstance(data, list):
            for item in data:
                if not isinstance(item, dict):
                    continue
                verse_num = item.get("verse", item.get("verse_num", item.get("number")))
                if verse_num is None:
                    continue
                try:
                    verse_num = int(verse_num)
                except (ValueError, TypeError):
                    continue
                text = _strip_html(_clean_verse_text(item.get("text", item.get("content", ""))))
                if text:
                    verses.append(BibleVerse(verse_num=verse_num, text=text))

        if not verses:
            logger.warning(
                "No verses parsed for chapter '%s'. Response keys: %s",
                usfm_chapter,
                list(data.keys()) if isinstance(data, dict) else "list",
            )

        # Sort by verse number in case they come out of order
        verses.sort(key=lambda v: v.verse_num)
        return verses

    @staticmethod
    def _extract_verse_num(verse_id: str) -> Optional[int]:
        """Extract verse number from a USFM verse ID like 'JHN.3.16'."""
        if not verse_id:
            return None
        parts = str(verse_id).split(".")
        if len(parts) >= 3:
            try:
                return int(parts[-1])
            except ValueError:
                pass
        return None

    @staticmethod
    def _filter_verses(
        all_verses: List[BibleVerse],
        verse_start: int,
        verse_end: Optional[int],
    ) -> List[BibleVerse]:
        """Filter verse list to the requested range."""
        end = verse_end if verse_end is not None else verse_start
        return [v for v in all_verses if verse_start <= v.verse_num <= end]


# ---------------------------------------------------------------------------
# Text utilities
# ---------------------------------------------------------------------------

_HTML_TAG_RE = re.compile(r"<[^>]+>")
_WHITESPACE_RE = re.compile(r"\s+")


def _strip_html(text: str) -> str:
    """Remove HTML tags and normalize whitespace."""
    text = _HTML_TAG_RE.sub(" ", text)
    text = _WHITESPACE_RE.sub(" ", text)
    return text.strip()


def _clean_verse_text(text: str) -> str:
    """Remove leading verse numbers (e.g. '16 For God...') from YouVersion output."""
    # Some responses include the verse number at the start – strip it
    text = text.strip()
    # Remove patterns like "16 " or "(16) " at the start
    text = re.sub(r"^\d+\s+", "", text)
    return text
