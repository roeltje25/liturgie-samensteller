"""Bible text fetching service using the YouVersion Platform API.

Uses the official YouVersion Platform API (api.youversion.com/v1) to fetch
Bible texts in multiple translations.  A free API key can be obtained from
https://developers.youversion.com.
"""

import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import requests

from ..logging_config import get_logger

logger = get_logger("bible_service")

# YouVersion Platform API base URL
YOUVERSION_API_BASE = "https://api.youversion.com/v1"

# bible.com share URL (unchanged – same system)
YOUVERSION_SHARE_BASE = "https://www.bible.com/bible"

# Request timeout in seconds
REQUEST_TIMEOUT = 15

# Default request headers
_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
    "Accept-Language": "en-US,en;q=0.9",
}

# ---------------------------------------------------------------------------
# Built-in translation catalog
# YouVersion numeric IDs (same IDs used on bible.com).
# ---------------------------------------------------------------------------
BUILTIN_TRANSLATIONS = [
    # English
    {"id": 1,    "abbreviation": "KJV",      "name": "King James Version",               "language": "en", "language_name": "English"},
    {"id": 59,   "abbreviation": "ESV",      "name": "English Standard Version",         "language": "en", "language_name": "English"},
    {"id": 111,  "abbreviation": "NIV",      "name": "New International Version",        "language": "en", "language_name": "English"},
    {"id": 116,  "abbreviation": "NLT",      "name": "New Living Translation",           "language": "en", "language_name": "English"},
    {"id": 100,  "abbreviation": "NASB",     "name": "New American Standard Bible",      "language": "en", "language_name": "English"},
    {"id": 206,  "abbreviation": "WEB",      "name": "World English Bible",              "language": "en", "language_name": "English"},
    {"id": 97,   "abbreviation": "MSG",      "name": "The Message",                      "language": "en", "language_name": "English"},
    {"id": 37,   "abbreviation": "HCSB",     "name": "Holman Christian Standard Bible",  "language": "en", "language_name": "English"},
    # Dutch
    {"id": 48,   "abbreviation": "NBG51",    "name": "NBG-vertaling 1951",               "language": "nl", "language_name": "Nederlands"},
    {"id": 328,  "abbreviation": "NBV",      "name": "Nieuwe Bijbelvertaling",            "language": "nl", "language_name": "Nederlands"},
    {"id": 1816, "abbreviation": "HSV",      "name": "Herziene Statenvertaling",         "language": "nl", "language_name": "Nederlands"},
    {"id": 524,  "abbreviation": "BGT",      "name": "Bijbel in Gewone Taal",            "language": "nl", "language_name": "Nederlands"},
    {"id": 278,  "abbreviation": "SV",       "name": "Statenvertaling",                  "language": "nl", "language_name": "Nederlands"},
    # German
    {"id": 51,   "abbreviation": "LUTH1545", "name": "Luther Bibel 1545",                "language": "de", "language_name": "Deutsch"},
    {"id": 157,  "abbreviation": "ELB",      "name": "Elberfelder Bibel",                "language": "de", "language_name": "Deutsch"},
    {"id": 70,   "abbreviation": "SCH2000",  "name": "Schlachter 2000",                  "language": "de", "language_name": "Deutsch"},
    # French
    {"id": 93,   "abbreviation": "LSG",      "name": "Louis Segond 1910",                "language": "fr", "language_name": "Français"},
    {"id": 1462, "abbreviation": "BDS",      "name": "Bible du Semeur",                  "language": "fr", "language_name": "Français"},
    # Spanish
    {"id": 128,  "abbreviation": "RVR60",    "name": "Reina-Valera 1960",                "language": "es", "language_name": "Español"},
    {"id": 503,  "abbreviation": "NVI",      "name": "Nueva Versión Internacional",      "language": "es", "language_name": "Español"},
    # Portuguese
    {"id": 212,  "abbreviation": "ARC",      "name": "Almeida Revisada Corrigida",       "language": "pt", "language_name": "Português"},
    # Italian
    {"id": 431,  "abbreviation": "NR2006",   "name": "Nuova Riveduta 2006",              "language": "it", "language_name": "Italiano"},
    # Russian
    {"id": 400,  "abbreviation": "SYNO",     "name": "Синодальный перевод",              "language": "ru", "language_name": "Русский"},
    # Chinese Simplified
    {"id": 1268, "abbreviation": "CNVS",     "name": "Chinese New Version (Simplified)", "language": "zh", "language_name": "中文"},
    # Greek
    {"id": 2097, "abbreviation": "SBLG",     "name": "SBL Greek New Testament",         "language": "el", "language_name": "Ελληνικά"},
    # Hebrew
    {"id": 2310, "abbreviation": "WLC",      "name": "Westminster Leningrad Codex",      "language": "he", "language_name": "עברית"},
    # Arabic
    {"id": 3,    "abbreviation": "AVDB",     "name": "Arabic Bible",                    "language": "ar", "language_name": "العربية"},
]


# ---------------------------------------------------------------------------
# USFM book code mappings
# ---------------------------------------------------------------------------
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
    language: str
    language_name: str


@dataclass
class BibleVerse:
    """A single Bible verse."""
    verse_num: int
    text: str


@dataclass
class BibleReference:
    """A parsed Bible reference."""
    book_usfm: str
    chapter: int
    verse_start: int
    verse_end: Optional[int] = None

    @property
    def book_name(self) -> str:
        return USFM_TO_BOOK_NAME.get(self.book_usfm, self.book_usfm)

    @property
    def usfm_chapter(self) -> str:
        return f"{self.book_usfm}.{self.chapter}"

    @property
    def usfm_range(self) -> str:
        """USFM range string for the Platform API (e.g. 'JHN.3.16-JHN.3.21')."""
        start = f"{self.book_usfm}.{self.chapter}.{self.verse_start}"
        if self.verse_end and self.verse_end != self.verse_start:
            end = f"{self.book_usfm}.{self.chapter}.{self.verse_end}"
            return f"{start}-{end}"
        return start

    @property
    def display_str(self) -> str:
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

    Raises ValueError if the reference cannot be parsed.
    """
    s = reference_str.strip()
    if not s:
        raise ValueError("Empty reference string")

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

    if digit_prefix:
        lookup_key = (digit_prefix + book_raw).replace(" ", "").lower()
    else:
        lookup_key = book_raw.replace(" ", "").lower()

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
    candidates = [key]
    if key.endswith("s") and len(key) > 3:
        candidates.append(key[:-1])
    if len(key) >= 4:
        candidates.append(key[:4])
    if len(key) >= 3:
        candidates.append(key[:3])
    return candidates


# ---------------------------------------------------------------------------
# BibleService
# ---------------------------------------------------------------------------

class BibleService:
    """Fetches Bible texts from the YouVersion Platform API.

    Requires a free API key from https://developers.youversion.com.
    Set api_key before calling get_verses(); without a key the API will
    return 401 Unauthorized errors.
    """

    def __init__(self, api_key: str = "") -> None:
        self._api_key = api_key
        self._passage_cache: Dict[str, List[BibleVerse]] = {}

    def set_api_key(self, api_key: str) -> None:
        """Update the API key and clear the cache."""
        if api_key != self._api_key:
            self._api_key = api_key
            self._passage_cache.clear()

    def has_api_key(self) -> bool:
        return bool(self._api_key and self._api_key.strip())

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
        """Fetch available translations for a language from the Platform API.

        Uses GET /v1/bibles?language_tag={lang}&page_size=100.
        Falls back to built-in catalog on error or missing API key.
        """
        if not self.has_api_key():
            return [
                t for t in self.get_builtin_translations()
                if t.language == language_tag
            ]

        try:
            results = []
            page_token: Optional[str] = None

            while True:
                params: dict = {
                    "language_tag": language_tag,
                    "page_size": 100,
                }
                if page_token:
                    params["page_token"] = page_token

                resp = requests.get(
                    f"{YOUVERSION_API_BASE}/bibles",
                    headers=self._get_headers(),
                    params=params,
                    timeout=REQUEST_TIMEOUT,
                )
                resp.raise_for_status()
                data = resp.json()

                bibles = data.get("data", data) if isinstance(data, dict) else data
                if isinstance(bibles, list):
                    for b in bibles:
                        if not isinstance(b, dict):
                            continue
                        lang_info = b.get("language", {})
                        lang_code = (
                            lang_info.get("tag", language_tag)
                            if isinstance(lang_info, dict)
                            else language_tag
                        )
                        lang_name = (
                            lang_info.get("name", language_tag)
                            if isinstance(lang_info, dict)
                            else language_tag
                        )
                        results.append(
                            BibleTranslation(
                                id=int(b.get("id", 0)),
                                abbreviation=b.get("abbreviation", ""),
                                name=b.get("title", b.get("local_title", "")),
                                language=lang_code,
                                language_name=lang_name,
                            )
                        )

                # Pagination
                page_token = data.get("next_page_token") if isinstance(data, dict) else None
                if not page_token:
                    break

            logger.info(
                "Fetched %d translations for language '%s' from Platform API",
                len(results),
                language_tag,
            )
            return results

        except Exception as exc:
            logger.warning(
                "Failed to fetch translations from Platform API for '%s': %s. "
                "Falling back to built-in catalog.",
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
        """Fetch verses for a reference and translation using the Platform API.

        Results are cached by (version_id, usfm_range).

        Raises:
            requests.RequestException: On network failure.
            ValueError: If the API key is missing or response is unexpected.
        """
        if not self.has_api_key():
            raise ValueError(
                "No YouVersion API key configured. "
                "Please add your API key in Settings → Bible Text Slides. "
                "Get a free key at https://developers.youversion.com"
            )

        usfm_range = reference.usfm_range
        cache_key = f"{version_id}:{usfm_range}"
        if cache_key not in self._passage_cache:
            self._passage_cache[cache_key] = self._fetch_passage(version_id, usfm_range, reference)
        return self._passage_cache[cache_key]

    def get_youversion_url(self, reference: BibleReference, version_id: int) -> str:
        """Return the bible.com share URL for the reference."""
        usfm = reference.usfm_chapter
        usfm_with_verses = f"{usfm}.{reference.verse_start}"
        if reference.verse_end and reference.verse_end != reference.verse_start:
            usfm_with_verses += f"-{reference.verse_end}"
        return f"{YOUVERSION_SHARE_BASE}/{version_id}/{usfm_with_verses}"

    def clear_cache(self) -> None:
        self._passage_cache.clear()

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _get_headers(self) -> dict:
        headers = dict(_HEADERS)
        if self._api_key:
            headers["X-YVP-App-Key"] = self._api_key.strip()
        return headers

    def _fetch_passage(
        self,
        version_id: int,
        usfm_range: str,
        reference: BibleReference,
    ) -> List[BibleVerse]:
        """Fetch a passage from the YouVersion Platform API."""
        url = f"{YOUVERSION_API_BASE}/bibles/{version_id}/passages/{usfm_range}"
        logger.info("Fetching passage: %s (version %s)", usfm_range, version_id)

        resp = requests.get(url, headers=self._get_headers(), timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        data = resp.json()
        return self._parse_passage_response(data, usfm_range, reference)

    def _parse_passage_response(
        self,
        data: dict,
        usfm_range: str,
        reference: BibleReference,
    ) -> List[BibleVerse]:
        """Parse the Platform API passage response into a list of BibleVerse.

        The Platform API may return verses nested in data.content, data.verses,
        or as a flat text block.  We try each format in order.
        """
        # Unwrap outer "data" envelope if present
        payload = data.get("data", data) if isinstance(data, dict) else data

        verses: List[BibleVerse] = []

        if isinstance(payload, dict):
            # Format A: {content: [{type:"verse", verse_id:"JHN.3.16", content:"..."}]}
            content_list = payload.get("content", [])
            if isinstance(content_list, list):
                for item in content_list:
                    if not isinstance(item, dict):
                        continue
                    if item.get("type") == "verse":
                        verse_id = item.get("verse_id", item.get("verseId", ""))
                        verse_num = _extract_verse_num_from_id(verse_id)
                        text = _clean_text(item.get("content", item.get("text", "")))
                        if verse_num is not None and text:
                            verses.append(BibleVerse(verse_num=verse_num, text=text))

            # Format B: {verses: [{verse_id:"JHN.3.16", text:"..."}]}
            if not verses:
                verse_list = payload.get("verses", [])
                if isinstance(verse_list, list):
                    for item in verse_list:
                        if not isinstance(item, dict):
                            continue
                        verse_id = item.get("verse_id", item.get("verseId", item.get("id", "")))
                        verse_num = (
                            item.get("verse", item.get("verse_num"))
                            or _extract_verse_num_from_id(str(verse_id))
                        )
                        try:
                            verse_num = int(verse_num) if verse_num is not None else None
                        except (ValueError, TypeError):
                            verse_num = None
                        text = _clean_text(item.get("text", item.get("content", "")))
                        if verse_num is not None and text:
                            verses.append(BibleVerse(verse_num=verse_num, text=text))

            # Format C: flat text block — split on embedded verse numbers
            if not verses:
                raw_text = payload.get("content", "")
                if isinstance(raw_text, str) and raw_text.strip():
                    verses = _split_text_into_verses(raw_text, reference)

        if not verses:
            logger.warning(
                "No verses parsed for %s. Response keys: %s",
                usfm_range,
                list(payload.keys()) if isinstance(payload, dict) else type(payload).__name__,
            )

        verses.sort(key=lambda v: v.verse_num)
        return verses


# ---------------------------------------------------------------------------
# Text utilities
# ---------------------------------------------------------------------------

_HTML_TAG_RE = re.compile(r"<[^>]+>")
_WHITESPACE_RE = re.compile(r"\s+")


def _clean_text(text: str) -> str:
    """Strip HTML tags, normalize whitespace, and remove leading verse numbers."""
    text = _HTML_TAG_RE.sub(" ", text)
    text = _WHITESPACE_RE.sub(" ", text)
    text = text.strip()
    # Remove leading verse number like "16 " or "(16) "
    text = re.sub(r"^\(?\d+\)?\s+", "", text)
    return text.strip()


def _extract_verse_num_from_id(verse_id: str) -> Optional[int]:
    """Extract verse number from a USFM ID like 'JHN.3.16'."""
    if not verse_id:
        return None
    parts = str(verse_id).split(".")
    if len(parts) >= 3:
        try:
            return int(parts[-1])
        except ValueError:
            pass
    return None


def _split_text_into_verses(text: str, reference: BibleReference) -> List[BibleVerse]:
    """Split a flat passage text block into individual verses.

    Looks for patterns like '[16]', '16 ', '(16)' at the start of verse segments.
    If no verse markers are found, returns the whole text as a single verse starting
    at verse_start.
    """
    # Try splitting on patterns like [16] or (16)
    parts = re.split(r"\[(\d+)\]|\((\d+)\)", text)
    if len(parts) > 1:
        verses = []
        # parts alternates: text, group1, group2, text, group1, group2, ...
        i = 0
        while i < len(parts):
            pre_text = parts[i].strip()
            if i + 2 < len(parts):
                num_str = parts[i + 1] or parts[i + 2]
                try:
                    verse_num = int(num_str)
                    verse_text = parts[i + 3].strip() if i + 3 < len(parts) else ""
                    if verse_text:
                        verses.append(BibleVerse(verse_num=verse_num, text=_clean_text(verse_text)))
                    i += 4
                    continue
                except (ValueError, TypeError):
                    pass
            i += 1
        if verses:
            return verses

    # Fallback: return as single verse at verse_start
    cleaned = _clean_text(text)
    if cleaned:
        return [BibleVerse(verse_num=reference.verse_start, text=cleaned)]
    return []
