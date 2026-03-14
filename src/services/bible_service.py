"""Bible text fetching service using the YouVersion Platform API.

Uses the official YouVersion Platform API (api.youversion.com/v1).
A free API key is obtainable from https://developers.youversion.com.
Authenticate with header  x-yvp-app-key: {key}

Response format (passage endpoint):
  {"id": "JHN.3.16", "content": "For God so loved...", "reference": "John 3:16"}
Multi-verse passages include inline verse numbers in the content text.
"""

import re
from dataclasses import dataclass
from typing import Dict, List, Optional

import requests

from ..logging_config import get_logger

logger = get_logger("bible_service")

# ---------------------------------------------------------------------------
# API configuration
# ---------------------------------------------------------------------------
YOUVERSION_API_BASE = "https://api.youversion.com/v1"
YOUVERSION_SHARE_BASE = "https://www.bible.com/bible"
REQUEST_TIMEOUT = 15

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
}

# ---------------------------------------------------------------------------
# Built-in translation catalog
# IDs are YouVersion Platform API numeric IDs.
# NOTE: IDs differ from the old bible.com web IDs.
#       Known correct Platform API IDs are marked (✓).
#       Others are best-effort and can be refreshed via "Fetch more".
# ---------------------------------------------------------------------------
BUILTIN_TRANSLATIONS = [
    # English
    {"id": 3034, "abbreviation": "NIV",      "name": "New International Version (2011)", "language": "en", "language_name": "English"},   # ✓ confirmed
    {"id": 1,    "abbreviation": "KJV",      "name": "King James Version",               "language": "en", "language_name": "English"},
    {"id": 59,   "abbreviation": "ESV",      "name": "English Standard Version",         "language": "en", "language_name": "English"},
    {"id": 116,  "abbreviation": "NLT",      "name": "New Living Translation",           "language": "en", "language_name": "English"},
    {"id": 100,  "abbreviation": "NASB",     "name": "New American Standard Bible",      "language": "en", "language_name": "English"},
    {"id": 206,  "abbreviation": "WEB",      "name": "World English Bible",              "language": "en", "language_name": "English"},
    {"id": 97,   "abbreviation": "MSG",      "name": "The Message",                      "language": "en", "language_name": "English"},
    # Dutch
    {"id": 328,  "abbreviation": "NBV",      "name": "Nieuwe Bijbelvertaling",            "language": "nl", "language_name": "Nederlands"},
    {"id": 1816, "abbreviation": "HSV",      "name": "Herziene Statenvertaling",         "language": "nl", "language_name": "Nederlands"},
    {"id": 48,   "abbreviation": "NBG51",    "name": "NBG-vertaling 1951",               "language": "nl", "language_name": "Nederlands"},
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
    # Arabic
    {"id": 3,    "abbreviation": "AVDB",     "name": "Arabic Bible",                    "language": "ar", "language_name": "العربية"},
    # Persian / Farsi
    {"id": 131,  "abbreviation": "PCB",      "name": "Persian Contemporary Bible",      "language": "fa", "language_name": "فارسی"},
    {"id": 341,  "abbreviation": "NMV",      "name": "Farsi New Millennium Version",    "language": "fa", "language_name": "فارسی"},
    # Chinese (Simplified)
    {"id": 2,    "abbreviation": "CUNPSS",   "name": "Chinese Union Version (Simplified)", "language": "zh-Hans", "language_name": "中文(简体)"},
    # Chinese (Traditional)
    {"id": 46,   "abbreviation": "CUVT",     "name": "Chinese Union Version (Traditional)", "language": "zh-Hant", "language_name": "中文(繁體)"},
    # Korean
    {"id": 136,  "abbreviation": "KRV",      "name": "Korean Revised Version",          "language": "ko", "language_name": "한국어"},
    {"id": 1877, "abbreviation": "NKRV",     "name": "New Korean Revised Version",      "language": "ko", "language_name": "한국어"},
    # Turkish
    {"id": 173,  "abbreviation": "TCL02",    "name": "Türkçe",                          "language": "tr", "language_name": "Türkçe"},
    # Polish
    {"id": 35,   "abbreviation": "UBG",      "name": "Uwspółcześniona Biblia Gdańska",  "language": "pl", "language_name": "Polski"},
    # Indonesian
    {"id": 306,  "abbreviation": "TB",       "name": "Terjemahan Baru",                 "language": "id", "language_name": "Indonesia"},
    # Romanian
    {"id": 191,  "abbreviation": "RMNN",     "name": "Biblia Nouă",                     "language": "ro", "language_name": "Română"},
    # Ukrainian
    {"id": 186,  "abbreviation": "UKR",      "name": "Ukrainian Bible",                 "language": "uk", "language_name": "Українська"},
    # Hungarian
    {"id": 90,   "abbreviation": "KAR",      "name": "Károli Biblia",                   "language": "hu", "language_name": "Magyar"},
    # Swedish
    {"id": 167,  "abbreviation": "SVL",      "name": "Svenska Folkbibeln",              "language": "sv", "language_name": "Svenska"},
    # Norwegian
    {"id": 65,   "abbreviation": "NB",       "name": "Norsk Bibel",                     "language": "no", "language_name": "Norsk"},
    # Danish
    {"id": 43,   "abbreviation": "DN1933",   "name": "Bibelen på dansk",                "language": "da", "language_name": "Dansk"},
    # Finnish
    {"id": 330,  "abbreviation": "FB92",     "name": "Raamattu 1992",                   "language": "fi", "language_name": "Suomi"},
    # Greek (Modern)
    {"id": 2079, "abbreviation": "GKPNT",    "name": "Greek New Testament (Tischendorf)", "language": "el", "language_name": "Ελληνικά"},
    # Hebrew
    {"id": 2302, "abbreviation": "WLC",      "name": "Westminster Leningrad Codex",     "language": "he", "language_name": "עברית"},
    # Hindi
    {"id": 40,   "abbreviation": "ERV-HI",   "name": "Easy-to-Read Version (Hindi)",    "language": "hi", "language_name": "हिन्दी"},
    # Tagalog / Filipino
    {"id": 54,   "abbreviation": "TAB",      "name": "Tagalog Bible",                   "language": "tl", "language_name": "Filipino"},
]

# ---------------------------------------------------------------------------
# Verse counts per chapter (0-indexed: index 0 = chapter 1)
# Source: standard Protestant KJV/NIV verse counts.
# Used by _fetch_cross_chapter to build exact valid API ranges.
# ---------------------------------------------------------------------------
CHAPTER_VERSE_COUNTS: Dict[str, List[int]] = {
    "GEN": [31,25,24,26,32,22,24,22,29,32,32,20,18,24,21,16,27,33,38,18,34,24,20,67,34,35,46,22,35,43,55,32,20,31,29,43,36,30,23,23,57,38,34,34,28,34,31,22,33,26],
    "EXO": [22,25,22,31,23,30,25,32,35,29,10,51,22,31,27,36,16,27,25,26,36,31,33,18,40,37,21,43,46,38,18,35,23,35,35,38,29,31,43,38],
    "LEV": [17,16,17,35,19,30,38,36,24,20,47,8,59,57,33,34,16,30,37,27,24,33,44,23,55,46,34],
    "NUM": [54,34,51,49,31,27,89,26,23,36,35,16,33,45,41,50,13,32,22,29,35,41,30,25,18,65,23,31,40,16,54,42,56,29,34,13],
    "DEU": [46,37,29,49,33,25,26,20,29,22,32,32,18,29,23,22,20,22,21,20,23,30,25,22,19,19,26,68,29,20,30,52,29,12],
    "JOS": [18,24,17,24,15,27,26,35,27,43,23,24,33,15,63,10,18,28,51,9,45,34,16,33],
    "JDG": [36,23,31,24,31,40,25,35,57,18,40,15,25,20,20,31,13,31,30,48,25],
    "RUT": [22,23,18,22],
    "1SA": [28,36,21,22,12,21,17,22,27,27,15,25,23,52,35,23,58,30,24,42,15,23,29,22,44,25,12,25,11,31,13],
    "2SA": [27,32,39,12,25,23,29,18,13,19,27,31,39,33,37,23,29,33,43,26,22,51,39,25],
    "1KI": [53,46,28,34,18,38,51,66,28,29,43,33,34,31,34,34,24,46,21,43,29,53],
    "2KI": [18,25,27,44,27,33,20,29,37,36,21,21,25,29,38,20,41,37,37,21,26,20,37,20,30],
    "1CH": [54,55,24,43,26,81,40,40,44,14,47,40,14,17,29,43,27,17,19,8,30,19,32,31,31,32,34,21,30],
    "2CH": [17,18,17,22,14,42,22,18,31,19,23,16,22,15,19,14,19,34,11,37,20,12,21,27,28,23,9,27,36,27,21,33,25,33,27,23],
    "EZR": [11,70,13,24,17,22,28,36,15,44],
    "NEH": [11,20,32,23,19,19,73,18,38,39,36,47,31],
    "EST": [22,23,15,17,14,14,10,17,32,3],
    "JOB": [22,17,16,21,20,29,20,37,17,12,22,13,20,22,13,17,11,14,17,23,19,11,17,21,22,17,22,13,14,17,17,23,17,21,28,16,33,21,22,25,29,22],
    "PSA": [6,12,8,8,12,10,17,9,20,18,7,8,6,7,5,11,15,50,14,9,13,31,6,10,22,12,14,9,11,12,24,11,22,22,28,12,40,22,13,17,13,11,5,26,17,11,9,14,20,23,19,9,6,7,23,13,11,11,17,12,8,12,11,10,13,20,7,35,36,5,24,20,28,23,10,12,20,72,13,19,16,8,18,12,13,17,7,18,52,17,16,15,5,23,11,13,12,9,9,5,8,28,22,35,45,48,43,13,31,7,10,10,9,8,18,19,2,29,176,7,8,9,4,8,5,6,5,6,8,8,3,18,3,3,21,26,9,8,24,13,10,7,12,15,21,10,20,14,9,6],
    "PRO": [33,22,35,27,23,35,27,36,18,32,31,28,25,35,23,17,27,22,21,22,15,25,17,25,6,11,23,12,27,13,31],
    "ECC": [18,26,22,16,20,12,29,17,18,20,10,14],
    "SNG": [17,17,11,16,16,13,13,14],
    "ISA": [31,22,26,6,30,13,25,22,21,34,16,6,22,32,9,14,14,7,25,6,17,25,18,23,12,21,13,29,24,33,9,20,24,17,10,22,38,22,8,31,29,25,28,28,25,13,15,22,26,11,23,15,12,17,13,12,21,14,21,22,11,12,19,12,25,24],
    "JER": [19,37,25,31,31,30,34,22,26,25,23,17,27,22,21,21,27,23,15,18,14,30,40,10,38,24,22,17,32,24,40,44,26,22,19,32,21,28,18,16,18,22,13,30,5,28,7,47,39,46,64,34],
    "LAM": [22,22,66,22,22],
    "EZK": [28,10,27,17,17,14,27,18,11,22,25,28,23,23,8,63,24,32,14,49,32,31,49,27,17,21,36,26,21,26,18,32,33,31,15,38,28,23,29,49,26,20,27,31,25,24,23,35],
    "DAN": [21,49,30,37,31,28,28,27,27,21,45,13],
    "HOS": [11,23,5,19,15,11,16,14,17,15,12,14,16,9],
    "JOL": [20,32,21],
    "AMO": [15,16,15,13,27,14,17,14,15],
    "OBA": [21],
    "JON": [17,10,10,11],
    "MIC": [16,13,12,13,15,16,20],
    "NAH": [15,13,19],
    "HAB": [17,20,19],
    "ZEP": [18,15,20],
    "HAG": [15,23],
    "ZEC": [21,13,10,14,11,15,14,23,17,12,17,14,9,21],
    "MAL": [14,17,18,6],
    # New Testament
    "MAT": [25,23,17,25,48,34,29,34,38,42,30,50,58,36,39,28,27,35,30,34,46,46,39,51,46,75,66,20],
    "MRK": [45,28,35,41,43,56,37,38,50,52,33,44,37,72,47,20],
    "LUK": [80,52,38,44,39,49,50,56,62,42,54,59,35,35,32,31,37,43,48,47,38,71,56,53],
    "JHN": [51,25,36,54,47,71,53,59,41,42,57,50,38,31,27,33,26,40,42,31,25],
    "ACT": [26,47,26,37,42,15,60,40,43,48,30,25,52,28,41,40,34,28,41,38,40,30,35,27,27,32,44,31],
    "ROM": [32,29,31,25,21,23,25,39,33,21,36,21,14,23,33,27],
    "1CO": [31,16,23,21,13,20,40,13,27,33,34,31,13,40,58,24],
    "2CO": [24,17,18,18,21,20,15,25,19,12,17,25,14],
    "GAL": [24,21,29,31,26,18],
    "EPH": [23,22,21,32,33,24],
    "PHP": [30,30,21,23],
    "COL": [29,23,25,18],
    "1TH": [10,20,13,18,28],
    "2TH": [12,17,18],
    "1TI": [20,15,16,16,25,21],
    "2TI": [18,26,17,22],
    "TIT": [16,15,15],
    "PHM": [25],
    "HEB": [14,18,19,16,14,20,28,13,28,39,40,29,25],
    "JAS": [27,26,18,17,20],
    "1PE": [25,25,22,19,14],
    "2PE": [21,22,18],
    "1JN": [10,29,24,21,21],
    "2JN": [13],
    "3JN": [14],
    "JUD": [25],
    "REV": [20,29,22,11,14,17,17,13,21,11,19,17,18,20,8,21,18,24,21,15,27,21],
}

# ---------------------------------------------------------------------------
# USFM book code mappings
# ---------------------------------------------------------------------------
BOOK_NAME_TO_USFM: Dict[str, str] = {
    "gen": "GEN", "genesis": "GEN",
    "exo": "EXO", "exodus": "EXO", "ex": "EXO",
    "lev": "LEV", "leviticus": "LEV",
    "num": "NUM", "numbers": "NUM", "numeri": "NUM",
    "deu": "DEU", "deuteronomy": "DEU", "deut": "DEU",
    "jos": "JOS", "joshua": "JOS", "joz": "JOS",
    "jdg": "JDG", "judges": "JDG", "richteren": "JDG",
    "rut": "RUT", "ruth": "RUT",
    "1sa": "1SA", "1sam": "1SA", "1samuel": "1SA", "1samuël": "1SA",
    "2sa": "2SA", "2sam": "2SA", "2samuel": "2SA", "2samuël": "2SA",
    "1ki": "1KI", "1kings": "1KI", "1kon": "1KI", "1koningen": "1KI",
    "2ki": "2KI", "2kings": "2KI", "2kon": "2KI", "2koningen": "2KI",
    "1ch": "1CH", "1chr": "1CH", "1chronicles": "1CH", "1kron": "1CH", "1kronieken": "1CH",
    "2ch": "2CH", "2chr": "2CH", "2chronicles": "2CH", "2kron": "2CH", "2kronieken": "2CH",
    "ezr": "EZR", "ezra": "EZR",
    "neh": "NEH", "nehemiah": "NEH", "nehemia": "NEH",
    "est": "EST", "esther": "EST", "ester": "EST",
    "job": "JOB",
    "psa": "PSA", "psalms": "PSA", "psalm": "PSA", "ps": "PSA", "pss": "PSA",
    "pro": "PRO", "proverbs": "PRO", "prov": "PRO", "spreuken": "PRO",
    "ecc": "ECC", "ecclesiastes": "ECC", "eccl": "ECC", "pred": "ECC", "prediker": "ECC",
    "son": "SNG", "sng": "SNG", "song": "SNG", "songofsolomon": "SNG",
    "hld": "SNG", "hooglied": "SNG",
    "isa": "ISA", "isaiah": "ISA", "jes": "ISA", "jesaja": "ISA",
    "jer": "JER", "jeremiah": "JER",
    "lam": "LAM", "lamentations": "LAM", "klaagl": "LAM", "klaagliederen": "LAM",
    "eze": "EZK", "ezk": "EZK", "ezekiel": "EZK", "ez": "EZK",
    "ezech": "EZK", "ezechiël": "EZK",
    "dan": "DAN", "daniel": "DAN",
    "hos": "HOS", "hosea": "HOS",
    "joe": "JOL", "jol": "JOL", "joel": "JOL",
    "amo": "AMO", "amos": "AMO",
    "oba": "OBA", "obadiah": "OBA", "obadja": "OBA",
    "jon": "JON", "jonah": "JON", "jona": "JON",
    "mic": "MIC", "micah": "MIC", "micha": "MIC",
    "nah": "NAH", "nahum": "NAH",
    "hab": "HAB", "habakkuk": "HAB", "habakuk": "HAB",
    "zep": "ZEP", "zephaniah": "ZEP", "zef": "ZEP", "zefanja": "ZEP",
    "hag": "HAG", "haggai": "HAG", "haggaï": "HAG",
    "zec": "ZEC", "zechariah": "ZEC", "zach": "ZEC", "zacharia": "ZEC",
    "mal": "MAL", "malachi": "MAL", "maleachi": "MAL",
    "mat": "MAT", "matthew": "MAT", "matt": "MAT", "mt": "MAT",
    "matteüs": "MAT", "mattheus": "MAT",
    "mar": "MRK", "mrk": "MRK", "mark": "MRK", "mk": "MRK",
    "marc": "MRK", "marcus": "MRK",
    "luk": "LUK", "luke": "LUK", "lc": "LUK", "lukas": "LUK",
    "joh": "JHN", "jhn": "JHN", "john": "JHN", "jn": "JHN",
    "johannes": "JHN",
    "act": "ACT", "acts": "ACT", "hand": "ACT", "handelingen": "ACT",
    "rom": "ROM", "romans": "ROM",
    "1co": "1CO", "1cor": "1CO", "1corinthians": "1CO",
    "1kor": "1CO", "1korintiërs": "1CO",
    "2co": "2CO", "2cor": "2CO", "2corinthians": "2CO",
    "2kor": "2CO", "2korintiërs": "2CO",
    "gal": "GAL", "galatians": "GAL", "galaten": "GAL",
    "eph": "EPH", "ephesians": "EPH", "ef": "EPH", "efeziërs": "EPH",
    "php": "PHP", "phil": "PHP", "philippians": "PHP", "fil": "PHP", "filippenzen": "PHP",
    "col": "COL", "colossians": "COL", "kol": "COL", "kolossenzen": "COL",
    "1th": "1TH", "1thess": "1TH", "1thessalonians": "1TH",
    "1tes": "1TH", "1tessalonicenzen": "1TH",
    "2th": "2TH", "2thess": "2TH", "2thessalonians": "2TH",
    "2tes": "2TH", "2tessalonicenzen": "2TH",
    "1ti": "1TI", "1tim": "1TI", "1timothy": "1TI",
    "2ti": "2TI", "2tim": "2TI", "2timothy": "2TI",
    "tit": "TIT", "titus": "TIT",
    "phm": "PHM", "philemon": "PHM", "filem": "PHM",
    "heb": "HEB", "hebrews": "HEB",
    "jas": "JAS", "james": "JAS", "jak": "JAS", "jakobus": "JAS",
    "1pe": "1PE", "1pet": "1PE", "1peter": "1PE", "1ptr": "1PE", "1petr": "1PE",
    "2pe": "2PE", "2pet": "2PE", "2peter": "2PE", "2ptr": "2PE", "2petr": "2PE",
    "1jo": "1JN", "1jn": "1JN", "1john": "1JN", "1joh": "1JN",
    "2jo": "2JN", "2jn": "2JN", "2john": "2JN", "2joh": "2JN",
    "3jo": "3JN", "3jn": "3JN", "3john": "3JN", "3joh": "3JN",
    "jude": "JUD", "judas": "JUD",
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
    id: int
    abbreviation: str
    name: str
    language: str
    language_name: str


@dataclass
class BibleVerse:
    verse_num: int
    text: str


@dataclass
class BibleReference:
    """A parsed Bible reference.

    Supports:
      John 3:16          – single verse
      John 3:16-21       – same-chapter range
      Matthew 1:10-2:4   – cross-chapter range  (chapter_end set)
      Psalm 46           – whole chapter         (whole_chapter=True)
    """
    book_usfm: str
    chapter: int
    verse_start: int = 1
    verse_end: Optional[int] = None       # None = to end of chapter if whole_chapter
    chapter_end: Optional[int] = None     # set for cross-chapter ranges
    whole_chapter: bool = False

    @property
    def book_name(self) -> str:
        return USFM_TO_BOOK_NAME.get(self.book_usfm, self.book_usfm)

    @property
    def usfm_chapter(self) -> str:
        return f"{self.book_usfm}.{self.chapter}"

    @property
    def usfm_passage(self) -> str:
        """USFM reference string for the Platform API passages endpoint."""
        if self.whole_chapter:
            # Whole chapter: use chapter-level USFM (no verse numbers)
            return f"{self.book_usfm}.{self.chapter}"
        start = f"{self.book_usfm}.{self.chapter}.{self.verse_start}"
        ch_end = self.chapter_end if self.chapter_end else self.chapter
        v_end = self.verse_end if self.verse_end else self.verse_start
        if ch_end == self.chapter and v_end == self.verse_start:
            return start   # single verse
        return f"{start}-{self.book_usfm}.{ch_end}.{v_end}"

    @property
    def display_str(self) -> str:
        book = self.book_name
        if self.whole_chapter:
            return f"{book} {self.chapter}"
        if self.chapter_end and self.chapter_end != self.chapter:
            return f"{book} {self.chapter}:{self.verse_start}–{self.chapter_end}:{self.verse_end}"
        if self.verse_end and self.verse_end != self.verse_start:
            return f"{book} {self.chapter}:{self.verse_start}–{self.verse_end}"
        return f"{book} {self.chapter}:{self.verse_start}"


# ---------------------------------------------------------------------------
# Reference parsers
# ---------------------------------------------------------------------------

# Pattern: Book Chapter:Verse[-[Chapter:]Verse]
_VERSE_PATTERN = re.compile(
    r"^(\d\s*)?([A-Za-zÀ-öø-ÿëïüäöé]+(?:\s+[A-Za-zÀ-öø-ÿëïüäöé]+)*)"
    r"\s+(\d+)\s*[:.]\s*(\d+)"
    r"(?:\s*[-–—]\s*(?:(\d+)\s*[:.]\s*)?(\d+))?$",
    re.IGNORECASE | re.UNICODE,
)

# Pattern: Book Chapter  (whole chapter)
_CHAPTER_PATTERN = re.compile(
    r"^(\d\s*)?([A-Za-zÀ-öø-ÿëïüäöé]+(?:\s+[A-Za-zÀ-öø-ÿëïüäöé]+)*)"
    r"\s+(\d+)\s*$",
    re.IGNORECASE | re.UNICODE,
)


def parse_reference(reference_str: str) -> BibleReference:
    """Parse a single Bible reference string.

    Supported formats:
      - "John 3:16"
      - "John 3:16-21"
      - "Matthew 1:10-2:4"  (cross-chapter)
      - "Psalm 46"           (whole chapter)
      - "1 Cor 13:1-13"
      - "Johannes 3:16"      (Dutch names supported)
    """
    s = reference_str.strip()
    if not s:
        raise ValueError("Empty reference string")

    m = _VERSE_PATTERN.match(s)
    if m:
        digit_prefix = (m.group(1) or "").strip()
        book_raw = m.group(2).strip()
        chapter = int(m.group(3))
        verse_start = int(m.group(4))
        chapter_end_str = m.group(5)
        verse_end_str = m.group(6)

        usfm = _resolve_book(digit_prefix, book_raw, s)

        verse_end: Optional[int] = None
        chapter_end: Optional[int] = None

        if verse_end_str:
            verse_end = int(verse_end_str)
            if chapter_end_str:
                chapter_end = int(chapter_end_str)
            elif verse_end < verse_start:
                raise ValueError(
                    f"End verse ({verse_end}) cannot be less than start verse ({verse_start})."
                )

        return BibleReference(
            book_usfm=usfm,
            chapter=chapter,
            verse_start=verse_start,
            verse_end=verse_end,
            chapter_end=chapter_end,
            whole_chapter=False,
        )

    m2 = _CHAPTER_PATTERN.match(s)
    if m2:
        digit_prefix = (m2.group(1) or "").strip()
        book_raw = m2.group(2).strip()
        chapter = int(m2.group(3))
        usfm = _resolve_book(digit_prefix, book_raw, s)
        return BibleReference(book_usfm=usfm, chapter=chapter, whole_chapter=True)

    raise ValueError(
        f"Cannot parse '{reference_str}'. "
        "Expected format: 'Book Chapter[:Verse[-Verse]]', e.g. 'John 3:16-21', 'Psalm 46'"
    )


def parse_references(reference_str: str) -> List[BibleReference]:
    """Parse a comma-separated list of Bible references.

    Examples:
      "John 3:16-21"
      "Psalm 1:1, Revelation 3:2, John 3:16"
      "Matthew 1:10-2:4"
    """
    parts = [p.strip() for p in reference_str.split(",") if p.strip()]
    if not parts:
        raise ValueError("Empty reference string")
    return [parse_reference(p) for p in parts]


def _resolve_book(digit_prefix: str, book_raw: str, original: str) -> str:
    key = (digit_prefix + book_raw).replace(" ", "").lower() if digit_prefix else book_raw.replace(" ", "").lower()
    for candidate in _lookup_candidates(key):
        if candidate in BOOK_NAME_TO_USFM:
            return BOOK_NAME_TO_USFM[candidate]
    raise ValueError(
        f"Unknown book name: '{digit_prefix}{book_raw}' in '{original}'. "
        "Please use a standard book name or abbreviation."
    )


def _lookup_candidates(key: str) -> List[str]:
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
    Authenticate via Settings → Bible Text Slides.
    """

    def __init__(self, api_key: str = "") -> None:
        self._api_key = api_key.strip()
        self._passage_cache: Dict[str, List[BibleVerse]] = {}

    def set_api_key(self, api_key: str) -> None:
        if api_key.strip() != self._api_key:
            self._api_key = api_key.strip()
            self._passage_cache.clear()

    def has_api_key(self) -> bool:
        return bool(self._api_key)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_builtin_translations(self) -> List[BibleTranslation]:
        return [BibleTranslation(**t) for t in BUILTIN_TRANSLATIONS]

    def fetch_translations_for_language(self, language_tag: str) -> List[BibleTranslation]:
        """Fetch available translations for a language from the Platform API.

        Falls back to built-in catalog if the API is unavailable.
        """
        if not self.has_api_key():
            return [t for t in self.get_builtin_translations() if t.language == language_tag]
        try:
            results = self._fetch_bibles_list(language_tag)
            logger.info("Fetched %d translations for '%s'", len(results), language_tag)
            return results
        except Exception as exc:
            logger.warning(
                "Failed to fetch translations for '%s': %s. Using built-in catalog.",
                language_tag, exc,
            )
            return [t for t in self.get_builtin_translations() if t.language == language_tag]

    def get_verses(self, reference: BibleReference, version_id: int) -> List[BibleVerse]:
        """Fetch verses for a reference and translation.

        Supports whole-chapter, same-chapter ranges, and cross-chapter ranges.
        Results are cached.

        Raises:
            ValueError: If no API key is configured.
            requests.HTTPError: On API error.
        """
        if not self.has_api_key():
            raise ValueError(
                "No YouVersion API key configured. "
                "Please add your key in Settings → Bible Text Slides. "
                "Get a free key at https://developers.youversion.com"
            )
        return self._fetch_for_reference(reference, version_id)

    def get_verses_multi(
        self, references: List[BibleReference], version_id: int
    ) -> List[BibleVerse]:
        """Fetch verses for multiple references, renumbering sequentially."""
        result: List[BibleVerse] = []
        seq = 1
        for ref in references:
            for v in self.get_verses(ref, version_id):
                result.append(BibleVerse(verse_num=seq, text=v.text))
                seq += 1
        return result

    def get_youversion_url(self, reference: BibleReference, version_id: int) -> str:
        """Return the bible.com share URL for QR codes."""
        usfm = f"{reference.book_usfm}.{reference.chapter}"
        if not reference.whole_chapter:
            usfm += f".{reference.verse_start}"
            if reference.verse_end and reference.verse_end != reference.verse_start:
                usfm += f"-{reference.verse_end}"
        return f"{YOUVERSION_SHARE_BASE}/{version_id}/{usfm}"

    def clear_cache(self) -> None:
        self._passage_cache.clear()

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _get_headers(self) -> dict:
        h = dict(_HEADERS)
        h["x-yvp-app-key"] = self._api_key
        return h

    def _fetch_for_reference(
        self, reference: BibleReference, version_id: int
    ) -> List[BibleVerse]:
        """Fetch a passage, dispatching cross-chapter refs to per-chapter calls."""
        if reference.chapter_end and reference.chapter_end != reference.chapter:
            return self._fetch_cross_chapter(reference, version_id)
        return self._fetch_passage(reference.usfm_passage, reference.verse_start,
                                   reference.verse_end, reference.whole_chapter, version_id)

    def _fetch_cross_chapter(
        self, reference: BibleReference, version_id: int
    ) -> List[BibleVerse]:
        """Fetch a cross-chapter range via one API call per chapter.

        The YouVersion API does not support cross-chapter USFM ranges in a single
        call.  We therefore split the range into per-chapter requests using exact
        verse numbers from CHAPTER_VERSE_COUNTS so the API never receives an
        out-of-range verse number.
        """
        ch_start = reference.chapter
        ch_end = reference.chapter_end
        all_verses: List[BibleVerse] = []
        seq = 1

        for ch in range(ch_start, ch_end + 1):
            v_from = reference.verse_start if ch == ch_start else 1
            v_to = (
                reference.verse_end
                if ch == ch_end
                else _get_last_verse(reference.book_usfm, ch)
            )

            if v_from > v_to:
                logger.warning(
                    "verse_start %s > last_verse %s for %s.%s — skipping chapter",
                    v_from, v_to, reference.book_usfm, ch,
                )
                continue

            if v_from == v_to:
                usfm = f"{reference.book_usfm}.{ch}.{v_from}"
            else:
                usfm = f"{reference.book_usfm}.{ch}.{v_from}-{reference.book_usfm}.{ch}.{v_to}"

            try:
                chapter_verses = self._fetch_passage(usfm, v_from, v_to, False, version_id)
            except Exception as exc:
                logger.warning("Could not fetch %s (version %s): %s", usfm, version_id, exc)
                chapter_verses = []

            for v in chapter_verses:
                all_verses.append(BibleVerse(verse_num=seq, text=v.text))
                seq += 1

        return all_verses

    def _fetch_passage(
        self,
        usfm_passage: str,
        verse_start: int,
        verse_end: Optional[int],
        whole_chapter: bool,
        version_id: int,
    ) -> List[BibleVerse]:
        """Fetch a single passage from the Platform API (cached)."""
        cache_key = f"{version_id}:{usfm_passage}"
        if cache_key in self._passage_cache:
            return self._passage_cache[cache_key]

        url = f"{YOUVERSION_API_BASE}/bibles/{version_id}/passages/{usfm_passage}"
        logger.info("Fetching passage %s (version %s)", usfm_passage, version_id)

        try:
            resp = requests.get(url, headers=self._get_headers(), timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
        except requests.HTTPError as exc:
            logger.error(
                "API error for passage %s (version %s): %s – Response: %s",
                usfm_passage, version_id, exc,
                exc.response.text[:300] if exc.response is not None else "no body",
            )
            raise

        data = resp.json()
        verses = self._parse_passage(data, verse_start, verse_end, whole_chapter)
        self._passage_cache[cache_key] = verses
        return verses

    def _parse_passage(
        self,
        data: dict,
        verse_start: int,
        verse_end: Optional[int],
        whole_chapter: bool,
    ) -> List[BibleVerse]:
        """Parse a Platform API passage response into BibleVerse objects.

        The Platform API returns:
          {"id": "JHN.3.16", "content": "For God so loved...", "reference": "John 3:16"}

        For multi-verse passages the content includes inline verse numbers:
          "16 For God so loved... 17 For God did not send..."

        Parse strategies in order:
          1. Structured verse array in content/verses/items
          2. Split flat content text on inline verse numbers
          3. Return the whole content as a single verse block
        """
        payload = data.get("data", data) if isinstance(data, dict) else data
        if not isinstance(payload, dict):
            payload = data if isinstance(data, dict) else {}

        # --- Strategy 1: structured verse array ---
        verses = self._try_parse_structured(payload, verse_start, verse_end, whole_chapter)
        if verses:
            return verses

        # --- Strategy 2: split flat content text on inline verse numbers ---
        content = payload.get("content", "")
        if isinstance(content, str) and content.strip():
            verses = _split_content_on_verse_nums(
                _clean_text(content), verse_start, verse_end, whole_chapter
            )
            if verses:
                return verses

        # --- Strategy 3: return as single block ---
        cleaned = _clean_text(content) if isinstance(content, str) else ""
        if cleaned:
            logger.warning(
                "Returning passage as single block (could not split into individual verses)"
            )
            return [BibleVerse(verse_num=verse_start, text=cleaned)]

        logger.warning("No verse text found in response. Keys: %s", list(payload.keys()))
        return []

    @staticmethod
    def _try_parse_structured(
        payload: dict,
        verse_start: int,
        verse_end: Optional[int],
        whole_chapter: bool,
    ) -> List[BibleVerse]:
        """Try to extract verse objects from a structured response."""
        verses: List[BibleVerse] = []

        content_list = payload.get("content", payload.get("items", payload.get("verses", [])))
        if not isinstance(content_list, list):
            return []

        for item in content_list:
            if not isinstance(item, dict):
                continue
            verse_id = item.get("verse_id", item.get("verseId", item.get("id", "")))
            verse_num = _extract_verse_num(str(verse_id))
            if verse_num is None:
                verse_num = item.get("verse", item.get("verse_num"))
            if verse_num is None:
                continue
            try:
                verse_num = int(verse_num)
            except (ValueError, TypeError):
                continue
            text = _clean_text(item.get("content", item.get("text", "")))
            if text:
                verses.append(BibleVerse(verse_num=verse_num, text=text))

        if not verses:
            return []

        return _filter_by_range(verses, verse_start, verse_end, whole_chapter)

    def _fetch_bibles_list(self, language_tag: str) -> List[BibleTranslation]:
        """Fetch bibles from the Platform API, filtered by language.

        Tries parameter name variants in order:
          1. language_tag=<tag>
          2. language=<tag>
          3. No filter (client-side filtering)
        """
        # Try parameter variants: language_tag first, then language
        for param_name in ("language_tag", "language"):
            params: dict = {param_name: language_tag}
            try:
                resp = requests.get(
                    f"{YOUVERSION_API_BASE}/bibles",
                    headers=self._get_headers(),
                    params=params,
                    timeout=REQUEST_TIMEOUT,
                )
                if resp.status_code == 422:
                    logger.warning(
                        "Parameter '%s=%s' returned 422: %s",
                        param_name, language_tag, resp.text[:300],
                    )
                    continue  # try next variant
                resp.raise_for_status()
                results = self._parse_bibles_list(resp.json(), language_tag)
                logger.info(
                    "Fetched %d bibles for '%s' using param '%s'",
                    len(results), language_tag, param_name,
                )
                return results
            except requests.HTTPError as exc:
                logger.warning(
                    "HTTP error with %s=%s: %s – %s",
                    param_name, language_tag, exc,
                    exc.response.text[:300] if exc.response is not None else "no body",
                )
                continue

        # Fallback: fetch all bibles and filter client-side
        return self._fetch_bibles_list_all(language_tag)

    def _fetch_bibles_list_all(self, language_tag: str) -> List[BibleTranslation]:
        """Fetch all bibles (no language filter) and filter client-side."""
        try:
            resp = requests.get(
                f"{YOUVERSION_API_BASE}/bibles",
                headers=self._get_headers(),
                timeout=REQUEST_TIMEOUT,
            )
            if not resp.ok:
                logger.warning(
                    "GET /bibles (no params) returned %s: %s",
                    resp.status_code, resp.text[:400],
                )
                resp.raise_for_status()
        except requests.HTTPError:
            raise

        return self._parse_bibles_list(resp.json(), language_tag)

    @staticmethod
    def _parse_bibles_list(data: dict, language_filter: str) -> List[BibleTranslation]:
        """Parse a /bibles API response into BibleTranslation objects."""
        bibles = data.get("data", data) if isinstance(data, dict) else data
        if not isinstance(bibles, list):
            return []

        results: List[BibleTranslation] = []
        for b in bibles:
            if not isinstance(b, dict):
                continue
            lang_info = b.get("language", {})
            lang_tag = lang_info.get("tag", "") if isinstance(lang_info, dict) else ""
            if language_filter and not lang_tag.startswith(language_filter):
                continue
            lang_name = lang_info.get("name", lang_tag) if isinstance(lang_info, dict) else lang_tag
            bid = b.get("id")
            if bid is None:
                continue
            results.append(BibleTranslation(
                id=int(bid),
                abbreviation=b.get("abbreviation", str(bid)),
                name=b.get("title", b.get("local_title", str(bid))),
                language=lang_tag,
                language_name=lang_name,
            ))
        return results


# ---------------------------------------------------------------------------
# Helper: last verse lookup
# ---------------------------------------------------------------------------

def _get_last_verse(book_usfm: str, chapter: int) -> int:
    """Return the last verse number for a chapter using CHAPTER_VERSE_COUNTS.

    Falls back to 176 (Psalm 119, the longest chapter in any Protestant Bible)
    if the book/chapter is not in the table.  A fallback of 176 will cause a
    404 from the API for most chapters, but is safer than an arbitrary low cap.
    """
    counts = CHAPTER_VERSE_COUNTS.get(book_usfm)
    if counts and 1 <= chapter <= len(counts):
        return counts[chapter - 1]
    logger.warning(
        "No verse count for %s ch %s in CHAPTER_VERSE_COUNTS; using fallback 176",
        book_usfm, chapter,
    )
    return 176


# ---------------------------------------------------------------------------
# Text utilities
# ---------------------------------------------------------------------------

_HTML_TAG_RE = re.compile(r"<[^>]+>")
_WHITESPACE_RE = re.compile(r"\s+")
# Verse number embedded inline: "16 text", not part of a longer word
_INLINE_VERSE_RE = re.compile(r"(?<!\w)(\d{1,3})(?=\s)")


def _clean_text(text: str) -> str:
    text = _HTML_TAG_RE.sub(" ", text)
    text = _WHITESPACE_RE.sub(" ", text)
    return text.strip()


def _extract_verse_num(verse_id: str) -> Optional[int]:
    parts = verse_id.split(".")
    if len(parts) >= 3:
        try:
            return int(parts[-1])
        except ValueError:
            pass
    return None


def _split_content_on_verse_nums(
    content: str,
    verse_start: int,
    verse_end: Optional[int],
    whole_chapter: bool,
) -> List[BibleVerse]:
    """Split a flat content string on inline verse numbers."""
    matches = list(_INLINE_VERSE_RE.finditer(content))
    if not matches:
        return []

    end = verse_end if not whole_chapter and verse_end is not None else 999
    relevant = [m for m in matches if verse_start <= int(m.group(1)) <= end]
    if not relevant:
        relevant = [m for m in matches if 1 <= int(m.group(1)) <= 200]
        if len(relevant) < 2:
            return []

    verses: List[BibleVerse] = []
    for i, m in enumerate(relevant):
        verse_num = int(m.group(1))
        text_start = m.end()
        text_end = relevant[i + 1].start() if i + 1 < len(relevant) else len(content)
        text = content[text_start:text_end].strip()
        if text:
            verses.append(BibleVerse(verse_num=verse_num, text=text))

    return _filter_by_range(verses, verse_start, verse_end, whole_chapter)


def _filter_by_range(
    verses: List[BibleVerse],
    verse_start: int,
    verse_end: Optional[int],
    whole_chapter: bool,
) -> List[BibleVerse]:
    if whole_chapter or verse_end is None:
        return [v for v in verses if v.verse_num >= verse_start]
    return [v for v in verses if verse_start <= v.verse_num <= verse_end]
