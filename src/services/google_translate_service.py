"""Google Translate helper using the free unofficial API.

Uses the same endpoint that many lightweight translation tools rely on.
No API key is required but usage should be moderate (this is unofficial).
"""

import re
from typing import List, Optional
from urllib.parse import quote

import requests

from ..logging_config import get_logger

logger = get_logger("google_translate_service")

_ENDPOINT = "https://translate.googleapis.com/translate_a/single"
_TIMEOUT = 10

# ISO 639-1 codes for RTL languages
RTL_LANGUAGES = {"ar", "he", "fa", "ur", "yi", "dv", "ps"}

# Human-readable target language options shown in the picker
TRANSLATE_LANGUAGES = [
    ("nl", "Nederlands"),
    ("en", "English"),
    ("de", "Deutsch"),
    ("fr", "Français"),
    ("es", "Español"),
    ("pt", "Português"),
    ("it", "Italiano"),
    ("ru", "Русский"),
    ("pl", "Polski"),
    ("ro", "Română"),
    ("cs", "Čeština"),
    ("hu", "Magyar"),
    ("sv", "Svenska"),
    ("da", "Dansk"),
    ("no", "Norsk"),
    ("fi", "Suomi"),
    ("tr", "Türkçe"),
    ("id", "Indonesian"),
    ("zh-CN", "中文 (简体)"),
    ("zh-TW", "中文 (繁體)"),
    ("ja", "日本語"),
    ("ko", "한국어"),
    ("ar", "العربية"),
    ("he", "עברית"),
    ("fa", "فارسی"),
]


def translate(text: str, target_lang: str, source_lang: str = "auto") -> str:
    """Translate *text* to *target_lang* using Google Translate's free endpoint.

    Args:
        text: Source text (plain, no HTML).
        target_lang: Target ISO 639-1 language code (e.g. "nl", "en", "de").
        source_lang: Source language code, or "auto" for detection.

    Returns:
        Translated string.

    Raises:
        requests.RequestException: On network failure.
        ValueError: If the response format is unexpected.
    """
    if not text.strip():
        return text

    params = {
        "client": "gtx",
        "sl": source_lang,
        "tl": target_lang,
        "dt": "t",
        "q": text,
    }

    resp = requests.get(_ENDPOINT, params=params, timeout=_TIMEOUT)
    resp.raise_for_status()

    data = resp.json()
    # Response: [[["translated_chunk", "original", ...], ...], ...]
    if not isinstance(data, list) or not data:
        raise ValueError(f"Unexpected Google Translate response: {data!r}")

    sentences = data[0]
    if not isinstance(sentences, list):
        raise ValueError(f"Unexpected sentence list: {sentences!r}")

    parts = []
    for item in sentences:
        if isinstance(item, list) and item:
            chunk = item[0]
            if isinstance(chunk, str):
                parts.append(chunk)

    return "".join(parts)


def translate_batch(texts: List[str], target_lang: str) -> List[str]:
    """Translate multiple texts to target_lang, returning a parallel list.

    Translates each text individually.  Failed translations fall back to the
    original text.
    """
    results = []
    for t in texts:
        try:
            results.append(translate(t, target_lang))
        except Exception as exc:
            logger.warning("Translation failed for a text chunk: %s", exc)
            results.append(t)
    return results


def is_rtl(language_code: str) -> bool:
    """Return True if *language_code* corresponds to a right-to-left language."""
    base = language_code.split("-")[0].lower()
    return base in RTL_LANGUAGES
