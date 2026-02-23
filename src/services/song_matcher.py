"""Fuzzy song matching utilities shared across the UI layer."""

import re
import unicodedata
from typing import List, Tuple

from ..models import Song


def normalize_for_search(text: str) -> str:
    """Normalize text for fuzzy multilingual search.

    Handles phonetic equivalences common in Dutch/German/English/Arabic songs:
    - y ↔ j (ya/ja, Yeshua/Jeshua)
    - i ↔ ee ↔ ie (Jesus/Jezus)
    - c ↔ k (Christ/Krist)
    - ou ↔ oe ↔ u (Dutch sounds)
    - ei ↔ ij ↔ y (Dutch)
    - ph ↔ f, th ↔ t
    - sch ↔ s
    - Removes diacritics (é→e, ü→u, etc.)
    """
    # Lowercase
    text = text.lower()

    # Remove diacritics (é→e, ü→u, ñ→n, etc.)
    text = unicodedata.normalize('NFD', text)
    text = ''.join(c for c in text if unicodedata.category(c) != 'Mn')

    # Multi-character replacements (order matters - do longer patterns first)
    replacements = [
        ('sch', 's'),      # German/Dutch sch → s
        ('gh', 'g'),       # Arabic transliteration
        ('ch', 'k'),       # Christ → Krist
        ('ph', 'f'),       # Pharao → Farao
        ('th', 't'),       # Thomas → Tomas
        ('oe', 'u'),       # Dutch oe → u sound
        ('ou', 'u'),       # French/Dutch ou → u
        ('ee', 'i'),       # ee → i (Geest → Gist)
        ('ie', 'i'),       # ie → i
        ('ei', 'y'),       # Dutch ei → y
        ('ij', 'y'),       # Dutch ij → y
        ('aa', 'a'),       # Double vowels → single
        ('oo', 'o'),
        ('uu', 'u'),
    ]

    for old, new in replacements:
        text = text.replace(old, new)

    # Single character replacements
    char_map = str.maketrans({
        'j': 'y',          # j → y (Dutch j sounds like English y)
        'c': 'k',          # c → k
        'q': 'k',          # q → k
        'x': 'ks',         # x → ks
        'z': 's',          # z → s (soften)
        'v': 'f',          # v → f (Dutch v often sounds like f)
        'w': 'v',          # w → v (German w)
    })
    text = text.translate(char_map)

    # Remove non-alphanumeric
    text = re.sub(r'[^a-z0-9]', '', text)

    return text


def fuzzy_match_score(query: str, text: str) -> float:
    """Calculate fuzzy match score between query and text.

    Returns a score from 0.0 (no match) to 1.0 (perfect match).
    Uses normalized forms for phonetic matching.
    """
    if not query:
        return 1.0

    query_norm = normalize_for_search(query)
    text_norm = normalize_for_search(text)

    if not query_norm:
        return 1.0

    # Exact normalized substring match - best score
    if query_norm in text_norm:
        return 1.0

    # Check if all query characters appear in order (subsequence match)
    # This handles typos and partial matches
    query_idx = 0
    matches = 0
    last_match_pos = -1
    consecutive_bonus = 0

    for i, char in enumerate(text_norm):
        if query_idx < len(query_norm) and char == query_norm[query_idx]:
            matches += 1
            # Bonus for consecutive matches
            if last_match_pos == i - 1:
                consecutive_bonus += 0.1
            last_match_pos = i
            query_idx += 1

    if matches == len(query_norm):
        # All characters found in order - good match
        base_score = 0.6
        score = min(1.0, base_score + consecutive_bonus)
        return score

    # Partial match - some characters found
    if matches > 0:
        return 0.3 * (matches / len(query_norm))

    return 0.0


def find_best_matches(
    query: str, songs: List[Song], limit: int = 3
) -> List[Tuple[Song, float]]:
    """Return up to `limit` (song, score) pairs, sorted descending by score."""
    scored = []
    for song in songs:
        score = max(
            fuzzy_match_score(query, song.display_title),
            fuzzy_match_score(query, song.name),
        )
        if score > 0:
            scored.append((song, score))
    scored.sort(key=lambda x: x[1], reverse=True)
    return scored[:limit]
