"""Service for YouTube search and link validation."""

import os
import re
import subprocess
import sys
import json
from typing import List, Optional, Tuple
from dataclasses import dataclass
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests


@dataclass
class YouTubeResult:
    """Represents a YouTube search result."""

    title: str
    url: str
    channel: str
    duration: str
    thumbnail_url: Optional[str] = None


class YouTubeService:
    """Service for YouTube operations."""

    # Regex to extract video ID from various YouTube URL formats
    YOUTUBE_REGEX = re.compile(
        r"(?:https?://)?(?:www\.)?(?:youtube\.com/watch\?v=|youtu\.be/|youtube\.com/embed/)([a-zA-Z0-9_-]{11})"
    )

    def __init__(self):
        self._yt_dlp_available = None

    def _get_yt_dlp_cmd(self) -> List[str]:
        """Get the command to run yt-dlp as a Python module."""
        return [sys.executable, "-m", "yt_dlp"]

    def is_yt_dlp_available(self) -> bool:
        """Check if yt-dlp is available."""
        if self._yt_dlp_available is None:
            try:
                # Try running as Python module first (works after pip install)
                result = subprocess.run(
                    self._get_yt_dlp_cmd() + ["--version"],
                    capture_output=True,
                    text=True,
                    timeout=5,
                )
                self._yt_dlp_available = result.returncode == 0
            except (subprocess.SubprocessError, FileNotFoundError):
                self._yt_dlp_available = False
        return self._yt_dlp_available

    def search(self, query: str, max_results: int = 5) -> List[YouTubeResult]:
        """
        Search YouTube for videos matching the query.
        Uses yt-dlp for searching.
        """
        if not self.is_yt_dlp_available():
            return []

        try:
            # Use yt-dlp to search (as Python module)
            result = subprocess.run(
                self._get_yt_dlp_cmd() + [
                    f"ytsearch{max_results}:{query}",
                    "--dump-json",
                    "--flat-playlist",
                    "--no-warnings",
                ],
                capture_output=True,
                text=True,
                timeout=30,
            )

            if result.returncode != 0:
                return []

            results = []
            for line in result.stdout.strip().split("\n"):
                if line:
                    try:
                        data = json.loads(line)
                        results.append(
                            YouTubeResult(
                                title=data.get("title", "Unknown"),
                                url=f"https://www.youtube.com/watch?v={data.get('id', '')}",
                                channel=data.get("channel", data.get("uploader", "Unknown")),
                                duration=self._format_duration(data.get("duration")),
                                thumbnail_url=data.get("thumbnail"),
                            )
                        )
                    except json.JSONDecodeError:
                        continue

            return results

        except subprocess.SubprocessError:
            return []

    def _format_duration(self, seconds: Optional[int]) -> str:
        """Format duration in seconds to MM:SS or HH:MM:SS."""
        if seconds is None:
            return "?"
        hours, remainder = divmod(int(seconds), 3600)
        minutes, secs = divmod(remainder, 60)
        if hours > 0:
            return f"{hours}:{minutes:02d}:{secs:02d}"
        return f"{minutes}:{secs:02d}"

    def extract_video_id(self, url: str) -> Optional[str]:
        """Extract video ID from a YouTube URL."""
        match = self.YOUTUBE_REGEX.search(url)
        if match:
            return match.group(1)
        return None

    def validate_link_fast(self, url: str) -> bool:
        """
        Fast validation using HEAD request.
        Returns True if the URL appears valid.
        """
        video_id = self.extract_video_id(url)
        if not video_id:
            return False

        try:
            # Check if the video page returns 200
            response = requests.head(
                f"https://www.youtube.com/watch?v={video_id}",
                timeout=5,
                allow_redirects=True,
            )
            return response.status_code == 200
        except requests.RequestException:
            return False

    def validate_link_thorough(self, url: str) -> Tuple[bool, Optional[str]]:
        """
        Thorough validation using yt-dlp.
        Returns (is_valid, error_message).
        """
        video_id = self.extract_video_id(url)
        if not video_id:
            return False, "Invalid YouTube URL"

        if not self.is_yt_dlp_available():
            # Fall back to fast validation
            return self.validate_link_fast(url), None

        try:
            result = subprocess.run(
                self._get_yt_dlp_cmd() + [
                    "--simulate",
                    "--no-warnings",
                    f"https://www.youtube.com/watch?v={video_id}",
                ],
                capture_output=True,
                text=True,
                timeout=15,
            )

            if result.returncode == 0:
                return True, None
            else:
                # Extract error message
                error = result.stderr.strip()
                if "Video unavailable" in error:
                    return False, "Video unavailable"
                elif "Private video" in error:
                    return False, "Private video"
                elif "removed" in error.lower():
                    return False, "Video removed"
                return False, error[:100] if error else "Unknown error"

        except subprocess.TimeoutExpired:
            return False, "Timeout"
        except subprocess.SubprocessError as e:
            return False, str(e)

    def validate_links_batch(
        self, urls: List[str], thorough: bool = False
    ) -> List[Tuple[str, bool, Optional[str]]]:
        """
        Validate multiple links in parallel.
        Returns list of (url, is_valid, error_message) tuples.
        """
        results = []

        validate_func = (
            self.validate_link_thorough if thorough else self._validate_fast_wrapper
        )

        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {executor.submit(validate_func, url): url for url in urls}

            for future in as_completed(futures):
                url = futures[future]
                try:
                    result = future.result()
                    if isinstance(result, tuple):
                        is_valid, error = result
                    else:
                        is_valid, error = result, None
                    results.append((url, is_valid, error))
                except Exception as e:
                    results.append((url, False, str(e)))

        return results

    def _validate_fast_wrapper(self, url: str) -> Tuple[bool, Optional[str]]:
        """Wrapper to make fast validation return same format as thorough."""
        is_valid = self.validate_link_fast(url)
        return is_valid, None if is_valid else "Link check failed"

    def read_youtube_file(self, song_folder: str) -> List[str]:
        """Read YouTube links from a song folder's youtube.txt."""
        youtube_file = os.path.join(song_folder, "youtube.txt")
        if os.path.exists(youtube_file):
            try:
                with open(youtube_file, "r", encoding="utf-8") as f:
                    return [line.strip() for line in f if line.strip()]
            except IOError:
                pass
        return []

    def write_youtube_file(self, song_folder: str, urls: List[str]) -> None:
        """Write YouTube links to a song folder's youtube.txt."""
        youtube_file = os.path.join(song_folder, "youtube.txt")
        with open(youtube_file, "w", encoding="utf-8") as f:
            for url in urls:
                f.write(f"{url}\n")
