"""Liturgie Samensteller - PowerPoint Mixer application."""

import subprocess
import os

__version__ = "1.1.17"


def _get_git_revision() -> str:
    """Get the current git revision hash."""
    try:
        # Get the directory where this file is located
        pkg_dir = os.path.dirname(os.path.abspath(__file__))
        result = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=pkg_dir,
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return "unknown"


def _get_git_date() -> str:
    """Get the date of the current git commit."""
    try:
        pkg_dir = os.path.dirname(os.path.abspath(__file__))
        result = subprocess.run(
            ["git", "log", "-1", "--format=%cd", "--date=short"],
            cwd=pkg_dir,
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return "unknown"


__revision__ = _get_git_revision()
__build_date__ = _get_git_date()
