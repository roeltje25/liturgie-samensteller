"""Service for sharing text via the Windows Share dialog."""

import subprocess
import tempfile
import os
import sys

from ..logging_config import get_logger

logger = get_logger("share_service")


def _find_share_helper() -> str | None:
    """Find the ShareHelper.exe binary.

    Checks:
    1. Next to the running executable (PyInstaller bundle)
    2. In src/resources/ (development)
    """
    # When frozen with PyInstaller, look next to the exe
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        path = os.path.join(exe_dir, "ShareHelper.exe")
        if os.path.isfile(path):
            return path

    # Development: look in src/resources/
    src_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(src_dir, "resources", "ShareHelper.exe")
    if os.path.isfile(path):
        return path

    return None


def share_text(text: str, title: str, hwnd: int) -> bool:
    """Show the Windows Share dialog with the given text.

    Uses a compiled C# helper (ShareHelper.exe) that creates a WinForms
    window and opens the Windows 10/11 Share dialog via
    IDataTransferManagerInterop. Requires .NET 9 runtime on the machine.

    Args:
        text: The text content to share.
        title: The title shown in the share dialog.
        hwnd: The window handle (unused; ShareHelper creates its own).

    Returns:
        True if the share dialog was opened successfully, False otherwise.
    """
    helper = _find_share_helper()
    if not helper:
        logger.warning("ShareHelper.exe not found")
        return False

    tmp_file = None
    try:
        # Write text to a temp file (avoids command-line escaping issues)
        tmp_file = tempfile.NamedTemporaryFile(
            mode="w", suffix=".txt", encoding="utf-8", delete=False
        )
        tmp_file.write(text)
        tmp_file.close()

        result = subprocess.run(
            [helper, "--title", title, "--file", tmp_file.name],
            capture_output=True,
            timeout=90,
        )

        if result.returncode == 0:
            logger.info("Share dialog opened successfully")
            return True
        else:
            logger.warning("Share dialog failed (rc=%d)", result.returncode)
            return False

    except subprocess.TimeoutExpired:
        logger.warning("Share dialog timed out")
        return False
    except Exception:
        logger.error("Failed to open share dialog", exc_info=True)
        return False
    finally:
        if tmp_file is not None:
            try:
                os.unlink(tmp_file.name)
            except OSError:
                pass
