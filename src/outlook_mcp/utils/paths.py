"""Path validation for attachment / output-dir parameters.

The MCP server runs with the user's full privileges. Tools that take
filesystem paths from the model are validated here so a confused or
adversarial agent can't read or write arbitrary locations.

Set ``OUTLOOK_MCP_ALLOW_ANY_PATH=1`` to disable the user-profile sandbox
(handy if you legitimately need to attach a file from a network share).
"""

from __future__ import annotations

import os
from pathlib import Path

from outlook_mcp.errors import OutlookError


def _profile_root() -> Path:
    return Path(os.environ.get("USERPROFILE") or Path.home()).resolve()


def _sandbox_disabled() -> bool:
    return os.environ.get("OUTLOOK_MCP_ALLOW_ANY_PATH", "").strip() in ("1", "true", "yes")


def _within_profile(path: Path) -> bool:
    try:
        path.resolve().relative_to(_profile_root())
    except ValueError:
        return False
    return True


def validate_attachment_path(raw: str) -> str:
    """Validate that ``raw`` points to an existing file under the user profile.

    Returns the normalized absolute path. Raises :class:`OutlookError`
    on rejection.
    """
    if not raw:
        raise OutlookError("Attachment path is empty.")
    p = Path(raw)
    if not p.is_absolute():
        raise OutlookError(
            f"Attachment path '{raw}' must be absolute "
            f"(e.g. C:\\Users\\you\\Documents\\file.pdf)."
        )
    if not p.is_file():
        raise OutlookError(
            f"Attachment not found: '{raw}'. Provide an absolute path to a "
            "file that exists on this machine."
        )
    if not _sandbox_disabled() and not _within_profile(p):
        raise OutlookError(
            f"Attachment '{raw}' is outside the user profile. Set the env "
            "var OUTLOOK_MCP_ALLOW_ANY_PATH=1 to allow paths anywhere."
        )
    return str(p.resolve())


def validate_output_dir(raw: str) -> str:
    """Validate (and create) an output directory under the user profile."""
    if not raw:
        raise OutlookError("Output directory path is empty.")
    p = Path(raw)
    if not p.is_absolute():
        raise OutlookError(
            f"Output directory '{raw}' must be absolute."
        )
    if not _sandbox_disabled() and not _within_profile(p):
        raise OutlookError(
            f"Output directory '{raw}' is outside the user profile. Set "
            "OUTLOOK_MCP_ALLOW_ANY_PATH=1 to allow arbitrary paths."
        )
    p.mkdir(parents=True, exist_ok=True)
    return str(p.resolve())
