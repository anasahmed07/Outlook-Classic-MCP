"""Detect installed MCP clients and merge the outlook server into their configs.

Run after `install.bat` (or any time you want to register the server with
additional clients). The script:

1. Probes a fixed list of MCP-aware editors / desktop apps for their
   config file (or, in Claude Code's case, the `claude` CLI).
2. Lists every detected client in a checkbox-style menu.
3. For each toggled-on client, snapshots the existing config to
   ``<file>.bak`` and deep-merges an ``mcpServers.outlook`` entry that
   points at the venv Python and ``-m outlook_mcp``.

Re-running is safe: the merge updates an existing entry instead of
duplicating it.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional


HERE = Path(__file__).resolve().parent
PROJECT_ROOT = HERE.parent
DEFAULT_VENV_PY = PROJECT_ROOT / ".venv" / "Scripts" / "python.exe"


def _venv_python() -> str:
    explicit = os.environ.get("OUTLOOK_MCP_PYTHON")
    if explicit:
        return explicit
    if DEFAULT_VENV_PY.exists():
        return str(DEFAULT_VENV_PY)
    # Fall back to current interpreter (the script may have been run by
    # the venv directly, in which case sys.executable is correct).
    return sys.executable


SERVER_ENTRY = {
    "command": _venv_python(),
    "args": ["-m", "outlook_mcp"],
}
SERVER_KEY = "outlook"


# ---------------------------------------------------------------------------
# Per-client adapters
# ---------------------------------------------------------------------------


@dataclass
class Client:
    name: str
    config_path: Optional[Path]
    install_fn: Callable[["Client"], str]
    detect_fn: Callable[[], bool]


def _path_join_env(env_var: str, *parts: str) -> Optional[Path]:
    base = os.environ.get(env_var)
    if not base:
        return None
    return Path(base, *parts)


def _claude_desktop_path() -> Optional[Path]:
    return _path_join_env("APPDATA", "Claude", "claude_desktop_config.json")


def _cursor_path() -> Optional[Path]:
    return _path_join_env("USERPROFILE", ".cursor", "mcp.json")


def _cline_path() -> Optional[Path]:
    return _path_join_env(
        "APPDATA",
        "Code",
        "User",
        "globalStorage",
        "saoudrizwan.claude-dev",
        "settings",
        "cline_mcp_settings.json",
    )


def _continue_path() -> Optional[Path]:
    return _path_join_env("USERPROFILE", ".continue", "config.json")


def _windsurf_path() -> Optional[Path]:
    return _path_join_env("USERPROFILE", ".codeium", "windsurf", "mcp_config.json")


def _claude_cli_present() -> bool:
    return shutil.which("claude") is not None


def _exists(p: Optional[Path]) -> bool:
    return p is not None and p.exists()


# Each "install" returns a one-line status string ----------------------------


def _install_via_json_merge(
    config_path: Path, schema_path: tuple[str, ...] = ("mcpServers",)
) -> str:
    """Merge SERVER_ENTRY into config_path under schema_path[*][SERVER_KEY].

    Creates the file (and parent dirs) if missing. Snapshots existing
    file to <file>.bak before writing.
    """
    config_path.parent.mkdir(parents=True, exist_ok=True)
    if config_path.exists():
        try:
            existing = json.loads(config_path.read_text(encoding="utf-8") or "{}")
        except json.JSONDecodeError:
            return f"  ! {config_path} is not valid JSON; aborted"
        backup = config_path.with_suffix(config_path.suffix + ".bak")
        shutil.copyfile(config_path, backup)
    else:
        existing = {}

    cursor = existing
    for key in schema_path:
        if not isinstance(cursor.get(key), dict):
            cursor[key] = {}
        cursor = cursor[key]
    cursor[SERVER_KEY] = SERVER_ENTRY

    config_path.write_text(json.dumps(existing, indent=2), encoding="utf-8")
    return f"  - wrote {config_path}"


def _install_claude_desktop(client: Client) -> str:
    return _install_via_json_merge(client.config_path)  # type: ignore[arg-type]


def _install_cursor(client: Client) -> str:
    return _install_via_json_merge(client.config_path)  # type: ignore[arg-type]


def _install_cline(client: Client) -> str:
    return _install_via_json_merge(client.config_path)  # type: ignore[arg-type]


def _install_continue(client: Client) -> str:
    # Continue nests MCP servers under experimental.modelContextProtocolServers.
    return _install_via_json_merge(
        client.config_path,  # type: ignore[arg-type]
        schema_path=("experimental", "modelContextProtocolServers"),
    )


def _install_windsurf(client: Client) -> str:
    return _install_via_json_merge(client.config_path)  # type: ignore[arg-type]


def _install_claude_code(client: Client) -> str:
    cmd = [
        "claude",
        "mcp",
        "add",
        SERVER_KEY,
        "--",
        SERVER_ENTRY["command"],
        *SERVER_ENTRY["args"],
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        # Likely an "already exists" — try `mcp remove` then re-add once.
        subprocess.run(["claude", "mcp", "remove", SERVER_KEY], capture_output=True)
        proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        return f"  ! claude mcp add failed: {(proc.stderr or proc.stdout).strip()}"
    return "  - registered with `claude mcp add`"


# ---------------------------------------------------------------------------
# Client registry
# ---------------------------------------------------------------------------


def discover() -> list[Client]:
    candidates: list[Client] = [
        Client(
            "Claude Desktop",
            _claude_desktop_path(),
            _install_claude_desktop,
            lambda: _exists(_claude_desktop_path()),
        ),
        Client(
            "Claude Code",
            None,
            _install_claude_code,
            _claude_cli_present,
        ),
        Client(
            "Cursor",
            _cursor_path(),
            _install_cursor,
            lambda: _exists(_cursor_path()),
        ),
        Client(
            "Cline (VS Code)",
            _cline_path(),
            _install_cline,
            lambda: _exists(_cline_path()),
        ),
        Client(
            "Continue (VS Code)",
            _continue_path(),
            _install_continue,
            lambda: _exists(_continue_path()),
        ),
        Client(
            "Windsurf",
            _windsurf_path(),
            _install_windsurf,
            lambda: _exists(_windsurf_path()),
        ),
    ]
    return [c for c in candidates if c.detect_fn()]


# ---------------------------------------------------------------------------
# Checkbox menu
# ---------------------------------------------------------------------------


def _print_menu(clients: list[Client], selected: set[int]) -> None:
    print()
    print("Select which clients to register outlook-mcp with:")
    for i, c in enumerate(clients, start=1):
        mark = "[x]" if i in selected else "[ ]"
        loc = c.config_path or "(via `claude` CLI)"
        print(f"  {mark} {i}. {c.name:<20} {loc}")
    print()
    print("Type a number to toggle, 'a' to select all, 'n' for none,")
    print("'enter' to confirm, 'q' to quit without changes.")


def _prompt_loop(clients: list[Client]) -> set[int]:
    # Default: nothing pre-selected — user must opt in to writing configs.
    selected: set[int] = set()
    while True:
        _print_menu(clients, selected)
        choice = input("> ").strip().lower()
        if choice == "":
            return selected
        if choice == "q":
            return set()
        if choice == "a":
            selected = set(range(1, len(clients) + 1))
            continue
        if choice == "n":
            selected = set()
            continue
        if choice.isdigit():
            idx = int(choice)
            if 1 <= idx <= len(clients):
                selected.symmetric_difference_update({idx})
            continue
        print(f"unrecognized: {choice!r}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> int:
    print("=" * 72)
    print(" outlook-mcp client installer")
    print("=" * 72)
    print(f" Python: {SERVER_ENTRY['command']}")
    print(f" Server: python -m outlook_mcp")

    clients = discover()
    if not clients:
        print()
        print("No MCP clients detected on this machine.")
        print("Install one of the following, then re-run this script:")
        print("  - Claude Desktop: https://claude.ai/download")
        print("  - Claude Code:    https://github.com/anthropics/claude-code")
        print("  - Cursor:         https://cursor.com")
        return 0

    selected = _prompt_loop(clients)
    if not selected:
        print("\nNo clients selected. Nothing changed.")
        return 0

    print()
    print("Installing ...")
    failures = 0
    for idx in sorted(selected):
        client = clients[idx - 1]
        print(f"* {client.name}:")
        try:
            msg = client.install_fn(client)
            print(msg)
            if msg.lstrip().startswith("!"):
                failures += 1
        except Exception as exc:  # noqa: BLE001
            print(f"  ! failed: {exc}")
            failures += 1

    print()
    if failures:
        print(f"Done with {failures} failure(s). Inspect output above.")
    else:
        print("Done. Restart each affected client to pick up the new tools.")
    return 0 if failures == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
