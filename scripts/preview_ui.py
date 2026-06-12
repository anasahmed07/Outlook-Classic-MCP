"""Preview the MCP Apps views in a normal browser, no MCP host needed.

Renders every ``ui://`` view into ``.ui-preview/`` next to a fake-host
harness (``scripts/ui_harness.html``) that answers the ``ui/initialize``
handshake, pushes sample data, and stubs ``tools/call``.

    python scripts/preview_ui.py
    start .ui-preview/harness.html?view=mail-list          (or any view)

Views: mail-list, mail-view, calendar, contacts, tasks.
Add ``&theme=dark`` to preview the dark palette.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from outlook_mcp import ui  # noqa: E402

OUT = ROOT / ".ui-preview"


def main() -> None:
    OUT.mkdir(exist_ok=True)
    for view in ui._VIEWS:
        (OUT / f"{view}.html").write_text(ui._render(view), encoding="utf-8")
        print(f"  {view}.html")
    shutil.copy(ROOT / "scripts" / "ui_harness.html", OUT / "harness.html")
    print(f"\nOpen {OUT / 'harness.html'}?view=mail-list in a browser.")


if __name__ == "__main__":
    main()
