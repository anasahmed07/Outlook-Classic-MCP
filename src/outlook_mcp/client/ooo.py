"""Out-of-Office (auto-reply) COM operations.

Reads the legacy MAPI property ``PR_OOF_STATE`` (proptag
``0x661D000B``). The property is read-only via PropertyAccessor on most
Exchange profiles; toggling OOO requires EWS or Graph and isn't done
here.
"""

from __future__ import annotations

from typing import Any

from outlook_mcp.constants import OOF_PROPTAG


def get_out_of_office(outlook: Any, namespace: Any) -> dict[str, Any]:
    store = namespace.DefaultStore
    try:
        state = store.PropertyAccessor.GetProperty(OOF_PROPTAG)
    except Exception:
        return {
            "out_of_office": None,
            "status": "unknown",
            "note": "Could not read OOF property. Check Outlook File > Automatic Replies directly.",
        }
    return {"out_of_office": bool(state), "status": "on" if state else "off"}
