"""Account / metadata COM operations."""

from __future__ import annotations

import datetime as dt
from typing import Any

from outlook_mcp.client.folders import _safe_get


def whoami(outlook: Any, namespace: Any) -> dict[str, Any]:
    accounts = []
    for acct in outlook.Session.Accounts:
        accounts.append(
            {
                "display_name": _safe_get(acct, "DisplayName"),
                "smtp_address": _safe_get(acct, "SmtpAddress"),
                "user_name": _safe_get(acct, "UserName"),
                "account_type": _safe_get(acct, "AccountType"),
            }
        )
    # All datetimes this server returns are in this (the user's) local
    # timezone with an explicit offset — surface it so agents never guess.
    now = dt.datetime.now().astimezone()
    return {
        "current_user": _safe_get(namespace.CurrentUser, "Name"),
        "accounts": accounts,
        "local_time": now.isoformat(),
        "timezone": str(now.tzinfo),
        "utc_offset": now.strftime("%z"),
    }
