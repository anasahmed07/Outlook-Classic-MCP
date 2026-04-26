"""Mail-rule COM operations.

Rules live on the default store. ``GetRules()`` returns a Rules
collection; ``Save()`` on the collection persists changes.
"""

from __future__ import annotations

from typing import Any

from outlook_mcp.errors import OutlookError


def list_rules(outlook: Any, namespace: Any) -> dict[str, Any]:
    rules = namespace.DefaultStore.GetRules()
    items = []
    for i in range(rules.Count):
        rule = rules.Item(i + 1)
        items.append(
            {
                "index": i + 1,
                "name": rule.Name,
                "enabled": bool(rule.Enabled),
            }
        )
    return {"count": len(items), "items": items}


def toggle_rule(
    outlook: Any,
    namespace: Any,
    *,
    rule_name: str,
    enabled: bool,
) -> dict[str, Any]:
    rules = namespace.DefaultStore.GetRules()
    for i in range(rules.Count):
        rule = rules.Item(i + 1)
        if rule.Name == rule_name:
            rule.Enabled = enabled
            rules.Save()
            return {
                "status": "updated",
                "rule": rule_name,
                "enabled": bool(enabled),
            }
    raise OutlookError(
        f"Rule '{rule_name}' not found. Use outlook_list_rules to see available rules."
    )
