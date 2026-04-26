"""Tool registration: each submodule exposes register(mcp, bridge)."""

from . import (
    account,
    calendar,
    categories,
    contacts,
    folders,
    mail,
    ooo,
    rules,
    tasks,
)


def register_all(mcp, bridge) -> None:
    for mod in (mail, folders, calendar, contacts, tasks, categories, rules, ooo, account):
        mod.register(mcp, bridge)
