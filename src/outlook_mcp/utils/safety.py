"""Safety helpers: DASL escaping and error-handling decorator."""

from __future__ import annotations

import functools
import logging
from typing import Any, Callable

from outlook_mcp.errors import OutlookError, format_com_error

logger = logging.getLogger("outlook_mcp.safety")


def safe_dasl(query: str) -> str:
    """Escape a string for safe inclusion in a DASL ``LIKE`` value.

    DASL uses ``%`` and ``_`` as SQL wildcards and pairs of single/double
    quotes for escaping. Without this, a search for ``50%_off`` would
    silently widen to a wildcard search.
    """
    if query is None:
        return ""
    query = query.replace("%", "[%]").replace("_", "[_]")
    return query.replace("'", "''").replace('"', '""')


def safe_call(fn: Callable[..., Any]) -> Callable[..., Any]:
    """Convert ``OutlookError`` / unexpected exceptions to friendly messages.

    The wrapper preserves ``fn``'s signature via :func:`functools.wraps`
    so FastMCP introspects the real parameters (not ``*args, **kwargs``).
    Errors are re-raised so the MCP host marks the response
    ``isError: true`` — we don't return ``"Error: ..."`` strings, which
    would look like a normal text response.
    """

    @functools.wraps(fn)
    async def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return await fn(*args, **kwargs)
        except OutlookError as exc:
            raise RuntimeError(str(exc)) from exc
        except TimeoutError:
            raise
        except BaseException as exc:  # noqa: BLE001
            logger.exception("Unhandled error in %s", fn.__name__)
            raise RuntimeError(format_com_error(exc)) from exc

    return wrapper
