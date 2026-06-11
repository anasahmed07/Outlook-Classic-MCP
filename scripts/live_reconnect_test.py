"""Live verification: bridge survives Outlook being closed mid-session.

Run manually:  uv run python _live_reconnect_test.py
Launches real OUTLOOK.EXE, so not part of the pytest suite.
"""

import asyncio
import logging
import subprocess
import sys
import time

from outlook_mcp.bridge import OutlookBridge
from outlook_mcp.client.account import whoami

logging.basicConfig(level="INFO", stream=sys.stderr)


def close_outlook() -> None:
    # Same effect as the user closing the window (graceful WM_CLOSE),
    # with taskkill /f as fallback.
    subprocess.run(["taskkill", "/im", "OUTLOOK.EXE"], capture_output=True)
    time.sleep(3)
    subprocess.run(["taskkill", "/f", "/im", "OUTLOOK.EXE"], capture_output=True)
    time.sleep(2)


async def main() -> int:
    bridge = OutlookBridge()
    print("1) starting bridge (may cold-launch Outlook)...", flush=True)
    bridge.start()

    me = await bridge.call(whoami)
    print(f"2) whoami before close: {me['current_user']}", flush=True)

    print("3) closing Outlook out from under the bridge...", flush=True)
    close_outlook()

    print("4) calling whoami again — expecting auto-reconnect...", flush=True)
    me2 = await bridge.call(whoami)
    print(f"5) whoami after reconnect: {me2['current_user']}", flush=True)

    bridge.stop()
    print("PASS: bridge survived Outlook being closed", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
