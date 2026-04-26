# Setting up the Outlook integration

Read this file when the user wants to use Outlook from Claude (or another MCP-aware agent) but the `outlook_*` tools aren't available in the session yet — meaning the underlying integration isn't installed.

This file walks an agent through helping the user install **`outlook-classic-mcp`**, the local MCP server that exposes classic Outlook desktop to MCP clients. After install, the user restarts their MCP client (Claude Desktop, Claude Code, Cursor, Cline, Continue, or Windsurf) and the `outlook_*` tools will appear.

## Prerequisites — verify these first

Ask the user to confirm each one. If any answer is no, stop and address that before proceeding; otherwise the install will succeed but the tools won't work.

1. **Windows 10 or 11.** macOS and Linux are not supported — there is no Outlook COM API there.
2. **Classic Outlook desktop is installed** (the `OUTLOOK.EXE` shipped with Microsoft 365 / Office). The "new" Outlook (`olk.exe`) does **not** expose COM and the integration cannot talk to it. Quick test: ask the user to open Outlook; if the title bar shows a "New Outlook" toggle and it's on, ask them to switch it off (toggle to "off" → Outlook restarts as classic).
3. **Python 3.10 or newer.** If they don't have it, the source-install path below uses `uv` to fetch Python 3.11 automatically — they don't need to install Python separately for that path.
4. **An Outlook account already signed in.** The integration piggybacks on whatever account Outlook is signed into. There is no separate auth flow, no Microsoft Graph, no Entra app, no OAuth tokens. If Outlook isn't signed in, sign in first.

The user does **not** need to have Outlook open before install; the integration auto-launches Outlook on its first call.

## Install path A — Claude Code plugin (simplest, recommended)

If the user is on Claude Code, the whole repo ships as a Claude Code **plugin** (single-plugin marketplace). Installing the plugin registers both the MCP server *and* this skill in one step. The MCP server itself is fetched and run on demand by `uvx` directly from PyPI ([outlook-classic-mcp](https://pypi.org/project/outlook-classic-mcp/)) — no `pip install`, no `git clone`, no `.venv` to maintain.

### Prereq: `uv` on PATH

The plugin's MCP server entry uses `uvx`, Astral's zero-install Python launcher. If the user doesn't have `uv` yet, install it once — it's a single command:

PowerShell (recommended on Windows):
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Then **open a fresh terminal** (PATH refresh) before continuing.

Confirm it's there: `uv --version` should print a version.

### Install the plugin

Two slash commands inside Claude Code, in this order:

```text
/plugin marketplace add anasahmed07/Outlook-Classic-MCP
/plugin install outlook@outlook-classic-mcp
```

`anasahmed07/Outlook-Classic-MCP` is the GitHub `owner/repo`. The first command adds the marketplace, the second picks the `outlook` plugin out of it.

Restart Claude Code (`/quit`, then reopen). On first call, `uvx` resolves and caches the package — there's a one-time 5–15 second pause while it installs. Subsequent calls are instant.

Confirm it works: `outlook_whoami` should return the bound mailbox.

### Updating later

```text
/plugin marketplace update outlook-classic-mcp
/plugin update outlook@outlook-classic-mcp
```

To pick up a new MCP server release on PyPI without changing the plugin, `uvx` re-resolves on its own cache TTL. To force a refresh: `uv cache clean outlook-classic-mcp`.

### Manual override: not using uv

If the user can't or won't install uv, edit the plugin's `mcpServers` entry (the file `/plugin list` points at) to use the system Python instead:

```json
"mcpServers": {
  "outlook": {
    "command": "python",
    "args": ["-m", "outlook_mcp"]
  }
}
```

…and have them `pip install outlook-classic-mcp` once. Functionally identical; loses the venv isolation that `uvx` gives you for free.

## Install path B — from PyPI (non-Claude-Code clients, or when plugin install isn't available)

For Cursor, Cline, Continue, Windsurf, Claude Desktop, or anywhere else, install the package directly from PyPI using `uv`:

```bat
uv pip install --system outlook-classic-mcp
python -m outlook_mcp.scripts.install_to_clients
```

The second command auto-detects which MCP clients are installed and registers the server with the ones the user picks (see "Registering with MCP clients" below). Re-running it is idempotent — it updates existing entries.

`--system` writes to the system Python install so `python -m outlook_mcp` resolves anywhere; drop the flag if the user is installing into an active venv. If they don't have `uv` yet, refer them to path A above for the one-line installer.

The package is published at https://pypi.org/project/outlook-classic-mcp/ — `uv pip install` should always succeed on a Windows machine with Python 3.10+.

## Install path C — from source

```bat
git clone https://github.com/anasahmed07/Outlook-Classic-MCP.git
cd Outlook-Classic-MCP
install.bat
```

`install.bat` does, in order:

1. Installs `uv` (Astral's Python installer) if it isn't already present.
2. Creates a `.venv` with Python 3.11.
3. Installs the package in editable mode (`pip install -e .`).
4. Pre-warms the pywin32 typelib cache.
5. Launches the smart client installer (next section).

The whole thing is unattended after they answer the client-selection menu.

## Registering with MCP clients

The smart client installer (`scripts/install_to_clients.py`, also reachable as `python -m outlook_mcp.scripts.install_to_clients`) auto-detects which MCP-aware clients are installed on the machine and shows a checkbox menu:

```
Select which clients to register outlook-mcp with:
  [ ] 1. Claude Desktop      C:\Users\you\AppData\Roaming\Claude\claude_desktop_config.json
  [ ] 2. Claude Code         (via `claude` CLI)
  [ ] 3. Cursor              C:\Users\you\.cursor\mcp.json
  ...

Type a number to toggle, 'a' to select all, 'n' for none,
'enter' to confirm, 'q' to quit without changes.
```

Supported clients: Claude Desktop, Claude Code, Cursor, Cline, Continue, Windsurf. The installer deep-merges the right entry into each config file (and runs `claude mcp add` for Claude Code), snapshotting any existing config to `<file>.bak` first. Re-running it is idempotent — it updates existing entries instead of duplicating.

## After install — making the tools appear

The MCP client only loads servers at startup, so the user must:

1. **Quit and reopen the MCP client** (e.g. Claude Desktop). For VS Code-based clients, reload the window or restart the editor.
2. Start a new conversation (the tool list is per-session in some clients).
3. The `outlook_*` tools will be available.

You can confirm it worked by calling `outlook_whoami` — it should return the bound mailbox.

## First-call behavior

The very first `outlook_*` call after a fresh install (or after a machine reboot) takes 5–10 seconds while Outlook's COM surface boots. Subsequent calls are fast. **Don't retry** — wait. If the user complains it's slow, this is normal once.

If Outlook isn't running, the integration auto-launches it. On tightly locked-down machines (UAC + group policy blocking COM activation), this may fail; in that case ask the user to open Outlook manually before retrying.

## Optional environment variables

Set these in the MCP client's config (alongside the `command` and `args` for the server) when needed:

- `OUTLOOK_MCP_ALLOW_ANY_PATH=1` — disables the user-profile sandbox on `attachments=` and `output_dir=` parameters. Use only when the user genuinely needs to attach files from / save files to locations outside `C:\Users\<them>\...` (e.g. corporate file shares).

The user must restart the MCP client after changing env vars.

## Troubleshooting tree

Walk the user through these in order when something doesn't work:

### "I installed it but `outlook_*` tools still don't show up."

1. Did they fully quit and reopen the MCP client (not just minimize / new chat)?
2. Did the install register with the client they're actually using? Re-run `python -m outlook_mcp.scripts.install_to_clients` and confirm the right client is checked.
3. For Claude Desktop / Cursor / Cline / Continue / Windsurf, ask them to open the MCP config file and confirm an entry like:
   ```json
   "outlook": {
     "command": ".venv/Scripts/python.exe",
     "args": ["-m", "outlook_mcp"]
   }
   ```
   exists under `mcpServers`. If it doesn't, the registration failed silently — re-run the installer.

### "Outlook COM thread did not become ready."

Outlook didn't auto-launch. Ask the user to:

1. Open Outlook manually and let it finish loading (sign in if prompted).
2. Restart the MCP client so it re-spawns the integration.
3. Try a tool call again.

### Send / reply / forward "succeeds" but nothing actually goes out

Outlook's Programmatic Access security is blocking the call. See `gotchas.md` → "*You said you sent it but I don't see it in Sent Items.*" The fix is in Outlook → File → Options → Trust Center → Programmatic Access. On corporate machines IT may need to whitelist the Python interpreter as a trusted publisher.

### "Inspector / debug tool shows ENOENT for the python path."

Known Windows quirk in some MCP debug tooling — the tool mangles backslashes. Tell the user to use forward slashes in the python path (`C:/Users/.../python.exe`) instead of backslashes. Doesn't affect normal usage; only relevant if they're debugging with `npx @modelcontextprotocol/inspector`.

### "I'm on the new Outlook and it doesn't work."

The integration only supports classic Outlook (`OUTLOOK.EXE`). The new Outlook (`olk.exe`) has no COM surface. Tell them:

> Open Outlook → top-right corner → toggle "New Outlook" off. Outlook will restart as classic. Then try again.

### "It worked, but every send pops up a security dialog."

That's Outlook's Programmatic Access guard asking to confirm. Permanently silencing it: Outlook → File → Options → Trust Center → Programmatic Access → "Never warn me about suspicious activity (not recommended)". On a corporate-managed machine, IT policy may prevent this — in which case the user has to click through the dialog each time, or get IT to add the Python interpreter as a trusted publisher.

## What you do not need to do

- **Don't** ask the user for OAuth tokens, Microsoft Graph credentials, or Entra app registrations. None of those are involved.
- **Don't** ask for SMTP passwords. The integration uses Outlook's own session.
- **Don't** assume Linux/macOS instructions translate. They don't — this is Windows-only.

## After install completes

Once the tools are available, you don't need this file anymore. Switch back to `tools.md` and `recipes.md` for actual operation, and `gotchas.md` for failure modes.
