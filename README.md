# teams-mcp-delegated

A [Model Context Protocol](https://modelcontextprotocol.io/introduction) server for **Microsoft
Teams** that can act **as the signed-in user** — reading their chats and calendar, scheduling
meetings, pulling meeting transcripts, and resolving colleagues by partial or misspelled name.

**This is a fork of [InditexTech/mcp-teams-server](https://github.com/InditexTech/mcp-teams-server)**
(Apache-2.0, © 2025 Industria de Diseño Textil S.A.). Upstream is a Teams *bot* that posts as itself.
This fork keeps that mode byte-for-byte and adds a second one: delegated per-user auth over Microsoft
Graph, plus 17 new tools and a Windows installer. Nothing upstream was removed.

> Portfolio copy. The Entra app ID and tenant ID this was originally installed against have been
> removed — the installer now takes them as parameters. Example names in tests and docstrings are
> fictional.

## What this fork adds

### 1. Delegated (per-user) authentication — the architectural change

Upstream only supports **app / service-principal** auth: a registered bot posting under its own
identity via `ClientSecretCredential` + `CloudAdapter`. That identity cannot read your chats, your
calendar, or a meeting transcript.

`src/mcp_teams_server/auth.py` (new, 186 lines) adds a credential factory selected by
`TEAMS_AUTH_MODE`:

- **`app`** (default) — `ClientSecretCredential`, exactly as upstream. Untouched.
- **`delegated`** — `DeviceCodeCredential` with `TokenCachePersistenceOptions`, persisting an
  `AuthenticationRecord` to `%LOCALAPPDATA%\.IdentityService\mcp-teams-server-auth-record.json`.
  The interactive device-code flow therefore runs **once per machine**, not once per invocation.

In delegated mode `app_lifespan` builds **no bot adapter at all** — it is pure Microsoft Graph with
the user's token, and resolves `/me` for message attribution.

Two details that were not obvious going in:

- **The device-code prompt cannot go to stdout.** stdout is the JSON-RPC channel; writing a prompt
  there corrupts the protocol. It goes to stderr — and because Claude Code hides stderr, the prompt
  additionally auto-opens the browser and raises a native dialog with the code (a detached
  `MessageBox` on Windows, `osascript` on macOS, `notify-send` on Linux).
- **Explicit cache warm-up.** `authenticate()` followed by `get_token(..., enable_cae=True)` runs at
  startup, so the Graph SDK's first CAE-enabled call does not re-prompt the user.

### 2. Tools: 6 → 23

Upstream's six (`start_thread`, `update_thread`, `read_thread`, `list_threads`,
`get_member_by_name`, `list_members`) plus seventeen:

| Area | Tools |
| --- | --- |
| Discovery | `list_teams`, `list_channels`, `resolve_teams_link`, `create_channel`, `add_channel_members` |
| Calendar | `find_available_times`, `create_meeting`, `cancel_meeting`, `list_events`, `update_event` |
| People | `resolve_user`, `send_direct_message` |
| Transcripts | `list_meeting_transcripts`, `get_meeting_transcript`, `get_transcript_content` |
| Chats | `list_chats`, `read_chat_messages` |

Each has dual internals (`_start_thread_graph` vs the bot path, `_send_dm_delegated` vs
`_send_dm_app`, …) so every tool works in both auth modes.

### 3. Three pieces of engineering worth the read

**Transcript access without `OnlineMeetings.Read`.** Fetching a transcript normally requires a
permission the app registration does not hold, and normally only the organizer can do it.
`_build_meeting_id_from_join_url()` reconstructs the Graph online-meeting ID locally — base64 of the
thread ID plus the organizer OID parsed out of the join URL — and `_resolve_meeting_id()` tries that
first, falling back to a Graph filter query. The result: **any attendee** can pull the transcript.
The content fetch sets an explicit `Accept: text/vtt` header, because the SDK's default is wrong for
this endpoint.

**Local fuzzy people-resolution.** `_build_directory_cache()` builds a directory from the caller's
own Teams memberships; `_fuzzy_score()` ranks candidates locally. So `"Alex"`, `"Alex Rivera"`,
`"alx rivra"` and `"alex.rivera"` all resolve to the same person, ranked by relevance to the
signed-in user — no directory-wide search permission needed.

**Outbound rendering and attribution.** Teams renders a markdown pipe table as literal `| --- |`
text, so `convert_markdown_tables_to_html()` turns it into real `<table>` markup and a text message
containing a table is auto-upgraded to HTML. Separately, `_attribution_footer()` appends
*"Posted via Teams MCP on behalf of &lt;user&gt;"* to every outbound message, in both the HTML and
plaintext bodies. An agent posting as a human should say so.

**Inbound card extraction.** CI and alerting bots post Adaptive Cards and legacy O365 Connector
Cards, which carry no plain-text body — so the model sees an empty message. The extraction layer
pulls readable text out of both.

### 4. Windows installer

`install-teams-mcp.ps1` (397 lines) is one command: detect or winget-install Python 3.10-3.12, create
an isolated venv under `%LOCALAPPDATA%`, find or build the wheel, install it, verify the import,
register the server with `claude mcp add --scope user`, and run the one-time device-code auth.

Two real Windows problems it solves:

- A `LongPathsEnabled` registry check with self-elevation guidance, because the msgraph SDK's paths
  exceed `MAX_PATH`.
- `Remove-LongPathDir`, which purges a >260-character venv via `robocopy /purge` against an empty
  directory — `Remove-Item -Recurse` simply cannot delete it.

## Install

```powershell
.\install-teams-mcp.ps1 -TeamsAppId <app-guid> -TeamsTenantId <tenant-guid>
```

Or set `TEAMS_APP_ID` / `TEAMS_APP_TENANT_ID` in the environment and run it with no arguments. You
need an Entra app registration with the delegated Graph scopes — see
[`doc/MS-Teams-setup.md`](doc/MS-Teams-setup.md) for the app registration, and
[`doc/delegated-auth.md`](doc/delegated-auth.md) for what delegated mode needs specifically.

Manual install:

```bash
uv sync --frozen
cp sample.env .env      # fill in your own IDs
uv run mcp-teams-server
```

## Configuration

| Variable | Mode | Meaning |
| --- | --- | --- |
| `TEAMS_AUTH_MODE` | both | `app` (default) or `delegated` |
| `TEAMS_APP_ID` | both | Entra app registration (client) ID |
| `TEAMS_APP_TENANT_ID` | both | Entra tenant ID |
| `TEAMS_APP_TYPE` | both | `SingleTenant` or `MultiTenant` |
| `TEAMS_APP_PASSWORD` | **app only** | Client secret. Not required in delegated mode. |
| `TEAM_ID` | both | Default Teams group ID |
| `TEAMS_CHANNEL_ID` | both | Default channel ID |
| `TEAMS_TOKEN_CACHE_NAME` | delegated | Token cache filename. Default `mcp-teams-server.cache`. |

## Verifying

```bash
python bin/e2e.py
```

Runs every check that works without a Teams tenant — imports, the 23-tool registry, both credential
factory branches, card extraction against captured payloads, markdown-table conversion — and prints
an explicit `SKIP` for each path that needs live Graph access, rather than omitting it. Exit code is
non-zero on any failure.

`uv run pytest` runs the same offline unit tests; the integration suite is excluded by default
(`-m "not integration"`) because it needs a real tenant.

## Known limitations

Stated rather than hidden:

- `teams.py` hardcodes `service_url = "https://smba.trafficmanager.net/emea/"`, inherited unchanged
  from upstream. It affects **app (bot) mode only**; a non-EMEA tenant needs this parameterised.
- Default timezone for calendar operations is `Central Standard Time`, not the caller's. Worth
  deriving from `/me/mailboxSettings`.
- Upstream's `doc/MS-Teams-setup.md` still describes the bot-mode setup. That is still accurate for
  app mode; delegated mode is documented in `doc/delegated-auth.md`.
- The CI workflows inherited from upstream that depended on upstream's org (the SonarCloud gate, the
  ghcr release job, OSSF Scorecard on `master`) are removed — they could never run here.
  `.github/workflows/PR-verify.yml` (ruff, pyright, pytest, repolinter, REUSE) is kept and does run.

## License and attribution

Apache-2.0, inherited from upstream. `LICENSE.txt`, `LICENSES/`, `NOTICE` and the per-file SPDX
headers are retained as Apache-2.0 §4 requires. Files substantially modified here carry an additional
`SPDX-FileCopyrightText: 2026 Jason Hastain (modifications)` line; files written here carry only that.

© 2025 INDUSTRIA DE DISEÑO TEXTIL S.A. (INDITEX S.A.) — original work
© 2026 Jason Hastain — modifications
