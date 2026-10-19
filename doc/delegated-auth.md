# Delegated auth mode

`doc/MS-Teams-setup.md` covers the **app (bot)** mode inherited from upstream. This covers the
delegated mode added by this fork: what it needs, how the token lifecycle works, and the two
non-obvious problems it had to solve.

## Why it exists

App mode authenticates as a registered bot. That identity is a separate principal from you, so it
cannot see your 1:1 chats, your calendar, or a meeting transcript — no amount of application
permission grants it *your* mailbox. Delegated mode authenticates **as you**, so everything you can
see in the Teams client is reachable.

The trade is symmetric and worth stating plainly: in delegated mode the server reads your chats,
calendar and transcripts, and posts messages under your name. `_attribution_footer()` appends
*"Posted via Teams MCP on behalf of \<you\>"* to every outbound message for exactly that reason.

## App registration

A single Entra app registration serves both modes. For delegated mode it needs:

- **Platform:** add a *Mobile and desktop applications* platform, and enable
  **Allow public client flows** — the device-code grant is a public-client flow and fails without it.
- **Delegated** Microsoft Graph permissions (not Application permissions), admin-consented. The
  credential requests `https://graph.microsoft.com/.default`, which means **the token carries
  whatever delegated permissions the app registration has been granted** — the server does not
  request scopes per feature. Grant the ones the tools you intend to use need: reading and sending
  channel messages, reading teams and channels, read/write chats, read/write calendars, basic user
  read, and `OnlineMeetingTranscript.Read.All` for the transcript tools (`TRANSCRIPT_SCOPES` in
  `auth.py` records that last one explicitly).
- **No client secret is required.** `TEAMS_APP_PASSWORD` is only read in app mode.

Then:

```
TEAMS_AUTH_MODE=delegated
TEAMS_APP_ID=<app registration client id>
TEAMS_APP_TENANT_ID=<tenant id>
TEAMS_APP_TYPE=SingleTenant
```

## Token lifecycle

```
first run ──▶ DeviceCodeCredential
              ├─ prompt written to stderr  (stdout is the JSON-RPC channel)
              ├─ browser opened automatically
              └─ native dialog raised with the code
                      │
                      ▼
              authenticate()  ──▶ AuthenticationRecord persisted to
                                  %LOCALAPPDATA%\.IdentityService\
                                      mcp-teams-server-auth-record.json
                      │
                      ▼
              get_token(.default, enable_cae=True)  ← explicit cache warm-up
                      │
later runs ──▶ record + TokenCachePersistenceOptions ──▶ silent refresh
```

Two details that are easy to get wrong:

**The prompt cannot go to stdout.** An MCP server over stdio uses stdout as the JSON-RPC channel.
Printing a device-code prompt there corrupts the protocol and the client disconnects. It goes to
stderr — and because Claude Code does not surface stderr to the user, `_device_code_prompt` also
opens the verification URL in the browser and raises a native dialog carrying the code: a detached
`System.Windows.Forms.MessageBox` on Windows, `osascript` on macOS, `notify-send` on Linux. Without
that, the first run looks like a hang.

**The cache needs warming explicitly.** Calling `authenticate()` alone leaves the CAE-enabled token
unfetched, so the Graph SDK's first continuous-access-evaluation call re-prompts. `create_credential`
therefore follows `authenticate()` with a `get_token(..., enable_cae=True)` at startup.

## Resetting

Delete the auth record and the server will re-prompt on next start:

```powershell
Remove-Item "$env:LOCALAPPDATA\.IdentityService\mcp-teams-server-auth-record.json"
```

## Transcript access

Graph's transcript endpoints normally need `OnlineMeetings.Read` and normally only answer for the
meeting organizer. This fork avoids both constraints:

1. `_build_meeting_id_from_join_url()` reconstructs the Graph online-meeting ID **locally** —
   base64 of the thread ID plus the organizer's object ID, both parseable out of the join URL.
2. `_resolve_meeting_id()` tries that first, and falls back to a Graph filter query if it fails.
3. The content fetch sets an explicit `Accept: text/vtt` header, because the SDK's default content
   type is wrong for this endpoint and returns an unhelpful error.

The result is that **any attendee** can pull a transcript, using only permissions they already have
as a participant.

## Both modes must keep working

Every tool has dual internals — `_start_thread_graph` versus the bot adapter path,
`_send_dm_delegated` versus `_send_dm_app`, and so on. `app_lifespan` branches once on
`config.is_delegated` and builds either a Graph client or a `CloudAdapter`, never both. App mode's
code path is byte-identical to upstream; if a change here would alter it, that is a bug.
