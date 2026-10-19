# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

This is a fork of [InditexTech/mcp-teams-server](https://github.com/InditexTech/mcp-teams-server).
Entries below `[1.1.0]` are upstream's. The fork is additive: no upstream behaviour was removed.

## [1.1.0] - 2026-10-19

### Added
- **Delegated (per-user) auth mode**, selected by `TEAMS_AUTH_MODE=delegated`. New
  `src/mcp_teams_server/auth.py`: credential factory, `DeviceCodeCredential` with a persisted
  `AuthenticationRecord` so the device-code flow runs once per machine, explicit CAE cache warm-up,
  and a device-code prompt that writes to stderr (stdout is the JSON-RPC channel) while also opening
  the browser and raising a native dialog.
- **17 new tools**, taking the registry from 6 to 23: `list_teams`, `list_channels`,
  `resolve_teams_link`, `create_channel`, `add_channel_members`, `find_available_times`,
  `create_meeting`, `cancel_meeting`, `list_events`, `update_event`, `resolve_user`,
  `send_direct_message`, `list_meeting_transcripts`, `get_meeting_transcript`,
  `get_transcript_content`, `list_chats`, `read_chat_messages`.
- **Meeting-transcript access for any attendee**, by reconstructing the Graph online-meeting ID
  locally from the join URL instead of requiring `OnlineMeetings.Read`.
- **Local fuzzy people-resolution** over the caller's own Teams memberships, so partial names,
  misspellings and email prefixes all resolve.
- **Markdown-table to HTML conversion** for outbound messages, because Teams renders a pipe table as
  literal text.
- **Attribution footer** on every outbound message: "Posted via Teams MCP on behalf of \<user\>".
- **Adaptive Card and O365 Connector Card text extraction**, so bot-posted CI and alert messages are
  readable to the model instead of arriving empty.
- **`install-teams-mcp.ps1`** — Windows one-command installer, including a `LongPathsEnabled` check
  and a `robocopy /purge` workaround for deleting a venv whose paths exceed `MAX_PATH`.
- **`bin/e2e.py`** — single-command check of everything that works without a Teams tenant, with
  explicit SKIP lines for what does not.
- Offline unit tests: `tests/test_extraction.py`, `tests/test_table_conversion.py`.

### Changed
- `config.py` rewritten from a `dict` subclass to a dataclass with attribute access; adds
  `APP_TYPE`, `AUTH_MODE`, `TOKEN_CACHE_NAME` and an `is_delegated` property.
- `TEAMS_APP_PASSWORD` is required only in app mode.
- Project renamed to `teams-mcp-delegated` and retargeted away from upstream's CI, container
  registry, maintainers and vulnerability-disclosure programme. Apache-2.0 attribution
  (`LICENSE.txt`, `LICENSES/`, `NOTICE`, per-file SPDX headers) is retained.

### Removed
- Upstream CI workflows that were gated on upstream's organisation and could never run here: the
  SonarCloud analysis, the ghcr release job, and the OSSF Scorecard scan on `master`.

## [1.0.0] - 2025-04-01

### Added
- First release of MCP Teams Server
- Basic Teams integration through MCP tools
- Documentation for setup and usage
- Security guidelines and policies
- Basic workflows for PR verify and release docker image
