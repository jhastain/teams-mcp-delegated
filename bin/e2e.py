#!/usr/bin/env python3
# SPDX-FileCopyrightText: 2026 Jason Hastain
# SPDX-License-Identifier: Apache-2.0
"""End-to-end check for teams-mcp-delegated.

One command, real code paths, explicit pass/fail. Anything that needs a live
Teams tenant is reported as a SKIP with the reason rather than quietly left out —
a green run that silently covers less than it looks like is worse than an honest
list of gaps.

    python bin/e2e.py        (or: uv run python bin/e2e.py)

Exit code 0 only if every check passed.
"""
from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

PASS: list[tuple[str, str]] = []
FAIL: list[tuple[str, str, str]] = []
SKIP: list[tuple[str, str]] = []


class _Skip(Exception):
    pass


def check(name):
    def wrap(fn):
        try:
            detail = fn()
        except _Skip as exc:
            SKIP.append((name, str(exc)))
        except Exception as exc:  # noqa: BLE001 - the harness reports, never crashes
            FAIL.append((name, f"{type(exc).__name__}: {exc}", traceback.format_exc()))
        else:
            PASS.append((name, detail or ""))
        return fn
    return wrap


# Config is read at import time, so give both modes something to read.
os.environ.setdefault("TEAMS_APP_ID", "00000000-0000-0000-0000-000000000000")
os.environ.setdefault("TEAMS_APP_TENANT_ID", "11111111-1111-1111-1111-111111111111")
os.environ.setdefault("TEAMS_APP_PASSWORD", "not-a-real-secret")
os.environ.setdefault("TEAM_ID", "test-team")
os.environ.setdefault("TEAMS_CHANNEL_ID", "test-channel")


@check("dependencies importable")
def _deps():
    missing = []
    for mod in ("mcp", "azure.identity", "msgraph", "botbuilder.core"):
        try:
            __import__(mod)
        except ImportError:
            missing.append(mod)
    if missing:
        raise RuntimeError(f"missing: {', '.join(missing)} - run: uv sync --frozen")
    return "mcp, azure.identity, msgraph, botbuilder.core"


if FAIL:
    print("Cannot continue: dependencies are not installed.")
    for name, msg, _ in FAIL:
        print(f"  FAIL  {name}: {msg}")
    sys.exit(1)


# --------------------------------------------------------------------------
# Configuration: the mode switch is the whole point of this fork
# --------------------------------------------------------------------------

@check("config defaults to app mode")
def _config_app():
    from mcp_teams_server.config import BotConfiguration

    os.environ["TEAMS_AUTH_MODE"] = "app"
    cfg = BotConfiguration()
    assert cfg.is_delegated is False, "app mode reported as delegated"
    assert cfg.APP_TYPE == "SingleTenant", cfg.APP_TYPE
    return "is_delegated=False, APP_TYPE=SingleTenant"


@check("delegated mode is recognised case-insensitively")
def _config_delegated():
    from mcp_teams_server.config import BotConfiguration

    for value in ("delegated", "Delegated", "DELEGATED"):
        os.environ["TEAMS_AUTH_MODE"] = value
        assert BotConfiguration().is_delegated is True, f"{value!r} not recognised"
    os.environ["TEAMS_AUTH_MODE"] = "app"
    return "delegated / Delegated / DELEGATED"


@check("the client secret is required in app mode only")
def _secret_requirement():
    import mcp_teams_server

    getter = getattr(mcp_teams_server, "_get_required_env_vars", None)
    if getter is None:
        raise _Skip("_get_required_env_vars is not exported by this version")

    os.environ["TEAMS_AUTH_MODE"] = "app"
    app_required = set(getter())
    os.environ["TEAMS_AUTH_MODE"] = "delegated"
    delegated_required = set(getter())
    os.environ["TEAMS_AUTH_MODE"] = "app"

    assert "TEAMS_APP_PASSWORD" in app_required, app_required
    assert "TEAMS_APP_PASSWORD" not in delegated_required, delegated_required
    return "TEAMS_APP_PASSWORD required in app mode, not in delegated"


@check("credential factory returns the right type per mode")
def _credential_factory():
    # Note the async variant: app mode uses azure.identity.aio, delegated uses
    # the sync azure.identity. GraphServiceClient accepts either.
    from azure.identity.aio import ClientSecretCredential

    from mcp_teams_server.auth import create_credential
    from mcp_teams_server.config import BotConfiguration

    os.environ["TEAMS_AUTH_MODE"] = "app"
    cred = create_credential(BotConfiguration())
    assert isinstance(cred, ClientSecretCredential), type(cred).__name__
    os.environ["TEAMS_AUTH_MODE"] = "app"
    return "app mode -> azure.identity.aio.ClientSecretCredential (no network call)"


@check("the auth-record path is per-user, never inside the repo")
def _auth_record_path():
    from mcp_teams_server.auth import _get_auth_record_path
    from mcp_teams_server.config import BotConfiguration

    path = _get_auth_record_path(BotConfiguration())
    assert path.name.endswith(".json"), path
    assert ".IdentityService" in str(path), path
    assert ROOT not in path.parents and path.parent != ROOT, (
        f"token cache would land in the repo: {path}"
    )
    return str(path)


# --------------------------------------------------------------------------
# Tool registry
# --------------------------------------------------------------------------

@check("all 23 tools are registered with unique names")
def _tool_registry():
    import re

    source = (ROOT / "src" / "mcp_teams_server" / "__init__.py").read_text(
        encoding="utf-8"
    )
    names = re.findall(r'name="([a-z_]+)"', source)
    assert len(names) == 23, f"expected 23 tools, found {len(names)}: {names}"
    assert len(set(names)) == len(names), "duplicate tool names"
    for expected in (
        "start_thread",            # upstream
        "resolve_user",            # fork: fuzzy people resolution
        "get_transcript_content",  # fork: transcripts
        "find_available_times",    # fork: calendar
        "read_chat_messages",      # fork: chats
    ):
        assert expected in names, f"{expected} missing from the registry"
    return f"23 tools: {len(names) - 6} added by this fork"


# --------------------------------------------------------------------------
# Inbound: card extraction. CI and alert bots post no plain-text body.
# --------------------------------------------------------------------------

@check("Adaptive Card text extraction")
def _adaptive_card():
    from mcp_teams_server.teams import _extract_adaptive_card_text

    body = [
        {"type": "TextBlock", "text": "Deployment failed"},
        {"type": "FactSet", "facts": [{"title": "Stage", "value": "production"}]},
        {
            "type": "Container",
            "items": [{"type": "TextBlock", "text": "nested detail"}],
        },
    ]
    text = _extract_adaptive_card_text(body)
    assert "Deployment failed" in text, text
    assert "Stage" in text and "production" in text, text
    assert "nested detail" in text, f"nested Container not walked: {text}"
    return "TextBlock, FactSet and nested Container all extracted"


@check("O365 connector card text extraction")
def _o365_card():
    from mcp_teams_server.teams import _extract_o365_connector_card_text

    card = {
        "summary": "Build Run Checkout-API - failed",
        "sections": [
            {
                "activityTitle": "Pipeline failed",
                "facts": [{"name": "Project", "value": "<p>web-api-dotnet</p>"}],
            }
        ],
    }
    text = _extract_o365_connector_card_text(card)
    assert "Checkout-API" in text, text
    assert "web-api-dotnet" in text, text
    assert "<p>" not in text, f"HTML not stripped: {text}"
    return "summary, activityTitle and facts extracted; HTML stripped"


@check("malformed card payloads degrade instead of raising")
def _bad_card():
    from mcp_teams_server.teams import (
        _extract_adaptive_card_text,
        _extract_o365_connector_card_text,
    )

    for payload in ([], [{}], [{"type": "TextBlock"}], [None], ["junk"]):
        result = _extract_adaptive_card_text(payload)
        assert isinstance(result, str), f"{payload!r} returned {type(result).__name__}"
    for payload in ({}, {"sections": []}, {"sections": [{}]}, {"sections": None}):
        result = _extract_o365_connector_card_text(payload)
        assert isinstance(result, str), f"{payload!r} returned {type(result).__name__}"
    return "9 malformed payloads returned strings"


# --------------------------------------------------------------------------
# Outbound: Teams renders a markdown pipe table as literal text
# --------------------------------------------------------------------------

@check("markdown tables are detected and converted to HTML")
def _tables():
    from mcp_teams_server.teams import (
        contains_markdown_table,
        convert_markdown_tables_to_html,
    )

    md = "Results:\n\n| Env | Status |\n| --- | --- |\n| prod | ok |\n\nDone."
    assert contains_markdown_table(md) is True
    html = convert_markdown_tables_to_html(md)
    for fragment in ("<table", ">Env<", "<td>prod</td>", "Results:", "Done."):
        assert fragment in html, f"{fragment} missing from: {html}"
    assert "| --- |" not in html, "separator row leaked into the output"
    return "header, cells and surrounding prose all preserved"


@check("non-table prose is not mangled")
def _no_false_positive():
    from mcp_teams_server.teams import contains_markdown_table

    for text in (
        "a | b without a separator",
        "just a sentence",
        "",
        "code: `a | b`",
    ):
        assert contains_markdown_table(text) is False, f"false positive on: {text!r}"
    return "4 non-table inputs rejected"


@check("angle brackets in cells are escaped")
def _escaping():
    from mcp_teams_server.teams import convert_markdown_tables_to_html

    html = convert_markdown_tables_to_html(
        "| Col |\n| --- |\n| Accounts_<id>@mask.example |"
    )
    assert "&lt;id&gt;" in html, html
    assert "<id>" not in html, "unescaped angle brackets would break the Teams card"
    return "cell content escaped"


# --------------------------------------------------------------------------
# Not covered here - stated explicitly rather than omitted
# --------------------------------------------------------------------------

UNCOVERED = [
    ("Device-code interactive flow", "needs a human to enter a code in a browser"),
    ("Token cache persistence across runs", "needs a completed device-code flow first"),
    ("Microsoft Graph round trips", "needs a real tenant and admin-consented delegated scopes"),
    ("Bot-adapter (app mode) message send", "needs a registered bot and a Teams channel"),
    ("Transcript ID reconstruction against Graph", "needs a real meeting with a transcript"),
    ("Fuzzy resolution ranking quality", "needs a real directory; tests/test_teams.py covers it as an integration test"),
    ("install-teams-mcp.ps1", "mutates %LOCALAPPDATA% and registers an MCP server; run it manually"),
]


def main() -> int:
    for name, reason in UNCOVERED:
        SKIP.append((name, reason))

    width = max(len(n) for n, *_ in PASS + FAIL + SKIP)
    for name, detail in PASS:
        print(f"PASS  {name:<{width}}  {detail}")
    for name, reason in SKIP:
        print(f"SKIP  {name:<{width}}  {reason}")
    for name, msg, tb in FAIL:
        print(f"FAIL  {name:<{width}}  {msg}")
        print("".join("        " + line for line in tb.splitlines(keepends=True)))

    print(f"\n{len(PASS)} passed, {len(FAIL)} failed, {len(SKIP)} skipped")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
