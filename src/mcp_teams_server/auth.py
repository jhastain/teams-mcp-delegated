# SPDX-FileCopyrightText: 2026 Jason Hastain
# SPDX-License-Identifier: Apache-2.0
"""Authentication helpers for mcp-teams-server.

Provides a credential factory that returns either a ClientSecretCredential
(app/service-principal mode) or a DeviceCodeCredential (delegated/user mode).

For delegated mode the module persists an AuthenticationRecord and uses
TokenCachePersistenceOptions so that the interactive device-code flow only
runs once per machine.
"""

import logging
import os
import platform
import sys
import webbrowser
from pathlib import Path

from azure.identity import (
    AuthenticationRecord,
    DeviceCodeCredential,
    TokenCachePersistenceOptions,
)
from azure.identity.aio import ClientSecretCredential

LOGGER = logging.getLogger(__name__)

DELEGATED_SCOPES = [
    "https://graph.microsoft.com/.default",
]

# Permissions needed per feature area (for reference and future tier system).
# When using .default scope, all permissions granted to the app registration
# are included in the token. These lists document what's actually required.
TRANSCRIPT_SCOPES = [
    "OnlineMeetingTranscript.Read.All",
]


def _get_auth_record_path(config) -> Path:
    """Return the path where the AuthenticationRecord JSON is persisted.

    On Windows: %LOCALAPPDATA%/.IdentityService/mcp-teams-server-auth-record.json
    Elsewhere:  ~/.IdentityService/mcp-teams-server-auth-record.json
    """
    if platform.system() == "Windows":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    else:
        base = Path.home()
    return base / ".IdentityService" / "mcp-teams-server-auth-record.json"


def _load_auth_record(path: Path) -> AuthenticationRecord | None:
    """Load a previously saved AuthenticationRecord, or return None."""
    if not path.exists():
        return None
    try:
        data = path.read_text(encoding="utf-8")
        return AuthenticationRecord.deserialize(data)
    except Exception:
        LOGGER.warning("Failed to load auth record from %s; re-authenticating", path)
        return None


def _save_auth_record(record: AuthenticationRecord, path: Path) -> None:
    """Persist an AuthenticationRecord to disk."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(record.serialize(), encoding="utf-8")
    LOGGER.info("Auth record saved to %s", path)


def _device_code_prompt(verification_uri: str, user_code: str, _expires_on) -> None:
    """Open the browser for device-code auth and print the code to stderr.

    stdout is reserved for the MCP JSON-RPC protocol, so all user-facing
    prompts must go to stderr. The browser is opened automatically so the
    user only needs to paste the code.
    """
    msg = (
        f"\n[Teams MCP] Authentication required.\n"
        f"A browser window has been opened. Enter code: {user_code}\n"
        f"If the browser didn't open, go to: {verification_uri}\n"
    )
    print(msg, file=sys.stderr, flush=True)

    # Show the code in a system notification / dialog so the user can see it
    # even when stderr is hidden (e.g. MCP server launched by Claude Code).
    try:
        if platform.system() == "Windows":
            import subprocess
            subprocess.Popen(
                [
                    "powershell.exe", "-NoProfile", "-Command",
                    (
                        "Add-Type -AssemblyName System.Windows.Forms; "
                        "[System.Windows.Forms.MessageBox]::Show("
                        f"'Enter this code in the browser:\\n\\n{user_code}\\n\\n"
                        f"URL: {verification_uri}', "
                        "'Teams MCP - Authentication Required', "
                        "'OK', 'Information') | Out-Null"
                    ),
                ],
                creationflags=0x00000008,  # DETACHED_PROCESS
            )
        else:
            # macOS/Linux: use notify-send or osascript
            if platform.system() == "Darwin":
                import subprocess
                subprocess.Popen([
                    "osascript", "-e",
                    f'display dialog "Teams MCP Auth Code: {user_code}" '
                    f'with title "Teams MCP" buttons {{"OK"}} default button "OK"',
                ])
            else:
                import subprocess
                subprocess.Popen([
                    "notify-send", "Teams MCP - Auth Required",
                    f"Enter code: {user_code}",
                ])
    except Exception:
        pass

    try:
        webbrowser.open(verification_uri)
    except Exception:
        pass


def create_credential(config):
    """Create the appropriate credential for the configured auth mode.

    Returns
    -------
    ClientSecretCredential (async, from azure.identity.aio)
        when ``config.is_delegated`` is False (app / service-principal mode).

    DeviceCodeCredential (sync, from azure.identity)
        when ``config.is_delegated`` is True (delegated / user mode).
        The GraphServiceClient accepts both sync and async credentials.
    """
    if not config.is_delegated:
        LOGGER.info("Using app (client-secret) authentication")
        return ClientSecretCredential(
            tenant_id=config.APP_TENANTID,
            client_id=config.APP_ID,
            client_secret=config.APP_PASSWORD,
        )

    LOGGER.info("Using delegated (device-code) authentication")

    cache_options = TokenCachePersistenceOptions(name=config.TOKEN_CACHE_NAME)
    auth_record_path = _get_auth_record_path(config)
    auth_record = _load_auth_record(auth_record_path)

    kwargs: dict = {
        "client_id": config.APP_ID,
        "tenant_id": config.APP_TENANTID,
        "cache_persistence_options": cache_options,
        "prompt_callback": _device_code_prompt,
    }

    if auth_record is not None:
        kwargs["authentication_record"] = auth_record
        LOGGER.info("Loaded existing auth record; silent auth should succeed")

    credential = DeviceCodeCredential(**kwargs)

    # If there is no cached record, perform an initial authenticate() so we
    # can persist the record for future runs.  enable_cae=True so that the
    # cached token satisfies the Graph SDK's CAE-enabled get_token() calls.
    if auth_record is None:
        LOGGER.info("No cached auth record found; initiating device-code flow")
        record = credential.authenticate(
            scopes=DELEGATED_SCOPES, enable_cae=True
        )
        _save_auth_record(record, auth_record_path)

    # Warm the token cache so the first Graph SDK call doesn't trigger
    # another interactive prompt. The SDK calls get_token() with .default
    # scope — this ensures a valid token is cached for that scope.
    try:
        credential.get_token(*DELEGATED_SCOPES, enable_cae=True)
        LOGGER.info("Token cache warmed successfully")
    except Exception as e:
        LOGGER.warning("Failed to warm token cache: %s", e)

    return credential
