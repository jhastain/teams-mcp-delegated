# SPDX-FileCopyrightText: 2025 INDUSTRIA DE DISEÑO TEXTIL, S.A. (INDITEX, S.A.)
# SPDX-FileCopyrightText: 2026 Jason Hastain (modifications)
# SPDX-License-Identifier: Apache-2.0
import os
from dataclasses import dataclass


@dataclass
class BotConfiguration:
    def __init__(self):
        self.APP_ID = os.environ.get("TEAMS_APP_ID", "")
        self.APP_PASSWORD = os.environ.get("TEAMS_APP_PASSWORD", "")
        self.APP_TYPE = os.environ.get("TEAMS_APP_TYPE", "SingleTenant")
        self.APP_TENANTID = os.environ.get("TEAMS_APP_TENANT_ID", "")
        self.TEAM_ID = os.environ.get("TEAM_ID", "")
        self.TEAMS_CHANNEL_ID = os.environ.get("TEAMS_CHANNEL_ID", "")
        self.AUTH_MODE = os.environ.get("TEAMS_AUTH_MODE", "app")
        self.TOKEN_CACHE_NAME = os.environ.get(
            "TEAMS_TOKEN_CACHE_NAME", "mcp-teams-server.cache"
        )

    @property
    def is_delegated(self) -> bool:
        return self.AUTH_MODE.lower() == "delegated"
