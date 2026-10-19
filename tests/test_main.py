# SPDX-FileCopyrightText: 2025 INDUSTRIA DE DISEÑO TEXTIL, S.A. (INDITEX, S.A.)
# SPDX-License-Identifier: Apache-2.0
import os
import sys
from unittest.mock import patch

import pytest

import mcp_teams_server
from mcp_teams_server import main


def test_main_should_exit_error_on_missing_env_vars():
    # Unset environment
    for var in mcp_teams_server.REQUIRED_ENV_VARS:
        os.environ.pop(var, None)

    test_args = ["main"]
    with patch.object(sys, "argv", test_args):
        with pytest.raises(SystemExit) as exit_code:
            main()

    assert exit_code.type is SystemExit
    assert exit_code.value.code == 1


@pytest.mark.asyncio
async def test_list_tools():
    tools = await mcp_teams_server.mcp.list_tools()

    assert tools is not None
