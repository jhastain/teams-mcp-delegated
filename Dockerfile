FROM ghcr.io/astral-sh/uv:python3.10-alpine

# ENV TEAMS_APP_ID="" TEAMS_APP_PASSWORD="" TEAMS_APP_TYPE="" TEAMS_APP_TENANT_ID="" TEAM_ID="" TEAMS_CHANNEL_ID=""

LABEL \
  org.opencontainers.image.source="https://github.com/jhastain/teams-mcp-delegated" \
  org.opencontainers.image.title="Teams MCP (delegated)" \
  org.opencontainers.image.description="MCP server for Microsoft Teams with delegated per-user auth. Fork of InditexTech/mcp-teams-server." \
  org.opencontainers.image.licenses="Apache-2.0"

# Settings for faster container start
ENV UV_COMPILE_BYTECODE=0 UV_PYTHON_DOWNLOADS=0 UV_LINK_MODE=copy

COPY pyproject.toml LICENSE.txt *.md uv.lock src /app/
WORKDIR /app
RUN uv sync --frozen --no-dev

CMD ["uv", "run", "--frozen", "--no-dev", "mcp-teams-server"]