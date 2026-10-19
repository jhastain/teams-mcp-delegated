# Contributing

This is a personal fork of
[InditexTech/mcp-teams-server](https://github.com/InditexTech/mcp-teams-server), kept as a portfolio
project. Issues and pull requests are welcome, with two caveats:

- **If the change belongs upstream, send it upstream.** Anything in code inherited unchanged —
  the bot-mode adapter path, the six original tools, the EMEA `service_url` — is better fixed in
  `InditexTech/mcp-teams-server` so everyone gets it.
- **Changes here should not break app mode.** Delegated mode was added alongside the bot mode, not
  on top of it. Both paths must keep working.

## Before opening a PR

```bash
uv sync --frozen --all-extras --dev
uv run ruff check .
uv run pyright
uv run pytest            # offline tests only; integration is excluded by default
python bin/e2e.py        # the same checks plus an explicit list of what is not covered
```

`.github/workflows/PR-verify.yml` runs all of the above, plus `repolinter` and a REUSE compliance
check. New files need an SPDX header:

```python
# SPDX-FileCopyrightText: 2026 <you>
# SPDX-License-Identifier: Apache-2.0
```

A file you substantially modify keeps upstream's copyright line and gains yours, marked
`(modifications)`. Add new paths to the right block in `REUSE.toml`.

## Integration tests

`tests/test_teams.py` needs a real tenant and is excluded by default. To run it, fill in `.env`
(see `sample.env`, including `TEST_THREAD_ID` / `TEST_MESSAGE_ID` / `TEST_USER_NAME`) and run
`uv run pytest -m integration`. It posts real messages — use a scratch channel.

## Code of conduct

[`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md), inherited from upstream.
