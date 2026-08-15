# Migration Guide

This guide covers breaking changes between versions and how to migrate.

## v0.2.0 �?v0.3.0

### MCP package relocated: `mcp/` �?`src/mcp/`

The local MCP server previously lived in a top-level `mcp/` package. Because
the official `mcp` SDK (added as a dependency) uses the same top-level module
name, the local package moved to `src/mcp/` and is now imported as `src.mcp.*`.

**Actions required:**

| Where | Before | After |
| ----- | ------ | ----- |
| CLI command | `python -m mcp.server` | `python -m src.mcp.server` |
| Claude Code / cursor MCP config `args` | `["-m", "mcp.server"]` | `["-m", "src.mcp.server"]` |
| Docker entrypoint | `python -m mcp.server --transport sse ...` | `python -m src.mcp.server --transport sse ...` |
| Custom imports | `from mcp.errors import ...` | `from src.mcp.errors import ...` |
| CLI alias | �?| `scribe-mcp` (equivalent to `python -m src.mcp.server`) |

**Test scripts** moved to `tests/integration/mcp/`:

| Before | After |
| ------ | ----- |
| `python mcp/test_server.py` | `python tests/integration/mcp/mcp_stdio_smoke.py` |
| `python mcp/test_sse.py` | `python tests/integration/mcp/test_sse.py` (pytest) |
| `python mcp/test_agent.py` | `python tests/integration/mcp/mcp_agent_sim.py` |

### Stricter tooling

- `pytest` now treats warnings as errors (`filterwarnings = ["error"]`).
- `ruff` enforces an expanded rule set; `ruff format` is required.
- `mypy` runs in a stricter configuration; new modules are fully strict.

Run before submitting changes:

```bash
ruff check .
ruff format --check .
mypy src/tianshang_scribe/
pytest -q
```

## v0.4.0 → v0.5.0

### Import package renamed: `src` → `tianshang_scribe`

The top-level import package was renamed to align with the distribution name
and TianshangCAD's layout. `src/` is now a pure build-isolation directory and
the importable package is `tianshang_scribe/` inside it.

**Actions required:**

| Where | Before | After |
| ----- | ------ | ----- |
| CLI command | `python -m src.cli.main` | `python -m tianshang_scribe.cli.main` |
| MCP server | `python -m src.mcp.server` | `python -m tianshang_scribe.mcp.server` |
| Claude Code / cursor MCP config `args` | `["-m", "src.mcp.server"]` | `["-m", "tianshang_scribe.mcp.server"]` |
| Docker entrypoint | `python -m src.mcp.server --host 0.0.0.0` | `python -m tianshang_scribe.mcp.server --host 0.0.0.0` |
| Custom imports | `from src.cli.main import ...` | `from tianshang_scribe.cli.main import ...` |
| PyInstaller entry | `src/cli/main.py` | `src/tianshang_scribe/cli/main.py` |
| Coverage source | `--cov=src` | `--cov=tianshang_scribe` |

The `tianshang-scribe` and `scribe-mcp` console scripts are unchanged; they
resolve to the new module internally.

## v0.5.0 → v0.6.0

### Environment variable prefix renamed: `SCRIBE_*` → `TIANSHANG_SCRIBE_*`

All server configuration environment variables now use the
`TIANSHANG_SCRIBE_` prefix, matching the distribution/import naming scheme.

**Actions required:**

| Where | Before | After |
| ----- | ------ | ----- |
| Auth token | `SCRIBE_AUTH_TOKEN` | `TIANSHANG_SCRIBE_AUTH_TOKEN` |
| Additional tokens | `SCRIBE_API_KEYS` | `TIANSHANG_SCRIBE_API_KEYS` |
| CORS origins | `SCRIBE_CORS_ORIGINS` | `TIANSHANG_SCRIBE_CORS_ORIGINS` |
| Transport / host / port | `SCRIBE_TRANSPORT` / `SCRIBE_HOST` / `SCRIBE_PORT` | `TIANSHANG_SCRIBE_TRANSPORT` / `TIANSHANG_SCRIBE_HOST` / `TIANSHANG_SCRIBE_PORT` |
| Rate limits | `SCRIBE_RATE_LIMIT_MAX` / `SCRIBE_RATE_LIMIT_WINDOW` | `TIANSHANG_SCRIBE_RATE_LIMIT_MAX` / `TIANSHANG_SCRIBE_RATE_LIMIT_WINDOW` |
| Logging | `SCRIBE_LOG_LEVEL` / `SCRIBE_LOG_JSON` | `TIANSHANG_SCRIBE_LOG_LEVEL` / `TIANSHANG_SCRIBE_LOG_JSON` |

Update shell exports, `docker compose` environment blocks, CI pipelines, and
any `.env` file accordingly.

### MCP server console script renamed: `scribe-mcp` → `tianshang-scribe-server`

The MCP server console script now mirrors TianshangCAD's `tianshangcad-server`.

| Where | Before | After |
| ----- | ------ | ----- |
| Console script | `scribe-mcp` | `tianshang-scribe-server` |
| Claude Code / cursor MCP config `args` | `["scribe-mcp"]` | `["tianshang-scribe-server"]` |
| Docker entrypoint | unchanged | `python -m tianshang_scribe.mcp.server ...` (still valid) |

The `python -m tianshang_scribe.mcp.server` module entry point is unchanged.

### Prometheus metric names renamed: `scribe_*` → `tianshang_scribe_*`

| Where | Before | After |
| ----- | ------ | ----- |
| Metric | `scribe_operation_duration_seconds` | `tianshang_scribe_operation_duration_seconds` |
| Metric | `scribe_operations_total` | `tianshang_scribe_operations_total` |

Update Prometheus scrape targets, dashboards and alert rules that reference
the old metric names.

### Docker named volume renamed: `scribe_output` → `tianshang_scribe_output`

`docker compose up` will create a fresh `tianshang_scribe_output` volume.
To preserve data from an existing volume, migrate it manually:

```bash
docker volume create tianshang_scribe_output
docker run --rm -v scribe_output:/from -v tianshang_scribe_output:/to alpine cp -a /from/. /to/
```

## Not yet released

Future releases will introduce MCP SDK 2.x Streamable HTTP transport,
document versioning, batch scheduling, and the script sandbox. Watch
`CHANGELOG.md` for the exact migration notes.
