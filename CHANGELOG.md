# Changelog

All notable changes to this project are documented in this file.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [0.7.1] - 2026-08-19

### Fixed
- `--comment` on PowerPoint: the slide index was passed as a string, causing a
  `TypeError` when adding speaker notes. It is now coerced to an integer
  (`src/tianshang_scribe/cli/main.py`). Non-numeric first tokens fall back to
  appending the whole string as text on slide 0.

### Changed
- Documentation reconciled with the implementation: `--split` documents that
  only Excel `by-sheet` is supported (Word by-page / PPT by-slide pending);
  `--merge` documents comma-separated inputs without glob/wildcard support;
  `--comment` documents that on PPT it appends to the notes area (overlapping
  with `--notes`); `docs/mcp/ROADMAP.md` `tools_available` corrected from 5 to 7.
- Bumped `SERVER_VERSION` to `0.7.1` so the MCP health endpoint reports the
  current release.

## [0.7.0] - 2026-08-19

### Added
- `--math-mtef`: embed formulas as real MathType OLE objects (MTEF binary,
  written to `word/embeddings/oleObject*.bin`) so they remain editable by legacy
  MathType (6.x and earlier); the default remains native Word OMML. Reads back
  through the same format `--extract math` understands
  (`src/tianshang_scribe/rendering/mtef/`).
- `--math-font "Times New Roman"`: configurable OMML math rendering font
  (default Cambria Math), switching to a MathType-style serif font via
  `<m:mathPr><m:mathFont>` (`src/tianshang_scribe/rendering/math_omml.py`).

### Changed
- The math converter is rewritten as a recursive-descent parser
  (expression → term → factor → atom) over an immutable nested token tree
  (fraction, root, N-ary, sub/sup, accent, styled, delimiter tokens), driven by
  an O(1) command dispatch table with precompiled regexes and zero-copy argument
  slicing; output is byte-for-byte stable across releases, guarded by a
  golden-snapshot regression suite.

## [0.6.0] - 2026-08-14

### Changed
- **Breaking: environment variable prefix `SCRIBE_*` → `TIANSHANG_SCRIBE_*`.**
  `src/tianshang_scribe/utils/config.py` sets `env_prefix='TIANSHANG_SCRIBE_'`,
  so `SCRIBE_AUTH_TOKEN` becomes `TIANSHANG_SCRIBE_AUTH_TOKEN`,
  `SCRIBE_API_KEYS` → `TIANSHANG_SCRIBE_API_KEYS`,
  `SCRIBE_CORS_ORIGINS` → `TIANSHANG_SCRIBE_CORS_ORIGINS`,
  `SCRIBE_HOST/PORT/TRANSPORT` → `TIANSHANG_SCRIBE_HOST/PORT/TRANSPORT`,
  `SCRIBE_LOG_LEVEL/LOG_JSON` → `TIANSHANG_SCRIBE_LOG_LEVEL/LOG_JSON`,
  `SCRIBE_RATE_LIMIT_MAX/WINDOW` → `TIANSHANG_SCRIBE_RATE_LIMIT_MAX/WINDOW`.
  Dockerfile and docker-compose environment variables updated in lockstep.
- **Breaking: MCP server console script renamed `scribe-mcp` →
  `tianshang-scribe-server`** (mirrors TianshangCAD's `tianshangcad-server`).
  `python -m tianshang_scribe.mcp.server` and the Docker `python -m` entry
  point are unchanged. Migration notes in `MIGRATION.md`.
- **Breaking: Prometheus metric names now use the `tianshang_scribe_` prefix**:
  `scribe_operation_duration_seconds` →
  `tianshang_scribe_operation_duration_seconds`, `scribe_operations_total` →
  `tianshang_scribe_operations_total` (`src/tianshang_scribe/mcp/metrics.py`).
- **Docker named volume renamed `scribe_output` → `tianshang_scribe_output`**
  (`docker-compose.yml`); existing volumes must be migrated manually.

## [0.5.0] - 2026-08-14

### Changed
- **Breaking: import package renamed `src` → `tianshang_scribe`.** The
  importable top-level package now lives at `src/tianshang_scribe/` (mirroring
  TianshangCAD's `src/tianshangcad/` layout); `src/` is a pure build-isolation
  directory. Imports change from `from src.x import y` to
  `from tianshang_scribe.x import y`; `python -m src.mcp.server` becomes
  `python -m tianshang_scribe.mcp.server`; coverage source is `tianshang_scribe`.
  The `tianshang-scribe` / `scribe-mcp` console scripts are unchanged.
  Migration notes in `MIGRATION.md`.

## [0.4.0] - 2026-08-14

### Added
- M2 cron scheduler + sandboxed script runner + persistent store
  (`src/tianshang_scribe/core/scheduler.py`, `src/tianshang_scribe/core/script_runner.py`,
  `src/tianshang_scribe/utils/store.py`): `--schedule-add/rm/list/run/run-all`,
  `--schedule-db`, `--run-script` CLI flags; `CronTrigger` parsing with
  timezone-aware due computation, subprocess isolation with an allow-listed
  import sandbox, and a WAL-backed SQLite store (56 unit tests).
- M3 document snapshots via `compare_documents` (Discriminated-Union
  action dispatch): `snapshot` / `list_snapshots` / `restore` sub-actions
  persisted as JSON under `~/.tianshang-scribe/snapshots/` with source sha1
  and paragraph state, keeping the tool count at 7.
- M3 RBAC role matrix (`src/tianshang_scribe/mcp/security.py`): `Role` enum (viewer/editor/
  owner), permission levels and a derived tool-permission matrix enforced by
  the `RbacMiddleware` on `tools/call` via the `X-Scribe-Role` header
  (defaults to owner for backward compatibility).
- SEP-1821 Tool Search on the MCP server: the `tools/list` method now accepts
  an optional `query` parameter and returns only matching tools ranked by
  relevance (name/title/description/input-schema weighted scoring), via a
  replacement handler registered through the SDK's `add_request_handler`
  replace semantics (`src/tianshang_scribe/mcp/tool_search.py`, 20 unit tests).
- Centralized runtime configuration via `pydantic-settings`
  (`src/tianshang_scribe/utils/config.py`): the `SCRIBE_*` environment variables, `.env` file
  and CLI defaults (transport, host, port, auth token, API keys, CORS origins,
  rate limits, MCP path) now resolve through a single `Settings` model;
  `src/tianshang_scribe/mcp/auth.py` and the `scribe-mcp` CLI read from it (15 unit tests).

### Changed
- M1 coverage gate raised 70% → 80% with full unit coverage on the CLI
  dispatcher, engines, REPL and MCP; full suite 830 passed, 92.93% coverage
  (12 unit test modules closed in `tests/unit/`).
- `scribe-mcp` CLI defaults (host/port/transport/rate limits/MCP path) now
  come from the `Settings` model instead of hard-coded constants, so they can
  be overridden via `SCRIBE_*` environment variables or a `.env` file.
- HTTP auth middleware now distinguishes missing credentials (`401`, no
  `Authorization` header) from invalid credentials (`403`, bad bearer token);
  `/health` and `/metrics` remain exempt from both auth and rate limiting
  (`src/tianshang_scribe/mcp/transport.py`, 20 unit tests, transport coverage 32%→89%).
- Structured logging via `structlog` (`src/tianshang_scribe/utils/logging.py`): console and
  JSON output modes driven by `SCRIBE_LOG_LEVEL` / `SCRIBE_LOG_JSON`; uvicorn
  access logs unified through the same format; middleware records
  `auth_rejected` / `rate_limited` events (10 unit tests).
- Docker image hardened: multi-stage build, OCI labels, writable tmpdir
  (`TMPDIR=/tmp/scribe`), `PYTHONDONTWRITEBYTECODE`, non-root runtime,
  HEALTHCHECK with start period, and `SCRIBE_*` env passthrough; the default
  transport is now Streamable HTTP (`http://localhost:8080/mcp`) in both the
  Dockerfile and `docker-compose.yml`.
- Test coverage raised above 70% (66.7%→70.6%, 414 tests) with new unit
  suites for `src/tianshang_scribe/transform/pdf.py` (92%) and `src/tianshang_scribe/rendering/template.py`
  (96%); the coverage gate moved from 60% to 70% in `pyproject.toml` and CI.
- Fixed a bug in Excel `{{#each}}` template loops where the first item was
  overwritten and subsequent items copied an already-rendered row; template
  cells are now captured once and the first data row is preserved.

## [0.3.0] - 2026-08-11

### Added
- Engineering baseline aligned with TianshangCAD: `src/tianshang_scribe/utils/` helpers,
  pytest `filterwarnings=error`, ruff expanded rule set, mypy strict for
  new modules, coverage configuration, and shared `tests/conftest.py` fixtures.
- Standard governance documents: `CHANGELOG.md`, `SECURITY.md`,
  `CODE_OF_CONDUCT.md`, `THIRD_PARTY_LICENSES.md`, `MIGRATION.md`.
- MCP server migrated to the official `mcp>=2.0` SDK under `src/tianshang_scribe/mcp/`
  (stdio + Streamable HTTP + SSE), removing the package-name collision with
  the SDK and aligning with TianshangCAD's `src/<pkg>/mcp/` layout.
- MCP hardening modules: `schemas.py` (pydantic models + `as_dict`
  normalization), `auth.py` (Bearer tokens), `rate_limit.py` (token bucket),
  `metrics.py` (Prometheus-style counters/histograms), `security.py`
  (read-only / destructive classification used in tool annotations).
- `src/tianshang_scribe/mcp/transport.py`: stdio / SSE / Streamable HTTP wiring plus ASGI
  middleware (auth, CORS, rate limiting, metrics) for the HTTP transports.
- `src/tianshang_scribe/mcp/prompts.py`: 5 prompt workflows (summarize, batch fill, convert,
  latex polish, compare) exposed via `prompts/list`.
- `src/tianshang_scribe/mcp/tools/_registry.py`: central tool registry; 7 tools now served
  (`create`, `edit`, `fill_template`, `convert`, `extract`, `validate`,
  `compare`). Tool `inputSchema`s are derived from `Annotated` signatures.
- Tool descriptions rewritten to a three-sentence template that discloses
  side effects (writes, `output_path` defaults, in-place overwrite, PDF engine
  dependency) and read-only guarantees, consistent with the `readOnlyHint` /
  `destructiveHint` annotations; server instructions now disclose bearer-token
  auth and rate limits. Guarded by `tests/test_mcp_descriptions.py`.
- `validate_template` / `compare_documents` promoted to first-class tools.
- `fill_template` writes its data payload to a unique temp file instead of a
  shared `mcp_template_data.json`, removing a cross-session race.
- MCP tests: `tests/integration/mcp/test_sse.py` (SDK SSE protocol: initialize +
  tools/list + tools/call), `tests/integration/mcp/mcp_stdio_smoke.py` (9 steps),
  `tests/integration/mcp/mcp_agent_sim.py` (11 end-to-end scenarios) via the shared
  `tests/integration/mcp/_mcp_client.py` stdio client.
- `--extract` full modes: `text` / `tables` / `images` / `structure` /
  `metadata` across Word, Excel and PPT (`extract_text`, `extract_tables`,
  `extract_images`, `extract_structure` on each engine).
- `--add-table`: add Word tables from inline `"H1,H2|a1,a2"` or `@file.csv`
  (`WordEngine.add_table_data`).
- Reverse conversions: `.md`/`.html` inputs auto-convert to Word and `.json`
  to Excel (`src/tianshang_scribe/transform/reverse.py`, `ExcelEngine.import_json`); new
  dependencies `htmldocx` + `markdown`.
- Batch processing: `--batch` flag and `--files "reports/*.docx"` glob;
  per-file execution continues on failure with a summary. Click/Typer's
  Windows argv glob expansion is disabled so patterns reach the CLI literally.
- `--compress-media` (PPT): recompress/resize images via Pillow
  (`PptEngine.compress_media`); `pillow` declared as a direct dependency.
- Interactive file session: `tianshang-scribe open <file>` enters a REPL
  (`src/tianshang_scribe/cli/repl.py` `InteractiveSession`) that holds the document in memory
  and supports `add`/`heading`/`table`/`math`/`replace`/`delete`/`style`/
  `extract`/`info`/`path`/`save`/`help`/`quit`. Explicit `save` persists;
  quitting with unsaved changes prompts first.

### Changed
- Local MCP package relocated from top-level `mcp/` to `src/tianshang_scribe/mcp/`.
  Entry point: `python -m src.mcp.server`; CLI alias `scribe-mcp`.
- `src/tianshang_scribe/utils/file_utils.py` introduced; `check_overwrite` now delegates to it.
- CLI restructured from a single Typer callback group into a one-shot plain
  command plus an `open` subcommand app, dispatched in `main_cli` by the first
  argument. Options may now appear before **or after** the positional
  `input_file` (e.g. `tianshang-scribe file.docx --add "hi"`).
- Shared CLI options (`--latex-style`, `-w/-e/-p`) are defined once as
  constants and reused by both apps so the parameters stay in sync.
- `parse_table_input` moved to `src/tianshang_scribe/cli/global_opts.py` and shared by the
  one-shot `--add-table`, the REPL `table` command, and batch mode.

### Removed
- Top-level `mcp/` namespace package (conflicted with the official `mcp` SDK).
- Legacy `src/tianshang_scribe/mcp/transport_sse.py` (superseded by `transport.py`).

## [0.2.0] - 2026-07

### Added
- MCP Server: SSE transport, Bearer Token auth, CORS whitelist,
  health checks, structured logging, progress notifications.
- LaTeX �?OMML math engine (110+ symbols, `\frac \sqrt \sum \int`).
- Template engine `{{#each}}` loops and `{{#if}}`/`{{#unless}}` conditions.
- `office2pdf` primary PDF engine with LibreOffice fallback.
- PyPI release and PyInstaller EXE packaging.

## [0.1.1] - 2026-07

- PyPI publishing pipeline, EXE packaging, CI 9-matrix, bilingual README.

## [0.1.0] - 2026-07

- MVP: Word/Excel/PPT basic CRUD, LaTeX-style markup, template filling.
