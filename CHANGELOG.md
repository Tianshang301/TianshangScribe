# Changelog

All notable changes to this project are documented in this file.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added
- REPL environment files: `tianshang-scribe open` now loads rc files
  (`--env-file` wins, then project `.scribe/repl.rc` over user
  `~/.tianshang-scribe/repl.rc`; corrupt files degrade to a warning). The INI
  file supports `[repl] latex_style`, `[aliases]` command aliases and
  `[startup]` commands executed before the interactive loop — default fonts
  are best set via a `style` startup command (`set_style` merges persistently).
- REPL `env` command: `env` shows the effective environment, `env alias <name>
  <command...>` / `env unalias <name>` manage session aliases at runtime
  (single-level expansion; aliasing `env` itself is refused, shadowing other
  built-ins warns).

### Changed
- The interactive session now enters the document's directory on start and
  restores the previous working directory on exit, so relative paths such as
  `table @data.csv`, `path out.docx` and `save rel.docx` resolve against the
  document location instead of the launch directory. Session paths are
  normalized to absolute at construction, keeping no-argument `save` in the
  original document location.

## [0.9.0] - 2026-08-23

### Added
- `compare_excel_workbooks` gains `formula` and `full` diff modes alongside
  `data`: formula mode streams both workbooks read-only and compares formula
  strings only (O(1) memory, hash pre-filter, size caps with a truncation
  warning), full mode reports every stored cell including cached values.
  Empirical note: openpyxl `read_only=True, data_only=False` does return
  formula strings, so PLAN.md D-5's assumption was corrected in the process.
- `extract_presentation_data` gains `notes` mode (full speaker-notes text per
  slide) and `master_info` mode (masters/layouts inventory with placeholder
  types).
- `analyze_excel_data` gains a `pivot_suggestion` mode proposing rows /
  columns / values / aggregation from per-column type inference.
- PowerPoint engine: `apply_theme` built-in themes (`office`, `dark`) rewrite
  `theme1.xml` color/font schemes directly; `add_movie` / `add_audio` insert
  click-to-play media (poster frame support for movies); master-level
  `set_master_options` toggles slide-number / footer / date placeholders.
- Excel engine: row/column grouping `group_rows` / `group_columns` / `ungroup`
  (outline level 1-7, optional collapse), `set_tab_color`, `set_print_area`,
  and a simplified `set_page_setup` (paper size name or raw int, orientation,
  margins subset, centred header/footer text).
- L2 schema wiring under the D-7 capability freeze: legacy `EditOperation`
  absorbs only four core actions (`group_rows`, `group_columns`, `ungroup`,
  `add_media`); edge capabilities stay exclusive to the dedicated tools via
  `ExcelEditOp` (17 actions) and `PptEditOp` (13 actions). Both wrappers share
  one edit session shell (`run_edit_session`) so operation order, backup and
  error mapping stay uniform.
- Cross-cutting error refinements: `error_response` accepts optional `field`
  and `documentation_url`; four refined codes — `EXCEL_INVALID_CELL_REF`
  (1007), `EXCEL_INVALID_RANGE` (1008), `PPT_INVALID_SLIDE_INDEX` (1009),
  `EXCEL_SHEET_NOT_FOUND` (1010); live engine errors are classified onto them
  with the offending `operations[i].field`. Formula syntax is deliberately not
  validated.
- New `_dryrun.py` pre-flight module: structural cell/range validation, sheet
  existence and slide-index bounds (with running `add_sheet` / `add_slide`
  context), per-operation impact estimates. `edit_*` and
  `create_office_document` dry-run responses now carry `validations`,
  `all_valid`, and `estimated_impacted_cells` next to their legacy keys.

### Changed
- Bumped `SERVER_VERSION` and package version to `0.9.0`; docs and AGENTS.md
  feature matrix now list 14 tools.

## [0.8.1] - 2026-08-22

### Fixed
- `parse_conditional_format` / `parse_data_validation`
  (`src/tianshang_scribe/mcp/tools/_parse.py`): the spec remainder is now split
  at most twice, so conditional-format formulas containing colons (e.g. time
  literals like `10:00`) survive intact and extra `:`-segments of a
  data-validation spec stay inside `formula2` instead of being dropped.

### Changed
- PowerPoint `add_textbox`: the `top` parameter is now optional (`None`);
  when omitted the box auto-stacks below previously placed boxes on the same
  slide via the internal text cursor, so consecutive text blocks never
  overlap. Explicit tops pin the box as before
  (`src/tianshang_scribe/core/ppt_engine.py`). The dedicated MCP tool's
  `PptTextBlock.top` defaults to `None` accordingly, and
  `create_presentation` no longer hardcodes `top=1.0`.
- MCP `write_cell` gains an optional `is_formula` flag (both on the unified
  `EditOperation` and the dedicated `ExcelEditOp`): `true` stores the string
  as a formula (must start with `=`), `false` forces a literal string cell
  even when it starts with `=`, omitted keeps the automatic behaviour —
  removing the ambiguity of writing `=`-prefixed plain text
  (`src/tianshang_scribe/mcp/tools/edit.py`,
  `src/tianshang_scribe/mcp/tools/excel_edit.py`).
- Registry description of `edit_office_document` now marks it as the legacy
  general-purpose editor and guides Agents to the dedicated
  `edit_excel_workbook` / `edit_presentation` tools (P1-010 plan A); the
  Excel/PPT capability advertising moved into those tools' descriptions.
- Bumped `SERVER_VERSION` and package version to `0.8.1`.

## [0.8.0] - 2026-08-22

### Fixed
- PowerPoint `merge_workbooks`: slides are now deep-cloned with their content
  (text, pictures, media, charts) via relationship-aware copying instead of
  creating blank slides from the layout (`src/tianshang_scribe/core/ppt_engine.py`).
- PowerPoint `add_text` / `add_styled_content`: multiple text blocks, headings and
  formulas are placed on distinct, vertically stacked text boxes (previously they
  all overlapped at the same fixed position); `add_text` gains a `slide_index`
  argument to append onto an existing slide's body placeholder.
- PowerPoint password protection: `set_protection` now writes a compliant
  ECMA-376 agile `modifyVerifier` (SHA-512 + random salt + 100k iterations,
  base64-encoded) instead of storing the plaintext password.
- PowerPoint `to_images`: exports **every** slide by rendering the deck to PDF
  first and rasterizing each page (PyMuPDF, else poppler `pdftoppm`), fixing the
  LibreOffice `--convert-to png` path that only produced the first slide.
- Excel `sort`: supports multi-column keys (`key_columns` / `orders`) and sorts
  whole rows intact; mixed value types are ordered deterministically
  (numbers < strings < other < None) instead of raising `TypeError`.

### Added
- Excel `--sheet` option to target a specific worksheet for the subsequent
  operations (write, formula, sort, chart, import/export, comment); the engine
  exposes `select_sheet` / `_ws()` (`src/tianshang_scribe/cli/main.py`,
  `src/tianshang_scribe/core/excel_engine.py`).
- Excel engine capabilities: `freeze_panes` (CLI `--freeze`), `set_number_format`
  (CLI `--number-format`), `add_conditional_format` (CLI `--conditional-format`,
  supporting `color_scale` / `data_bar` / `cell_is` / `formula`), `add_data_validation`
  (CLI `--data-validation`, `list`/`whole`/`decimal`/`date`/`text_length`),
  `set_range_style` (border/fill), chart-type extension (`area`/`scatter`/`doughnut`
  alongside `bar`/`line`/`pie`), `add_hyperlink`, `set_named_range`, `auto_fit`.
- PowerPoint engine capabilities: `add_textbox` (precise inch positioning),
  `add_table` (CLI `--ppt-table`), `add_chart` (CLI `--ppt-chart`,
  `bar`/`column`/`line`/`pie`/`area`/`doughnut`), `add_picture`, `add_shape`
  (autoshapes), and `replace_text` now preserves run-level styles across runs
  that span a match.
- MCP Server: the existing 7 tools now expose the Excel/PPT engine capabilities
  above. `create_office_document`'s `ContentBlock` and `edit_office_document`'s
  `EditOperation` gained optional fields (`slide_index`, `slide_layout`, `notes`,
  `transition`, `sheet_name`, `cell`, `formula`, `chart_type`, `chart_data_range`,
  `chart_data`, `rows`, `number_format`, `conditional_format`, `data_validation`,
  `freeze`, `hyperlink`, `named_range`) so agents can build formula cells, frozen
  panes, number formats, conditional formats, data validation, Excel/PPT charts,
   PPT tables/pictures/shapes/layouts/transitions/notes. Fixed a crash where
   `create_office_document` raised on a PPT `table` block (signature mismatch with
   `PptEngine.add_table`); PPT content blocks now stack onto a single slide.
   Parsing helpers live in `src/tianshang_scribe/mcp/tools/_parse.py`.
- MCP Server hardening (post-review P0/P1):
  - Fixed `_edit_write_cell` writing to the active sheet instead of the
    explicitly selected sheet (now resolves the target worksheet per-op without
    mutating engine selection); a `write_cell` value starting with `=` is stored
    as a formula.
  - PPT blocks in `create_office_document` now default to the current stacked
    slide (`current_slide_index`) unless an explicit `slide_index` is given
    (table and capability fields included); `parse_ppt_chart` no longer returns a
    dead chart-type hint.
  - `compare_documents` / snapshot now return a clearer `UNSUPPORTED_FORMAT`
    message stating Excel/PPT comparison is not yet available.
  - `EditOperation` schema documents the per-action field mapping; the registry
    tool descriptions for `create_office_document` / `edit_office_document` /
    `compare_documents` now mention the new Excel/PPT capabilities and limits.
  - `write_cell` honors an optional `style` (font/fill/alignment) applied to the
    target cell.
- Excel `set_range_style` now supports font name/size/bold/italic/color,
  horizontal alignment and number format (previously border/fill only),
  matching the cell-level style system.
- PowerPoint `add_textbox` / internal `_place_textbox` derive the default width
  from the slide width so 4:3 decks no longer overflow.
- MCP Server: five dedicated, document-type-specific tools (12 tools total,
  the 7 unified tools remain as legacy):
  - `create_excel_workbook` builds a new .xlsx from typed `ExcelSheetSpec`
    sheets (headers/rows/formulas/freeze/number_format/conditional_format/
    data_validation/column_widths) instead of the generic `ContentBlock`.
  - `edit_excel_workbook` applies typed `ExcelEditOp` operations
    (`write_cell`/`set_formula`/`freeze_panes`/`add_chart`/`conditional_format`/
    `data_validation`/`add_table`/`sort`/`add_sheet`/`set_range_style`/
    `number_format`) by reusing the unified edit dispatch.
  - `create_presentation` builds a new .pptx from typed `PptSlideSpec` slides
    (layout/title/bullets/text_blocks/table/chart/picture/notes/transition).
  - `edit_presentation` applies typed `PptEditOp` operations (`add_slide`,
    `add_text`, `replace_text`, `add_table`, `add_chart`, `add_picture`,
    `add_shape`, `apply_layout`, `set_transition`, `add_notes`).
  - `analyze_excel_data` profiles a workbook read-only: per-sheet row/column
    counts, headers, inferred column types (numeric min/max/mean, categorical),
    null counts, sample rows, and duplicate-row detection.
  - Permission classes (STANDARD×2 / DESTRUCTIVE×2 / READ_ONLY×1), RBAC and
    idempotency tables, description gates (≤90 words, side-effect disclosure,
    read-only affirmation, sibling pointers) and stdio/SSE smoke assertions all
    extended to the 12-tool registry; `SERVER_VERSION` stays `0.8.0`.
- Bumped `SERVER_VERSION` to `0.8.0` so the MCP health endpoint reports the
  current release.

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
