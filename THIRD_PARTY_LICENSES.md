# Third-Party Licenses

TianshangScribe (Apache-2.0) bundles or depends on the following third-party
packages. Each is used under its own license; full texts are distributed with
the respective packages.

| Package       | License        | Purpose                              |
| ------------- | -------------- | ------------------------------------ |
| typer         | MIT            | CLI framework                        |
| rich          | MIT            | Terminal output formatting           |
| python-docx   | MIT            | Word (.docx) read/write              |
| openpyxl      | MIT            | Excel (.xlsx) read/write/formulas    |
| python-pptx   | MIT            | PowerPoint (.pptx) operations        |
| lxml          | BSD-3-Clause   | XML/Office XML processing            |
| mammoth       | BSD-2-Clause   | Word → HTML conversion               |
| mcp           | MIT            | Model Context Protocol server SDK    |
| pydantic      | MIT            | Data validation / schemas            |
| pydantic-settings | MIT        | Centralized config via env / .env    |
| structlog     | Apache-2.0     | Structured logging                   |
| uvicorn       | BSD-3-Clause   | ASGI server for HTTP transports      |
| deepdiff      | MIT            | Document version differencing        |
| prometheus-client | Apache-2.0 | Metrics endpoint                     |
| APScheduler   | MIT            | Batch scheduling                     |
| Jinja2        | BSD-3-Clause   | Template scripting                   |
| PyYAML        | MIT            | YAML template data                   |
| htmldocx      | MIT            | HTML → Word reverse conversion       |
| markdown      | BSD-3-Clause   | Markdown → HTML for MD → Word        |
| pillow        | HPND           | Image (re)compression for PPT media   |

## Compliance Notes

- This project is distributed under the **Apache License 2.0**. See `LICENSE`.
- MIT/BSD components retain their own copyright notices; they are preserved in
  the installed package metadata and in `THIRD_PARTY_LICENSES.md`.
- The MCP SDK is used per its MIT license for the server implementation under
  `src/tianshang_scribe/mcp/`.

If you redistribute this software, please keep this notice and the license
files of the bundled dependencies.
