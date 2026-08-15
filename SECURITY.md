# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 0.6.x   | :white_check_mark: |
| 0.5.x   | :white_check_mark: |
| 0.4.x   | :white_check_mark: |
| < 0.4   | :x:                |

## Reporting a Vulnerability

Please **do not** open a public issue for security vulnerabilities.
Instead, report privately via email to the maintainers listed on the
GitHub repository, or open a private security advisory at
<https://github.com/Tianshang301/TianshangScribe/security/advisories/new>.

Include, if possible:

- Affected version(s) and platforms.
- Steps to reproduce (minimal example).
- Impact description and any suggested mitigation.

We aim to acknowledge reports within 3 business days and to ship a fix
in a patch release once validated.

## Security Considerations for This Project

- **MCP Server**: never deploy with `TIANSHANG_SCRIBE_AUTH_TOKEN` unset in
  non-loopback environments. Use the Bearer token auth option
  (`--auth-token`) and restrict CORS origins (`--cors-origins`).
- **Untrusted documents**: office documents are ZIP archives and may
  contain crafted XML. Always validate inputs from untrusted sources.
- **Script sandbox** (`script_runner`): only run user scripts that you
  authored or reviewed; they execute with your process privileges.
- **Secrets**: never commit tokens or credentials. Prefer environment
  variables or secret managers.

## Dependency Vulnerability Management

Dependency versions are pinned in `pyproject.toml`. We monitor
GitHub Advisory Database and update `THIRD_PARTY_LICENSES.md`
alongside security-relevant upgrades.
