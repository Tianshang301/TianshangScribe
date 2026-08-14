# TianshangScribe Docker Image
# Multi-stage build: builder installs deps, runtime ships a slim, non-root image.

FROM python:3.12-slim AS builder
WORKDIR /app
COPY pyproject.toml README.md LICENSE ./
COPY src/ ./src/
RUN pip install --no-cache-dir -e .
FROM python:3.12-slim
LABEL org.opencontainers.image.title="tianshang-scribe" \
      org.opencontainers.image.description="Cross-platform Office document processing CLI + MCP Server" \
      org.opencontainers.image.licenses="Apache-2.0"

WORKDIR /app
COPY --from=builder /usr/local/lib/python3.12/site-packages /usr/local/lib/python3.12/site-packages
COPY --from=builder /usr/local/bin /usr/local/bin
COPY --from=builder /app/src ./src

# Runtime tuning: unbuffered stdout, no bytecode cache in image, temp dir under
# the writable volume so document scratch files survive container restarts.
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    TMPDIR=/tmp/scribe \
    SCRIBE_LOG_LEVEL=${SCRIBE_LOG_LEVEL:-INFO} \
    SCRIBE_LOG_JSON=${SCRIBE_LOG_JSON:-0}

# Streamable HTTP MCP Server (default transport; override via CMD)
EXPOSE 8080
ENV SCRIBE_TRANSPORT=streamable-http

# Non-root user; scratch/output dir owned by that user.
RUN useradd -m scribe && mkdir -p /tmp/scribe && chown -R scribe:scribe /app /tmp/scribe
USER scribe

HEALTHCHECK --interval=30s --timeout=3s --start-period=10s --retries=3 \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8080/health', timeout=2)" || exit 1

ENTRYPOINT ["python", "-m", "tianshang_scribe.mcp.server", "--host", "0.0.0.0"]
CMD ["--transport", "streamable-http"]