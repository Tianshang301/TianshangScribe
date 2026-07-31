# TianshangScribe Docker Image
# Multi-stage: builder for deps, runtime for clean image

FROM python:3.12-slim AS builder
WORKDIR /app
COPY pyproject.toml ./
COPY src/ ./src/
COPY mcp/ ./mcp/
RUN pip install --no-cache-dir -e .

FROM python:3.12-slim
WORKDIR /app
COPY --from=builder /usr/local/lib/python3.12/site-packages /usr/local/lib/python3.12/site-packages
COPY --from=builder /app/src ./src
COPY --from=builder /app/mcp ./mcp

# Install office2pdf for PDF conversion (optional)
# RUN curl -L https://github.com/xxx/office2pdf/releases/latest/download/office2pdf-linux -o /usr/local/bin/office2pdf && chmod +x /usr/local/bin/office2pdf

# SSE MCP Server (default)
EXPOSE 8080
ENV PYTHONUNBUFFERED=1

# Switch to non-root user
RUN useradd -m scribe && chown -R scribe:scribe /app
USER scribe

HEALTHCHECK --interval=30s --timeout=3s \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8080/health')" || exit 1

ENTRYPOINT ["python", "-m", "mcp.server", "--transport", "sse", "--host", "0.0.0.0"]
