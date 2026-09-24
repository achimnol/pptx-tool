# syntax=docker/dockerfile:1.7
# Web UI image of pptx-tool: the package with the `web` extra plus the built frontend.
# The chart in lablup/internal-deploy runs `pptx-tool serve ...` explicitly, so the entrypoint
# only has to make `pptx-tool` available on PATH.

# --- Frontend ---------------------------------------------------------------------------------
FROM node:24-bookworm-slim AS frontend
WORKDIR /src/frontend
# packageManager in package.json pins the pnpm version for corepack.
RUN corepack enable pnpm
COPY frontend/package.json frontend/pnpm-lock.yaml ./
RUN --mount=type=cache,target=/root/.local/share/pnpm/store \
    pnpm install --frozen-lockfile
COPY frontend/ ./
# vite writes to ../pptx_tool/web/static (see vite.config.ts).
RUN pnpm build

# --- Python package ---------------------------------------------------------------------------
FROM ghcr.io/astral-sh/uv:python3.13-bookworm-slim AS build
ENV UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    UV_PYTHON_DOWNLOADS=never
# The venv scripts hard-code this path in their shebang, so the runtime stage copies it to the same place.
WORKDIR /app
COPY pyproject.toml uv.lock README.md LICENSE ./
# Dependencies first, so that source changes do not invalidate this layer.
RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --frozen --no-dev --extra web --no-install-project
COPY pptx_tool/ ./pptx_tool/
COPY --from=frontend /src/pptx_tool/web/static/ ./pptx_tool/web/static/
RUN --mount=type=cache,target=/root/.cache/uv \
    uv sync --frozen --no-dev --extra web --no-editable

# --- Runtime ----------------------------------------------------------------------------------
FROM python:3.13-slim-bookworm
ENV PATH="/app/.venv/bin:$PATH" \
    PYTHONUNBUFFERED=1
RUN groupadd --gid 10001 app && useradd --uid 10001 --gid app --no-create-home --shell /usr/sbin/nologin app
COPY --from=build --chown=root:root /app/.venv /app/.venv
USER 10001:10001
EXPOSE 8000
ENTRYPOINT ["pptx-tool"]
# The loopback default is useless inside a container; the chart overrides these anyway.
CMD ["serve", "--host=0.0.0.0", "--port=8000"]
