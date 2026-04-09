# ──────────────────────────────────────────────────────────────
# hls-m365-seedkit Dev Container
# All deployment dependencies pre-installed for CLI-driven seeding.
# ──────────────────────────────────────────────────────────────
FROM mcr.microsoft.com/devcontainers/python:1-3.12-bookworm

# Avoid interactive prompts during package installation
ENV DEBIAN_FRONTEND=noninteractive

# ── Pre-install Python build tools ──────────────────────────
RUN pip install --no-cache-dir --upgrade pip setuptools wheel

# ── Copy project files for dependency install ────────────────
COPY pyproject.toml /tmp/build/pyproject.toml
COPY README.md /tmp/build/README.md
RUN mkdir -p /tmp/build/m365seed && touch /tmp/build/m365seed/__init__.py

# ── Install project dependencies (runtime + dev) ────────────
RUN pip install --no-cache-dir -e "/tmp/build[dev]" \
    && rm -rf /tmp/build

# ── Labels ───────────────────────────────────────────────────
LABEL maintainer="HLS Demo Team" \
      description="M365 Demo Tenant Seeding Tool — Dev Container" \
      python.version="3.12"
