# ──────────────────────────────────────────────────────────────
# HLS-M365-Seed Dev Container
# All deployment dependencies pre-installed for CLI-driven seeding.
# ──────────────────────────────────────────────────────────────
FROM mcr.microsoft.com/devcontainers/python:1-3.12-bullseye

# Avoid interactive prompts during package installation
ENV DEBIAN_FRONTEND=noninteractive

# ── System packages ──────────────────────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
        curl \
        gnupg \
        lsb-release \
        apt-transport-https \
        jq \
    && rm -rf /var/lib/apt/lists/*

# ── Azure CLI (latest stable) ───────────────────────────────
RUN curl -sL https://aka.ms/InstallAzureCLIDeb | bash

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
