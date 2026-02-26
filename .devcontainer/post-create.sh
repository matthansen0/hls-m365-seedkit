#!/usr/bin/env bash
# ──────────────────────────────────────────────────────────────
# post-create.sh — Dev container post-create setup
# Runs automatically when the container is first created.
# ──────────────────────────────────────────────────────────────
set -euo pipefail

echo "──────────────────────────────────────────────────"
echo "  HLS-M365-Seed — Post-Create Setup"
echo "──────────────────────────────────────────────────"

# ── Install the project in editable mode ─────────────────────
echo "▶ Installing m365seed in editable mode (with dev extras)…"
pip install -e ".[dev]"

# ── Verify CLI is available ──────────────────────────────────
echo "▶ Verifying m365seed CLI…"
if command -v m365seed &>/dev/null; then
    echo "  ✓ m365seed CLI installed: $(m365seed --help | head -1)"
else
    echo "  ✗ m365seed CLI not found on PATH — check pyproject.toml [project.scripts]"
    exit 1
fi

# ── Verify Azure CLI ────────────────────────────────────────
echo "▶ Verifying Azure CLI…"
if command -v az &>/dev/null; then
    echo "  ✓ Azure CLI $(az version --query '\"azure-cli\"' -o tsv)"
else
    echo "  ⚠ Azure CLI not available — some setup wizard features may be limited"
fi

# ── Copy example config if no config exists ──────────────────
if [ ! -f seed-config.yaml ]; then
    echo "▶ No seed-config.yaml found — copying example config…"
    cp seed-config.example.yaml seed-config.yaml
    echo "  ✓ Created seed-config.yaml from example"
    echo "  → Edit seed-config.yaml with your tenant details, or run:"
    echo "    m365seed setup"
else
    echo "▶ seed-config.yaml already exists — skipping copy."
fi

# ── Run tests to verify environment ─────────────────────────
echo "▶ Running tests to verify environment…"
if python -m pytest tests/ -q --tb=short 2>/dev/null; then
    echo "  ✓ All tests passing"
else
    echo "  ⚠ Some tests failed — check output above"
fi

echo ""
echo "──────────────────────────────────────────────────"
echo "  Setup complete! Next steps:"
echo ""
echo "  1. Run the interactive setup wizard:"
echo "     m365seed setup"
echo ""
echo "  2. Or manually edit seed-config.yaml and run:"
echo "     m365seed validate -c seed-config.yaml"
echo "     m365seed seed-all -c seed-config.yaml --dry-run"
echo "──────────────────────────────────────────────────"
