#!/bin/bash
# Run tests for MCP tools
# Usage: ./run_tests.sh [pytest args]

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
MCP_DIR="$(dirname "$SCRIPT_DIR")"

cd "$MCP_DIR"

echo "=========================================="
echo "Running ruff linter..."
echo "=========================================="
python -m ruff check tools/ tests/ --fix || true
python -m ruff format tools/ tests/ || true

echo ""
echo "=========================================="
echo "Running pytest..."
echo "=========================================="
python -m pytest tests/ "$@"

echo ""
echo "=========================================="
echo "Tests complete!"
echo "=========================================="
