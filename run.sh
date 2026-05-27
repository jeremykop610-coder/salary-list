#!/usr/bin/env bash
set -euo pipefail

PYTHON="/Users/jeremyzhu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
exec "$PYTHON" app.py
