#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
python3 "$SCRIPT_DIR/scripts/install_mxm_tool.py" --repo-dir "$SCRIPT_DIR" --run-gui "$@"
