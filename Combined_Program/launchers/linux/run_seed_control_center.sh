#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")/../.."

PYTHON="python3"
PYTHONPATH_VAL="/home/sheamcg/RLM/venv-linux"

while true; do
    echo "Launching SEED Control Center..."
    MPLCONFIGDIR="$TMPDIR" PYTHONPATH="$PYTHONPATH_VAL" $PYTHON src/main.py
    echo
    echo "Application exited. Press Enter to restart, or Ctrl+C to quit."
    read -r
    clear
done
