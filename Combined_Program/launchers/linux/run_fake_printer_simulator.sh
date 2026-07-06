#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")/../.."

PYTHON="python3"
PYTHONPATH_VAL="/home/sheamcg/RLM/venv-linux"
PORT="${1:-/tmp/ttyV0}"

while true; do
    echo "Starting Fake Printer Emulator (port: $PORT)..."
    MPLCONFIGDIR="$TMPDIR" PYTHONPATH="$PYTHONPATH_VAL" $PYTHON sim/fake_printer.py --port "$PORT"
    echo
    echo "Application exited. Press Enter to restart, or Ctrl+C to quit."
    read -r
    clear
done
