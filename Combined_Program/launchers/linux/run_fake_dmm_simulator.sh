#!/usr/bin/env bash
# dmm_manager.py's DMM_LIST hard-codes six instruments at IDs 20-25 (VINV,
# VINP, IINP, VSYS, SAUX, SINV). Each needs its own simulator process bound
# to 127.0.0.<id>:5025 — one loopback octet per instrument, fixed port 5025.
# Point the real app's "DMM IP Prefix" field (Sender > DMM panel) at 127.0.0
# (no port) to reach all six. No code changes needed — DmmInst.connect()
# already accepts ip_prefix as a parameter.
#
# Usage: run_fake_dmm_simulator.sh            # launches all 6 real channels
#        run_fake_dmm_simulator.sh <id> <name> # launches one ad-hoc instance
set -euo pipefail
cd "$(dirname "$0")/../.."

PYTHON="python3"
PYTHONPATH_VAL="/home/sheamcg/RLM/venv-linux"

if [ $# -ge 1 ]; then
    DMM_ID="$1"
    DMM_NAME="${2:-DMM}"
    echo "Starting Fake DMM Simulator (127.0.0.$DMM_ID:5025, name=$DMM_NAME)..."
    MPLCONFIGDIR="$TMPDIR" PYTHONPATH="$PYTHONPATH_VAL" $PYTHON sim/fake_dmm.py --id "$DMM_ID" --name "$DMM_NAME"
    exit 0
fi

echo "Starting 6 Fake DMM Simulators (127.0.0.20-25:5025)..."
declare -A CHANNELS=( [20]=VINV [21]=VINP [22]=IINP [23]=VSYS [24]=SAUX [25]=SINV )
for id in "${!CHANNELS[@]}"; do
    MPLCONFIGDIR="$TMPDIR" PYTHONPATH="$PYTHONPATH_VAL" $PYTHON sim/fake_dmm.py --id "$id" --name "${CHANNELS[$id]}" &
done
wait
