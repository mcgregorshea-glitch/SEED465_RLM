#!/usr/bin/env bash
# Stands up the fake printer + fake hub simulators behind socat PTY pairs,
# then launches the real SEED Control Center pre-wired to talk to them.
#
# Usage:
#   run_dev_sim_stack.sh          # (re)start the whole stack, returns immediately
#   run_dev_sim_stack.sh --stop   # tear down a running stack
#
# See seed-control-center/CLAUDE.md "Dev Simulator Stack" for details.
set -uo pipefail
cd "$(dirname "$0")/../.."

PYTHON="python3"
PYTHONPATH_VAL="/home/sheamcg/RLM/venv-linux"

STACK_DIR="/tmp/seed_sim_stack"
PIDFILE="$STACK_DIR/stack.pid"

PRINTER_SIM_PORT="/tmp/ttyV0"   # sim side, fed to fake_printer.py
PRINTER_APP_PORT="/tmp/ttyV1"   # app side, exported as SEED_SIM_PORT
HUB_SIM_PORT="/tmp/ttyH0"       # sim side, fed to fake_hub.py
HUB_APP_PORT="/tmp/ttyH1"       # app side, exported as SEED_SIM_HUB_PORT

stop_stack() {
    if [[ -f "$PIDFILE" ]]; then
        echo "Stopping existing sim stack..."
        while read -r pid; do
            [[ -n "$pid" ]] && kill "$pid" 2>/dev/null
        done < "$PIDFILE"
        sleep 1
        while read -r pid; do
            [[ -n "$pid" ]] && kill -9 "$pid" 2>/dev/null
        done < "$PIDFILE"
        rm -f "$PIDFILE"
    fi
    rm -f "$PRINTER_SIM_PORT" "$PRINTER_APP_PORT" "$HUB_SIM_PORT" "$HUB_APP_PORT"
}

if [[ "${1:-}" == "--stop" ]]; then
    stop_stack
    echo "Stack stopped."
    exit 0
fi

# Force-restart: always tear down any prior stack before starting fresh.
stop_stack
mkdir -p "$STACK_DIR"
: > "$PIDFILE"

add_pid() { echo "$1" >> "$PIDFILE"; }

echo "Starting socat PTY pairs..."
socat -d -d pty,raw,echo=0,link="$PRINTER_SIM_PORT" pty,raw,echo=0,link="$PRINTER_APP_PORT" \
    > "$STACK_DIR/socat_printer.log" 2>&1 &
add_pid $!
socat -d -d pty,raw,echo=0,link="$HUB_SIM_PORT" pty,raw,echo=0,link="$HUB_APP_PORT" \
    > "$STACK_DIR/socat_hub.log" 2>&1 &
add_pid $!

for i in $(seq 1 20); do
    [[ -e "$PRINTER_SIM_PORT" && -e "$PRINTER_APP_PORT" && -e "$HUB_SIM_PORT" && -e "$HUB_APP_PORT" ]] && break
    sleep 0.2
done
if [[ ! -e "$PRINTER_APP_PORT" || ! -e "$HUB_APP_PORT" ]]; then
    echo "ERROR: socat PTY pairs did not come up. Is socat installed? (sudo apt-get install -y socat)"
    stop_stack
    exit 1
fi

if [[ -z "${DISPLAY:-}" ]]; then
    echo "No DISPLAY set, starting Xvfb on :99..."
    Xvfb :99 -screen 0 1280x900x24 > "$STACK_DIR/xvfb.log" 2>&1 &
    add_pid $!
    sleep 1
    export DISPLAY=:99
fi

echo "Starting Fake Printer Emulator (port: $PRINTER_SIM_PORT)..."
MPLCONFIGDIR="${TMPDIR:-/tmp}" PYTHONPATH="$PYTHONPATH_VAL" $PYTHON sim/fake_printer.py --port "$PRINTER_SIM_PORT" \
    > "$STACK_DIR/fake_printer.log" 2>&1 &
add_pid $!

echo "Starting Fake Hub Simulator (port: $HUB_SIM_PORT)..."
MPLCONFIGDIR="${TMPDIR:-/tmp}" PYTHONPATH="$PYTHONPATH_VAL" $PYTHON sim/fake_hub.py --port "$HUB_SIM_PORT" \
    > "$STACK_DIR/fake_hub.log" 2>&1 &
add_pid $!

echo "Launching SEED Control Center (SEED_SIM_PORT=$PRINTER_APP_PORT, SEED_SIM_HUB_PORT=$HUB_APP_PORT)..."
MPLCONFIGDIR="${TMPDIR:-/tmp}" PYTHONPATH="$PYTHONPATH_VAL" \
    SEED_SIM_PORT="$PRINTER_APP_PORT" SEED_SIM_HUB_PORT="$HUB_APP_PORT" \
    $PYTHON src/main.py > "$STACK_DIR/app.log" 2>&1 &
add_pid $!

echo "Sim stack up. PIDs recorded in $PIDFILE."
echo "  Printer: sim=$PRINTER_SIM_PORT  app=$PRINTER_APP_PORT (SEED_SIM_PORT)"
echo "  Hub:     sim=$HUB_SIM_PORT  app=$HUB_APP_PORT (SEED_SIM_HUB_PORT)"
echo "  Logs:    $STACK_DIR/"
echo "Stop with: $0 --stop"
