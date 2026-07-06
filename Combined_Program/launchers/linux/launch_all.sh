#!/usr/bin/env bash
SCRIPT_DIR="$(dirname "$0")"
bash "$SCRIPT_DIR/run_fake_printer_simulator.sh" &
bash "$SCRIPT_DIR/run_fake_hub_simulator.sh" &
bash "$SCRIPT_DIR/run_seed_control_center.sh"
