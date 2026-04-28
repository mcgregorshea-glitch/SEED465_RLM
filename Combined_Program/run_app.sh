#!/bin/bash
# Description: Launcher script for the SEED Combined Program
# This script ensures the script runs from its directory and activates the venv

cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
    echo "Virtual environment not found! Please run setup_pi.sh first."
    echo "Press Enter to exit..."
    read -r
    exit 1
fi

source .venv/bin/activate
python3 main.py
