#!/bin/bash
# Description: Robust setup script for Raspberry Pi
# Creates a Python virtual environment and installs required packages

set -e # Exit immediately if a command exits with a non-zero status

# If not running in a terminal, open one and relaunch this script
if [ ! -t 1 ]; then
    # We are likely running from a desktop double-click without terminal
    exec lxterminal -e bash "$0" "$@"
    exit 0
fi

echo "Starting SEED Program Setup for Raspberry Pi..."

echo "----------------------------------------"
echo "Step 0: Updating code from GitHub..."
echo "----------------------------------------"
git fetch origin || echo "Warning: Could not fetch from origin. Check network."
git reset --hard origin/master || echo "Warning: Could not reset to origin/master."
echo "Update complete."
echo "----------------------------------------"

# 1. Ensure we are in the correct directory
if [ ! -f "requirements.txt" ]; then
    echo "ERROR: requirements.txt not found in current directory!"
    echo "Please run this script from inside the Combined_Program folder."
    exit 1
fi

# 2. Check if python3-venv is installed
if ! dpkg -s python3-venv >/dev/null 2>&1; then
    echo "Installing python3-venv (requires sudo privileges)..."
    sudo apt-get update
    sudo apt-get install -y python3-venv
fi

# 3. Create virtual environment
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment '.venv'..."
    python3 -m venv .venv
else
    echo "Virtual environment already exists."
fi

# 4. Verify .venv was created successfully
if [ ! -f ".venv/bin/activate" ]; then
    echo "ERROR: Virtual environment was not created successfully."
    echo "Try running: sudo apt install python3-venv -y"
    exit 1
fi

# 5. Activate environment
echo "Activating virtual environment..."
source .venv/bin/activate

# 6. Install dependencies
echo "Upgrading pip and installing requirements..."
pip install --upgrade pip
pip install -r requirements.txt

echo ""
echo "==========================================================="
echo "Setup Complete!"
echo "To run the app: source .venv/bin/activate && python main.py"
echo "==========================================================="

# Keep the window open so we can see the result if launched from desktop
read -p "Press Enter to Launch Program..."
python main.py
exit 0
