#!/bin/bash
# Description: Setup script for Raspberry Pi
# Creates a Python virtual environment and installs required packages

echo "Starting SEED Program Setup for Raspberry Pi..."

# Check if python3-venv is installed, install if missing
if ! dpkg -s python3-venv >/dev/null 2>&1; then
    echo "Installing python3-venv..."
    sudo apt-get update
    sudo apt-get install -y python3-venv
fi

# Create virtual environment if it doesn't exist
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment '.venv'..."
    python3 -m venv .venv
else
    echo "Virtual environment already exists."
fi

# Activate virtual environment
echo "Activating virtual environment..."
source .venv/bin/activate

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip

# Install requirements
echo "Installing Python dependencies..."
pip install -r requirements.txt

# Additional instructions for the user
echo ""
echo "==========================================================="
echo "Setup Complete!"
echo ""
echo "To run the application, ALWAYS activate the environment first:"
echo "    source .venv/bin/activate"
echo "Then launch it using:"
echo "    python main.py"
echo "==========================================================="
