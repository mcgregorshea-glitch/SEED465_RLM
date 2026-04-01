#!/bin/bash

# Update script for SEED465_RLM
# This script pulls the latest changes from the GitHub repository.

echo "Updating code from GitHub..."

# Fetch all changes and reset to origin/master to ensure we are exactly in sync with the remote
# (This handles cases where local changes might otherwise cause a merge conflict)
git fetch origin
git reset --hard origin/master

if [ $? -eq 0 ]; then
    echo "----------------------------------------"
    echo "Update Successful! Your code is now in sync with GitHub."
    echo "----------------------------------------"
else
    echo "----------------------------------------"
    echo "ERROR: Update failed. Please check your internet connection or git status."
    echo "----------------------------------------"
    exit 1
fi
