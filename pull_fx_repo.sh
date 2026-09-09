#!/bin/bash
# Pulls the latest commit from the fx-rate-sync repo.
# Assumes this repo has already been cloned into your OneDrive-synced folder.

REPO_DIR="$HOME/OneDrive/fx-rate-sync"   # <-- update this path to match your OneDrive folder location

cd "$REPO_DIR" || { echo "ERROR: repo folder not found at $REPO_DIR"; exit 1; }
git pull origin main >> "$HOME/fx-pull.log" 2>&1
