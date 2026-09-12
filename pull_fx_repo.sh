#!/bin/bash
MIRROR_DIR="$HOME/fx-rate-sync-automation"
ONEDRIVE_FILE="$HOME/Library/CloudStorage/OneDrive-HeritageCeramicsPteLtd/Documents - Work/FX_Rate_Sync/fx-rate-sync/SGDUSD_Tracker.xlsx"
LOG="$HOME/Scripts/fx-pull.log"

{
  cd "$MIRROR_DIR" || { echo "ERROR: mirror folder not found at $MIRROR_DIR"; exit 1; }
  git pull origin main
  cp "$MIRROR_DIR/SGDUSD_Tracker.xlsx" "$ONEDRIVE_FILE"
  echo "Copied to OneDrive at $(date)"
} >> "$LOG" 2>&1
