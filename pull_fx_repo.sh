#!/bin/bash
MIRROR_DIR="$HOME/FX Rate Sync - Do not delete/repo"
ONEDRIVE_FILE="$HOME/Library/CloudStorage/OneDrive-HeritageCeramicsPteLtd/Documents - Work/FX_Rate_Sync/fx-rate-sync/SGDUSD_Tracker.xlsx"
LOG="$HOME/FX Rate Sync - Do not delete/scripts/fx-pull.log"

{
  echo "=== Run at $(date) ==="
  cd "$MIRROR_DIR" || { echo "FAILED: mirror folder not found at $MIRROR_DIR"; exit 1; }

  if git pull origin main; then
    cp "$MIRROR_DIR/SGDUSD_Tracker.xlsx" "$ONEDRIVE_FILE"
    echo "OK: pulled latest and copied to OneDrive at $(date)"
  else
    echo "FAILED: git pull did not succeed (see error above) — OneDrive file was NOT touched, still showing the last good copy. Will retry automatically."
  fi
} >> "$LOG" 2>&1
