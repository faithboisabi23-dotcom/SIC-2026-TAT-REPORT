#!/usr/bin/env bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "============================================"
echo " TAT Dashboard - Data Refresh"
echo "============================================"
echo ""

# Check Python is available
if ! command -v python3 &>/dev/null; then
    echo "ERROR: python3 not found. Please install Python 3.9 or higher."
    exit 1
fi

# Install / upgrade dependencies
echo "Installing dependencies..."
pip3 install -r "$SCRIPT_DIR/requirements.txt" --quiet

echo ""
echo "Running data export..."
python3 "$SCRIPT_DIR/scripts/export_dashboard_json.py"

echo ""
echo "============================================"
echo " Done! dashboard/data/ has been updated."
echo " Commit and push to refresh GitHub Pages."
echo "============================================"