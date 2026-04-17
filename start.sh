#!/bin/bash
# ══════════════════════════════════════════════
#  Payless Automation Hub — Start
# ══════════════════════════════════════════════

set -e
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$DIR"

if ! command -v python3 &>/dev/null; then
  echo "ERROR: Python 3 not found."
  exit 1
fi

if [ ! -d ".venv" ]; then
  echo "Setting up virtual environment..."
  python3 -m venv .venv
fi

source .venv/bin/activate
echo "Checking dependencies..."
pip install -q -r requirements.txt

echo ""
echo "  ✓ Payless Automation Hub starting..."
echo "  ✓ Open: http://localhost:5050"
echo ""

sleep 1 && open "http://localhost:5050" &
python3 server.py
