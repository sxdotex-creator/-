#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
if [ ! -d ".venv" ]; then
  echo ".venv not found. Please create the environment first with:"
  echo "  python3 -m venv .venv"
  exit 1
fi
source .venv/bin/activate
python -m excel2word.gui
