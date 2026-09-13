#!/bin/zsh
cd "$(dirname "$0")/.."
PYTHON_BIN="python3"
if [ -x ".venv/bin/python" ]; then
  PYTHON_BIN=".venv/bin/python"
fi
"$PYTHON_BIN" app/process_emails.py --dashboard
