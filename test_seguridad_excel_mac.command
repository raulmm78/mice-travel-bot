#!/bin/zsh
cd "$(dirname "$0")"
PYTHON_BIN="python3"
if [ -x ".venv/bin/python" ]; then
  PYTHON_BIN=".venv/bin/python"
fi
"$PYTHON_BIN" app/safety_test.py
echo
echo "Puedes cerrar esta ventana."
