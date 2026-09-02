#!/bin/zsh
for port in 8765 8766 8767 8768 8769; do
  for pid in $(lsof -ti tcp:$port 2>/dev/null); do
    kill "$pid" 2>/dev/null || true
  done
done
echo "MICE Travel Bot detenido. Puedes cerrar esta ventana."
