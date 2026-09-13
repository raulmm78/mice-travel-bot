#!/bin/zsh
cd "$(dirname "$0")"

echo "Instalador MICE Travel Bot"
echo "=========================="
echo

if ! command -v python3 >/dev/null 2>&1; then
  echo "No se ha encontrado Python 3."
  echo "Instalalo desde https://www.python.org/downloads/macos/ o con Homebrew."
  exit 1
fi

python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r app/requirements.txt

mkdir -p config
if [ ! -f config/.env ]; then
  cp config/.env.example config/.env
  echo
  echo "Se ha creado config/.env. Editalo antes de arrancar."
else
  echo
  echo "Ya existe config/.env. No se ha tocado."
fi

echo
echo "Instalacion terminada."
echo "Puedes abrir MICE Travel Bot con abrir_panel_mac.command o con la app del Escritorio."
