#!/bin/bash
# Logistics Invoice Updater — Instalação (Mac)
# Usa Python oficial do python.org (tkinter estável, sem Xcode).

set -e

echo ""
echo "============================================"
echo " Logistics Invoice Updater — Instalação"
echo "============================================"
echo ""

# Os .py estão um nível acima (pasta APP)
PROJECT_DIR="$(cd "$(dirname "$0")/.." && pwd)"

if [ ! -f "$PROJECT_DIR/update_logistics_gui.py" ]; then
    echo "ERRO: update_logistics_gui.py não encontrado em $PROJECT_DIR"
    read -p "Pressiona Enter para fechar..."
    exit 1
fi

echo "Projecto: $PROJECT_DIR"
echo ""

# Procurar Python oficial do python.org (tem tkinter a funcionar)
PY=""
for v in 3.13 3.12 3.11 3.10; do
    CAND="/Library/Frameworks/Python.framework/Versions/$v/bin/python3"
    if [ -x "$CAND" ]; then
        PY="$CAND"
        break
    fi
done

if [ -z "$PY" ]; then
    echo "Python oficial (python.org) não encontrado."
    echo "A descarregar instalador..."
    PKG="/tmp/python-installer.pkg"
    curl -L -o "$PKG" "https://www.python.org/ftp/python/3.12.7/python-3.12.7-macos11.pkg"
    echo ""
    echo "============================================"
    echo " A abrir o instalador do Python."
    echo " Clica 'Continue' até ao fim e 'Install'."
    echo " Vai pedir a tua password do Mac."
    echo " Depois volta a correr este instalador."
    echo "============================================"
    open "$PKG"
    read -p "Pressiona Enter para fechar..."
    exit 0
fi

echo "Python encontrado: $PY"
echo ""

cd "$PROJECT_DIR"

echo "A criar ambiente virtual..."
"$PY" -m venv .venv

echo "A actualizar pip..."
.venv/bin/python -m pip install --upgrade pip

echo "A instalar dependências..."
.venv/bin/python -m pip install pdfplumber openpyxl

echo ""
echo "============================================"
echo " Instalação concluída!"
echo " Usa 'Open Logistics Updater.command'"
echo "============================================"
read -p "Pressiona Enter para fechar..."
