#!/bin/bash
# Logistics Invoice Updater — Diagnóstico
# Corre isto e envia o output completo.

echo ""
echo "============================================"
echo " Diagnóstico — $(date '+%Y-%m-%d %H:%M')"
echo "============================================"
echo ""

PROJECT_DIR="$(cd "$(dirname "$0")/.." && pwd)"

echo "--- macOS ---"
sw_vers
echo ""

echo "--- Xcode CLT ---"
if xcode-select -p &>/dev/null; then
    echo "Instalado: $(xcode-select -p)"
else
    echo "NÃO instalado"
fi
echo ""

echo "--- Python (Apple) ---"
if [ -x /usr/bin/python3 ]; then
    echo "Caminho: /usr/bin/python3"
    /usr/bin/python3 --version 2>&1
else
    echo "NÃO encontrado"
fi
echo ""

echo "--- Python (python.org) ---"
FOUND_PY=""
for v in 3.13 3.12 3.11 3.10; do
    P="/Library/Frameworks/Python.framework/Versions/$v/bin/python3"
    if [ -x "$P" ]; then
        echo "Encontrado: $P"
        "$P" --version 2>&1
        FOUND_PY="$P"
        break
    fi
done
[ -z "$FOUND_PY" ] && echo "NÃO encontrado"
echo ""

echo "--- tkinter ---"
if [ -n "$FOUND_PY" ]; then
    "$FOUND_PY" -c "import tkinter; print('tkinter OK')" 2>&1
elif [ -x /usr/bin/python3 ]; then
    /usr/bin/python3 -c "import tkinter; print('tkinter OK')" 2>&1
else
    echo "Sem Python para testar"
fi
echo ""

echo "--- Pasta do projecto ---"
echo "Localização: $PROJECT_DIR"
echo ""
echo "Ficheiros:"
ls -la "$PROJECT_DIR/" 2>&1
echo ""

echo "--- .venv ---"
if [ -d "$PROJECT_DIR/.venv" ]; then
    echo "Encontrado: $PROJECT_DIR/.venv"
    "$PROJECT_DIR/.venv/bin/python" --version 2>&1
    "$PROJECT_DIR/.venv/bin/python" -c "import pdfplumber; print('pdfplumber OK')" 2>&1
    "$PROJECT_DIR/.venv/bin/python" -c "import openpyxl; print('openpyxl OK')" 2>&1
    "$PROJECT_DIR/.venv/bin/python" -c "import tkinter; print('tkinter OK')" 2>&1
else
    echo "NÃO encontrado (corre install_mac.command primeiro)"
fi
echo ""

echo "--- Quarentena ---"
xattr -l "$PROJECT_DIR/MAC/install_mac.command" 2>&1
xattr -l "$PROJECT_DIR/MAC/Open Logistics Updater.command" 2>&1
echo ""

echo "--- Permissões ---"
ls -l "$PROJECT_DIR/MAC/"*.command 2>&1
echo ""

echo "============================================"
echo " Fim do diagnóstico"
echo " Copia tudo acima e envia."
echo "============================================"
read -p "Pressiona Enter para fechar..."
