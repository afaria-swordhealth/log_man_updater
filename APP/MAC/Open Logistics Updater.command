#!/bin/bash
# Logistics Invoice Updater — Lançador (Mac)

# Os .py e .venv estão um nível acima (pasta APP)
PROJECT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
SCRIPT="$PROJECT_DIR/update_logistics_gui.py"
VENV_PY="$PROJECT_DIR/.venv/bin/python"

if [ ! -f "$SCRIPT" ]; then
    echo "ERRO: update_logistics_gui.py não encontrado em $PROJECT_DIR"
    read -p "Pressiona Enter para fechar..."
    exit 1
fi

if [ ! -x "$VENV_PY" ]; then
    echo "ERRO: Ambiente virtual (.venv) não encontrado."
    echo "Corre 'install_mac.command' primeiro."
    read -p "Pressiona Enter para fechar..."
    exit 1
fi

"$VENV_PY" "$SCRIPT"

if [ $? -ne 0 ]; then
    echo ""
    echo "ERRO ao correr a aplicação."
    read -p "Pressiona Enter para fechar..."
fi
