#!/bin/bash
# Logistics Invoice Updater — Lançador (Mac)

# Directório deste script → dois níveis acima ficam os .py
SCRIPT_DIR="$(cd "$(dirname "$0")/../.." && pwd)"
SCRIPT="$SCRIPT_DIR/update_logistics_gui.py"

if [ ! -f "$SCRIPT" ]; then
    echo "ERRO: Não foi possível encontrar update_logistics_gui.py"
    echo "Esperado em: $SCRIPT"
    read -p "Pressiona Enter para fechar..."
    exit 1
fi

python3 "$SCRIPT"

if [ $? -ne 0 ]; then
    echo ""
    echo "ERRO ao iniciar a aplicação."
    echo "Verifica que Python 3 e dependências estão instalados."
    echo "Corre 'install_mac.command' se ainda não o fizeste."
    read -p "Pressiona Enter para fechar..."
fi
