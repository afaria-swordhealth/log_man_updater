#!/bin/bash
# Logistics Invoice Updater — Reset completo
# Remove tudo o que os scripts de instalação criaram.

echo ""
echo "============================================"
echo " Logistics Invoice Updater — Reset"
echo "============================================"
echo ""
echo "Isto vai remover:"
echo "  - Ambiente virtual (.venv) desta pasta"
echo "  - Pacotes pip instalados com --user"
echo "  - uv (se foi instalado)"
echo "  - Cache do pip"
echo ""
read -p "Tens a certeza? (s/n) " CONFIRM
if [ "$CONFIRM" != "s" ] && [ "$CONFIRM" != "S" ]; then
    echo "Cancelado."
    read -p "Pressiona Enter para fechar..."
    exit 0
fi

echo ""

# Pasta do projecto (um nível acima de MAC/)
PROJECT_DIR="$(cd "$(dirname "$0")/.." && pwd)"

# 1. Apagar .venv
if [ -d "$PROJECT_DIR/.venv" ]; then
    echo "A remover .venv..."
    rm -rf "$PROJECT_DIR/.venv"
else
    echo ".venv não encontrado (OK)"
fi

# 2. Apagar pacotes pip --user
if [ -d "$HOME/Library/Python" ]; then
    echo "A remover pacotes pip --user..."
    rm -rf "$HOME/Library/Python"
else
    echo "Pacotes pip --user não encontrados (OK)"
fi

# 3. Apagar uv
if [ -f "$HOME/.local/bin/uv" ] || [ -d "$HOME/.local/share/uv" ]; then
    echo "A remover uv..."
    rm -rf "$HOME/.local/bin/uv" "$HOME/.local/share/uv" "$HOME/.cache/uv"
else
    echo "uv não encontrado (OK)"
fi

# 4. Apagar cache pip
if [ -d "$HOME/Library/Caches/pip" ]; then
    echo "A limpar cache pip..."
    rm -rf "$HOME/Library/Caches/pip"
else
    echo "Cache pip não encontrado (OK)"
fi

echo ""
echo "============================================"
echo " Reset concluído!"
echo " Para reinstalar, corre 'install_mac.command'"
echo "============================================"
read -p "Pressiona Enter para fechar..."
