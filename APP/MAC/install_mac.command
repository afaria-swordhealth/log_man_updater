#!/bin/bash
# Logistics Invoice Updater — Instalação (Mac)

echo ""
echo "============================================"
echo " Logistics Invoice Updater — Instalação"
echo "============================================"
echo ""

# Verificar Python 3
if ! command -v python3 &>/dev/null; then
    echo "ERRO: Python 3 não encontrado."
    echo ""
    echo "Instala o Python 3 em: https://www.python.org/downloads/"
    read -p "Pressiona Enter para fechar..."
    exit 1
fi

echo "Python 3 encontrado. A instalar dependências..."
echo ""

pip3 install pdfplumber openpyxl

if [ $? -ne 0 ]; then
    echo ""
    echo "ERRO: Falha na instalação. A tentar com --user..."
    pip3 install --user pdfplumber openpyxl
fi

echo ""
echo "============================================"
echo " Instalação concluída!"
echo " Podes agora usar 'Open Logistics Updater.command'"
echo "============================================"
echo ""
read -p "Pressiona Enter para fechar..."
