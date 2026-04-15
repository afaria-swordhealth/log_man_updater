@echo off
title Logistics Invoice Updater — Instalação
echo.
echo ============================================
echo  Logistics Invoice Updater — Instalação
echo ============================================
echo.

REM Os .py estão um nível acima (pasta APP)
set PROJECT_DIR=%~dp0..

REM Verificar se Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: Python não encontrado.
    echo.
    echo Por favor instala o Python 3.10 ou superior:
    echo https://www.python.org/downloads/
    echo.
    echo Assegura que marcas "Add Python to PATH" durante a instalação.
    pause
    exit /b 1
)

echo Python encontrado. A instalar dependências...
echo.
pip install pdfplumber openpyxl

if %errorlevel% neq 0 (
    echo.
    echo ERRO: Falha na instalação das dependências.
    echo Tenta correr este ficheiro como Administrador.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Instalação concluída com sucesso!
echo  Podes agora correr "Run Logistics Updater.bat"
echo ============================================
echo.
pause
