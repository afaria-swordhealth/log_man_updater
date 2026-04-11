@echo off
title Logistics Invoice Updater

REM Caminho para o script Python (dois níveis acima desta pasta)
set SCRIPT_DIR=%~dp0..\..
set SCRIPT=%SCRIPT_DIR%\update_logistics_gui.py

python "%SCRIPT%"

if %errorlevel% neq 0 (
    echo.
    echo ERRO ao iniciar a aplicação.
    echo Verifica que o Python e as dependências estão instalados.
    echo Corre "install_windows.bat" se ainda não o fizeste.
    pause
)
