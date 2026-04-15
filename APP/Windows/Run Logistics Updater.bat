@echo off
title Logistics Invoice Updater

REM Os .py estão um nível acima (pasta APP)
set SCRIPT=%~dp0..\update_logistics_gui.py

if not exist "%SCRIPT%" (
    echo ERRO: update_logistics_gui.py não encontrado.
    echo Esperado em: %SCRIPT%
    pause
    exit /b 1
)

python "%SCRIPT%"

if %errorlevel% neq 0 (
    echo.
    echo ERRO ao iniciar a aplicação.
    echo Verifica que o Python e as dependências estão instalados.
    echo Corre "install_windows.bat" se ainda não o fizeste.
    pause
)
