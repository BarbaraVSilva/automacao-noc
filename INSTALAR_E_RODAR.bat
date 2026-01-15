@echo off
title Instalador Assistente NOC
echo --------------------------------------------------
echo   PREPARANDO AMBIENTE PARA AUTOMACAO NOC
echo --------------------------------------------------

:: Verifica se o Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado! 
    echo Por favor, instale o Python via 'Software Center' ou site oficial.
    pause
    exit
)

echo [1/3] Instalando bibliotecas necessarias...
pip install openpyxl pyperclip keyboard configparser --quiet

echo [2/3] Verificando arquivos...
if not exist automacao_noc.py (
    echo [ERRO] Arquivo 'automacao_noc.py' nao encontrado na pasta!
    pause
    exit
)

echo [3/3] Iniciando o Assistente...
echo --------------------------------------------------
python automacao_noc.py

pause