@echo off
chcp 65001 >nul
title Automacao Lancamento Financeiro Manual
color 0A

cd /d "%~dp0"

echo =======================================================
echo          AUTOMACAO LANCAMENTO FINANCEIRO MANUAL
echo     (c) 2025 Warreno Hendrick Costa Lima Guimaraes
echo                      Versao 1.4
echo =======================================================
echo.
echo Verificando arquivos necessarios...

REM Verifica se o Node.js portable existe
if not exist "node-portable\node.exe" (
    echo ERRO: Node.js portable nao encontrado!
    echo Execute primeiro: Instalar_Dependencias.bat
    pause
    exit /b 1
)

REM Verifica se a planilha existe
if not exist "projeto\dados_lancamento.xlsx" (
    echo ERRO: Planilha dados_lancamento.xlsx nao encontrada!
    echo Coloque sua planilha na pasta 'projeto'
    pause
    exit /b 1
)

REM Navega para a pasta do projeto
cd projeto

echo.
echo Iniciando automacao...
echo.

REM Executa a automacao usando Node.js portable
"..\node-portable\node.exe" script.js

echo.
echo Automacao finalizada!
pause