@echo off
chcp 65001 >nul
title Setup Inicial - Automacao Lancamento Financeiro Manual

cd /d "%~dp0"

echo ============================================================
echo    SETUP INICIAL - AUTOMACAO LANCAMENTO FINANCEIRO MANUAL
echo        (c) 2025 Warreno Hendrick Costa Lima Guimaraes
echo                         Versao 1.4
echo ============================================================
echo.
echo Este processo ira:
echo 1. Baixar a versao mais recente do Node.js portable
echo 2. Instalar dependencias necessarias
echo 3. Configurar o ambiente
echo.
echo ATENCAO: Requer conexao com internet!
echo.
pause

REM Cria pasta para Node.js portable se nao existir
if not exist "node-portable" mkdir node-portable

echo.
echo [1/4] Verificando Node.js portable...

REM Verifica se ja tem Node.js
if exist "node-portable\node.exe" (
    echo Node.js portable ja esta instalado!
    goto install_deps
)

echo.
echo Node.js nao encontrado. Buscando versao mais recente...

REM Busca a versao mais recente do Node.js
echo Obtendo informacoes da versao mais recente...
for /f "delims=" %%i in ('powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (Invoke-WebRequest 'https://nodejs.org/dist/index.json' -UseBasicParsing | ConvertFrom-Json)[0].version"') do set NODE_VERSION=%%i

if "%NODE_VERSION%"=="" (
    echo.
    echo ERRO: Nao foi possivel obter a versao mais recente do Node.js!
    echo Usando versao padrao v22.21.0...
    set NODE_VERSION=v22.21.0
)

echo Versao encontrada: %NODE_VERSION%

REM URL do Node.js portable (versao mais recente)
set NODE_URL=https://nodejs.org/dist/%NODE_VERSION%/node-%NODE_VERSION%-win-x64.zip
set ZIP_FILE=node-portable\node.zip

echo Baixando Node.js %NODE_VERSION%...
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $webClient = New-Object System.Net.WebClient; $webClient.Headers.Add('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'); $webClient.DownloadFile('%NODE_URL%', '%ZIP_FILE%')"

if errorlevel 1 (
    echo.
    echo ERRO: Falha ao baixar Node.js %NODE_VERSION%!
    echo Tentando versao LTS v22.21.0...
    
    set NODE_VERSION=v22.21.0
    set NODE_URL=https://nodejs.org/dist/%NODE_VERSION%/node-%NODE_VERSION%-win-x64.zip
    
    powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $webClient = New-Object System.Net.WebClient; $webClient.Headers.Add('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'); $webClient.DownloadFile('%NODE_URL%', '%ZIP_FILE%')"
    
    if errorlevel 1 (
        echo.
        echo ERRO: Falha ao baixar Node.js!
        pause
        exit /b 1
    )
)

echo Extraindo Node.js...
tar -xf "%ZIP_FILE%" -C "node-portable"

if errorlevel 1 (
    echo.
    echo ERRO: Falha ao extrair Node.js com tar. Tentando metodo alternativo...
    powershell -Command "Expand-Archive -Path '%ZIP_FILE%' -DestinationPath 'node-portable' -Force"
    
    if errorlevel 1 (
        echo.
        echo ERRO: Falha ao extrair Node.js!
        pause
        exit /b 1
    )
)

REM Move conteudo da subpasta para a raiz da node-portable
for /f "delims=" %%i in ('dir /ad /b "node-portable\node-*" 2^>nul') do (
    xcopy "node-portable\%%i\*" "node-portable\" /E /H /C /I /Y >nul
    rmdir /s /q "node-portable\%%i"
)

REM Remove o arquivo zipado
del "%ZIP_FILE%"

echo Node.js %NODE_VERSION% portable instalado com sucesso!

:install_deps
echo.
echo [2/4] Navegando para pasta do projeto...
cd projeto

echo.
echo [3/4] Instalando dependencias...
call "..\node-portable\npm.cmd" install

if errorlevel 1 (
    echo.
    echo ERRO: Falha ao instalar dependencias!
    echo Verifique sua conexao com internet.
    pause
    exit /b 1
)

echo.

echo [4/4] Instalando navegadores do Playwright...

REM Configura PATH temporariamente para incluir Node.js portable
set "OLD_PATH=%PATH%"
set "PATH=%~dp0node-portable;%PATH%"

REM Primeiro tenta com npx
call "..\node-portable\npx.cmd" playwright install chromium

if errorlevel 1 (
    echo.
    echo ERRO com npx. Tentando metodo alternativo...
    
    REM Metodo alternativo: executar diretamente via node
    "..\..\node-portable\node.exe" "./node_modules/.bin/playwright" install chromium
    
    if errorlevel 1 (
        echo.
        echo ERRO: Falha ao instalar navegador chromium!
        echo.
        echo SOLUCOES POSSIVEIS:
        echo 1. Execute manualmente: npx playwright install chromium
        echo 2. Verifique sua conexao com internet
        echo 3. Execute como administrador
        echo.
        set "PATH=%OLD_PATH%"
        pause
        exit /b 1
    )
)

REM Restaura PATH original
set "PATH=%OLD_PATH%"

cd ..

echo.
echo ==========================================
echo         SETUP CONCLUIDO COM SUCESSO!
echo ==========================================
echo.
echo Agora voce pode:
echo 1. Colocar sua planilha 'dados_lancamento.xlsx' na pasta 'projeto'
echo 2. Executar 'EXECUTAR.bat' para iniciar a automacao
echo.
echo Arquivos importantes:
echo - EXECUTAR.bat       : Inicia a automacao
echo - projeto/           : Pasta com script e planilha
echo - node-portable/     : Node.js portavel (%NODE_VERSION%)
echo.
pause