@echo off
setlocal enabledelayedexpansion

echo --- Localizando Python no Sistema ---
set "PY_EXE="
set "PYW_EXE="

:: 1. BUSCA PRIORITARIA: Caminho local do usuario (AppData)
for %%v in (315 314 313 312 311 310) do (
    if not defined PY_EXE (
        set "BASE_DIR=%LOCALAPPDATA%\Programs\Python\Python%%v"
        if exist "!BASE_DIR!\python.exe" (
            set "PY_EXE=!BASE_DIR!\python.exe"
            set "PYW_EXE=!BASE_DIR!\pythonw.exe"
            echo Detectado Python %%v no perfil do usuario.
        )
    )
)

:: 2. BUSCA SECUNDARIA: Comandos globais (PATH)
if not defined PY_EXE (
    where python >nul 2>nul
    if %ERRORLEVEL% equ 0 (
        :: Se encontrou, pega o caminho completo para evitar erro 'pythonw' nao encontrado
        for /f "delims=" %%i in ('where python') do set "PY_EXE=%%i"
        
        where pythonw >nul 2>nul
        if %ERRORLEVEL% equ 0 (
            for /f "delims=" %%i in ('where pythonw') do set "PYW_EXE=%%i"
        ) else (
            set "PYW_EXE=!PY_EXE:python.exe=pythonw.exe!"
        )
        echo Detectado Python no PATH global.
    )
)

:: 3. BUSCA TERCIARIA: Program Files
if not defined PY_EXE (
    for %%v in (315 314 313 312 311 310) do (
        if not defined PY_EXE (
            set "BASE_DIR=C:\Program Files\Python%%v"
            if exist "!BASE_DIR!\python.exe" (
                set "PY_EXE=!BASE_DIR!\python.exe"
                set "PYW_EXE=!BASE_DIR!\pythonw.exe"
                echo Detectado Python %%v em Program Files.
            )
        )
    )
)

if not defined PY_EXE (
    echo.
    echo [ERRO] Nao foi possivel encontrar o Python instalado.
    echo Por favor, instale o Python em https://www.python.org/
    echo e marque a opcao "Add Python to PATH" durante a instalacao.
    echo.
    pause
    exit /b 1
)

:: Garantir que o pythonw existe de fato
if not exist "!PYW_EXE!" (
    echo [AVISO] 'pythonw.exe' nao encontrado. Usando 'python.exe' padrao.
    set "PYW_EXE=!PY_EXE!"
)

echo Usando Executavel: "!PY_EXE!"

:: Verifica se o servidor ja esta rodando na porta 5000
echo Verificando se servidor ja esta em execucao...
"!PY_EXE!" -c "import socket; s=socket.socket(); s.settimeout(1); r=s.connect_ex(('127.0.0.1',5000)); s.close(); exit(0 if r==0 else 1)" >nul 2>nul

if %ERRORLEVEL% equ 0 (
    echo Servidor ja esta em execucao. Abrindo navegador...
    start http://localhost:5000
    timeout /t 2 >nul
    exit
)

echo --- Iniciando Servidor (Modo Silencioso) ---

:: Instala dependencias
"!PY_EXE!" -m pip install -r requirements.txt --quiet

:: Inicia o servidor de forma invisivel usando pythonw
start "" /b "!PYW_EXE!" server.py

:: Aguarda ate 15 segundos para o servidor ficar pronto
echo Aguardando servidor iniciar...
for /l %%i in (1,1,15) do (
    timeout /t 1 /nobreak >nul
    "!PY_EXE!" -c "import socket; s=socket.socket(); s.settimeout(0.5); r=s.connect_ex(('127.0.0.1',5000)); s.close(); exit(0 if r==0 else 1)" >nul 2>nul
    if !ERRORLEVEL! equ 0 goto :server_ready
)

:server_ready
:: Abre o navegador
start http://localhost:5000

echo.
echo Servidor iniciado! Feche esta janela a qualquer momento.
timeout /t 3 >nul
exit
