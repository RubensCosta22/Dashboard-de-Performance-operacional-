@echo off
chcp 65001 >nul
title Sistema de Descarga — Demo

echo.
echo  ================================================
echo   Sistema de Descarga Operacional — Demo
echo  ================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERRO] Python nao encontrado.
    pause & exit /b 1
)

echo  Verificando dependencias...
pip show flask >nul 2>&1 || pip install flask openpyxl --quiet
echo  OK.
echo.

:: Pasta do script
set SCRIPT_DIR=%~dp0

:: ── Pasta local do banco ────────────────────────────────────────────────────
set LOCAL_DATA=%SCRIPT_DIR%data
set LOCAL_EXCEL=%LOCAL_DATA%\relatorios.xlsx
set REDE_EXCEL=%SCRIPT_DIR%data\relatorios.xlsx

:: Criar pasta local se não existir
if not exist "%LOCAL_DATA%" (
    echo  Criando pasta local %LOCAL_DATA%...
    mkdir "%LOCAL_DATA%"
)

:: Se banco local não existe mas existe na rede, copiar para local
if not exist "%LOCAL_EXCEL%" (
    if exist "%REDE_EXCEL%" (
        echo  Copiando banco inicial para pasta de dados...
        copy /Y "%REDE_EXCEL%" "%LOCAL_EXCEL%" >nul
        echo  Banco copiado: %LOCAL_EXCEL%
    ) else (
        echo  Banco sera criado automaticamente em %LOCAL_EXCEL%
    )
)

echo  Banco de dados: %LOCAL_EXCEL%
echo.
echo  Iniciando servidor...
echo  Relatorio:  http://localhost:5050
echo  Dashboard:  http://localhost:5050/dashboard
echo.
echo  Ctrl+C para encerrar.
echo.

cd /d "%~dp0"
python server.py
pause
