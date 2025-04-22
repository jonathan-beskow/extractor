@echo off
chcp 65001 >nul
setlocal

set SCRIPT_DIR=%~dp0extrator-files

if not exist "%SCRIPT_DIR%" (
    echo ❌ A pasta "extrator-files" não foi encontrada.
    pause
    exit /b 1
)

cd /d "%SCRIPT_DIR%"
echo =============================================
echo Diretório atual: %cd%
echo Executando extracao de horas por aplicacao...
echo =============================================

python extrator_local.py

IF %ERRORLEVEL% NEQ 0 (
    echo ❌ Ocorreu um erro durante a execução do script Python.
) ELSE (
    echo ✅ Execução concluída com sucesso.
)

pause
