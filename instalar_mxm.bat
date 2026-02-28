@echo off
setlocal EnableExtensions

for %%I in ("%~dp0.") do set "SCRIPT_DIR=%%~fI"
set "REPO_DIR=%SCRIPT_DIR%"
set "LOG_FILE=%SCRIPT_DIR%\instalar_mxm.log"

echo ==============================================
echo Instalador MXM - inicio
echo Pasta: %SCRIPT_DIR%
echo Repo : %REPO_DIR%
echo Log: %LOG_FILE%
echo ==============================================

echo [INFO] Iniciando instalacao... > "%LOG_FILE%"
echo [INFO] Script dir: %SCRIPT_DIR% >> "%LOG_FILE%"
echo [INFO] Repo dir: %REPO_DIR% >> "%LOG_FILE%"

where py >nul 2>nul
if %ERRORLEVEL%==0 (
    set "PY_CMD=py"
) else (
    set "PY_CMD=python"
)

echo [INFO] Interpretador selecionado: %PY_CMD%
echo [INFO] Interpretador selecionado: %PY_CMD% >> "%LOG_FILE%"

echo [RUN] %PY_CMD% "%SCRIPT_DIR%\scripts\install_mxm_tool.py" --repo-dir "%REPO_DIR%" --run-gui %* >> "%LOG_FILE%"
%PY_CMD% "%SCRIPT_DIR%\scripts\install_mxm_tool.py" --repo-dir "%REPO_DIR%" --run-gui %* >> "%LOG_FILE%" 2>&1
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo [ERRO] Instalacao falhou com codigo %RC%.
    echo [ERRO] Consulte o log: "%LOG_FILE%"
    echo.
    type "%LOG_FILE%"
    echo.
    pause
    exit /b %RC%
)

echo.
echo [OK] Instalacao concluida com sucesso.
echo [OK] Log em: "%LOG_FILE%"
echo.
pause
exit /b 0
