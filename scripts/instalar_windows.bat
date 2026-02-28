@echo off
setlocal EnableExtensions

for %%I in ("%~dp0.") do set "SCRIPT_DIR=%%~fI"
for %%I in ("%SCRIPT_DIR%\..") do set "REPO_DIR=%%~fI"
set "LOG_FILE=%SCRIPT_DIR%\instalar_windows.log"

echo [INFO] Instalador auxiliar Windows iniciado > "%LOG_FILE%"
echo [INFO] Script dir: %SCRIPT_DIR% >> "%LOG_FILE%"
echo [INFO] Repo dir: %REPO_DIR% >> "%LOG_FILE%"

where py >nul 2>nul
if %ERRORLEVEL%==0 (
    set "PY_CMD=py"
) else (
    set "PY_CMD=python"
)

echo [INFO] Python: %PY_CMD% >> "%LOG_FILE%"
echo [RUN] %PY_CMD% "%SCRIPT_DIR%\install_mxm_tool.py" --repo-dir "%REPO_DIR%" %* >> "%LOG_FILE%"
%PY_CMD% "%SCRIPT_DIR%\install_mxm_tool.py" --repo-dir "%REPO_DIR%" %* >> "%LOG_FILE%" 2>&1
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo [ERRO] Falha no instalador. Veja "%LOG_FILE%".
    type "%LOG_FILE%"
    pause
    exit /b %RC%
)

echo [OK] Instalador concluido. Log: "%LOG_FILE%"
pause
exit /b 0
