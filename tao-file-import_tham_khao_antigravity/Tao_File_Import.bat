@echo off
setlocal
chcp 65001 >nul

set "WORK_DIR=%~dp0"
if "%WORK_DIR:~-1%"=="\" if not "%WORK_DIR:~-2%"==":\" set "WORK_DIR=%WORK_DIR:~0,-1%"

set "SCRIPT_PATH=%~dp0scripts\tao_file_import.py"
if not exist "%SCRIPT_PATH%" set "SCRIPT_PATH=D:\AI_AntiGravity skill\tao-file-import\scripts\tao_file_import.py"

if not exist "%SCRIPT_PATH%" (
    echo [LOI] Khong tim thay script xu ly: %SCRIPT_PATH%
    echo.
    pause
    exit /b 1
)

python -X utf8 "%SCRIPT_PATH%" "%WORK_DIR%"
echo.
pause
