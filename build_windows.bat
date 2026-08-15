@echo off
setlocal
cd /d "%~dp0"

rem Find the main script: the only .py that is not template_data.py
set "MAIN_PY="
for %%f in (*.py) do (
    if /I not "%%~nxf"=="template_data.py" set "MAIN_PY=%%f"
)

if not defined MAIN_PY (
    echo No main Python file found next to this script.
    pause
    exit /b 1
)

echo Installing build dependencies...
python -m pip install --upgrade pyinstaller openpyxl requests
if errorlevel 1 goto :error

echo Building TransferTool.exe...
python -m PyInstaller --noconfirm --clean --onefile --windowed --name TransferTool "%MAIN_PY%"
if errorlevel 1 goto :error

copy /Y dist\TransferTool.exe TransferTool.exe >nul
if errorlevel 1 goto :error

echo.
echo Build OK: TransferTool.exe
pause
exit /b 0

:error
echo.
echo Build failed.
pause
exit /b 1
