@echo off
cd /d "%~dp0"

call ".venv\Scripts\activate.bat"

python -m app.main

echo.
echo ========================================
echo Processo finalizado.
pause